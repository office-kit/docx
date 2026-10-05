/**
 * Theme colours: reading the document theme's colour scheme, and the tint /
 * shade arithmetic WordprocessingML's `w:themeTint` / `w:themeShade` describe
 * (ECMA-376 Part 1 §17.3.2.6): the colour is converted to HSL and its
 * luminance moved toward white (tint) or black (shade) by the given
 * fraction of 255.
 */

import {
  childElementsOf,
  type Docx,
  getElementAttr,
  getRawPartRoot,
  type ThemeColor,
  type XmlElement,
  xmlPartNames,
} from "@office-kit/docx";

/** The twelve colour slots of a theme's `<a:clrScheme>`. */
export type ThemeSlot =
  | "dk1"
  | "lt1"
  | "dk2"
  | "lt2"
  | "accent1"
  | "accent2"
  | "accent3"
  | "accent4"
  | "accent5"
  | "accent6"
  | "hlink"
  | "folHlink";

export type ThemePalette = Readonly<Record<ThemeSlot, string>>;

/**
 * Word's "Office" theme (Microsoft 365, 2023 onward), used when the document
 * has no theme part: Word opens such a document with its default theme.
 */
export const DEFAULT_THEME_COLORS: ThemePalette = {
  dk1: "000000",
  lt1: "FFFFFF",
  dk2: "0E2841",
  lt2: "E8E8E8",
  accent1: "156082",
  accent2: "E97132",
  accent3: "196B24",
  accent4: "0F9ED5",
  accent5: "A02B93",
  accent6: "4EA72E",
  hlink: "467886",
  folHlink: "96607D",
};

// w:themeColor names → scheme slots, with Word's default colour mapping
// (text1 = dark 1, background1 = light 1, …; §17.15.1.20).
const SLOT_OF: Readonly<Record<ThemeColor, ThemeSlot | undefined>> = {
  dark1: "dk1",
  text1: "dk1",
  light1: "lt1",
  background1: "lt1",
  dark2: "dk2",
  text2: "dk2",
  light2: "lt2",
  background2: "lt2",
  accent1: "accent1",
  accent2: "accent2",
  accent3: "accent3",
  accent4: "accent4",
  accent5: "accent5",
  accent6: "accent6",
  hyperlink: "hlink",
  followedHyperlink: "folHlink",
  none: undefined,
};

const THEME_PART_DIR = "/word/theme/";
const HEX6 = /^[0-9A-Fa-f]{6}$/;

function child(el: XmlElement | undefined, local: string): XmlElement | undefined {
  return el && childElementsOf(el).find((c) => c.name.local === local);
}

/** The root of the document's theme part, if it has one. */
export function themeRoot(doc: Docx): XmlElement | undefined {
  // The document part has exactly one theme relationship (§14.2.7); Word
  // always stores it under /word/theme/.
  const name = xmlPartNames(doc).find((n) => n.startsWith(THEME_PART_DIR));
  return name ? getRawPartRoot(doc, name) : undefined;
}

/** The document theme's colour scheme, slot by slot, falling back to Word's default theme. */
export function themePalette(doc: Docx): ThemePalette {
  const scheme = child(child(themeRoot(doc), "themeElements"), "clrScheme");
  const out: Record<ThemeSlot, string> = { ...DEFAULT_THEME_COLORS };
  for (const slot of Object.keys(DEFAULT_THEME_COLORS) as ThemeSlot[]) {
    const def = child(scheme, slot);
    const color = def && childElementsOf(def)[0];
    // System colours (dk1 / lt1 are usually `<a:sysClr lastClr>`) carry their
    // last computed RGB value.
    const value = color && (getElementAttr(color, "lastClr") ?? getElementAttr(color, "val"));
    if (value && HEX6.test(value)) out[slot] = value.toUpperCase();
  }
  return out;
}

/** The scheme slot a `w:themeColor` value names, if any. */
export function themeSlot(color: ThemeColor): ThemeSlot | undefined {
  return SLOT_OF[color];
}

function toHsl(hex: string): [number, number, number] {
  const n = Number.parseInt(hex, 16);
  const r = ((n >> 16) & 0xff) / 255;
  const g = ((n >> 8) & 0xff) / 255;
  const b = (n & 0xff) / 255;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) return [0, 0, l];
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  const h =
    max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
  return [h / 6, s, l];
}

function hueToChannel(p: number, q: number, t: number): number {
  const x = t < 0 ? t + 1 : t > 1 ? t - 1 : t;
  if (x < 1 / 6) return p + (q - p) * 6 * x;
  if (x < 1 / 2) return q;
  if (x < 2 / 3) return p + (q - p) * (2 / 3 - x) * 6;
  return p;
}

function fromHsl(h: number, s: number, l: number): string {
  let r = l;
  let g = l;
  let b = l;
  if (s !== 0) {
    const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
    const p = 2 * l - q;
    r = hueToChannel(p, q, h + 1 / 3);
    g = hueToChannel(p, q, h);
    b = hueToChannel(p, q, h - 1 / 3);
  }
  return [r, g, b]
    .map((c) =>
      Math.round(c * 255)
        .toString(16)
        .padStart(2, "0"),
    )
    .join("")
    .toUpperCase();
}

const UCHAR = 255;

/** `hex` lightened by a theme tint (0–255; 255 leaves it unchanged). */
export function applyTint(hex: string, tint: number): string {
  const [h, s, l] = toHsl(hex);
  const f = tint / UCHAR;
  return fromHsl(h, s, l * f + (1 - f));
}

/** `hex` darkened by a theme shade (0–255; 255 leaves it unchanged). */
export function applyShade(hex: string, shade: number): string {
  const [h, s, l] = toHsl(hex);
  return fromHsl(h, s, l * (shade / UCHAR));
}

/** Relative luminance (HSL lightness) of a colour, 0–1. */
export function lightness(hex: string): number {
  return toHsl(hex)[2];
}

/** One cell of Word's theme colour grid. */
export interface ThemeSwatch {
  readonly rgb: string;
  readonly themeColor: ThemeColor;
  readonly themeTint?: number;
  readonly themeShade?: number;
  /** Percentage for the tooltip: positive = lighter, negative = darker, 0 = the base colour. */
  readonly percent: number;
}

// The ten columns of Word's Theme Colors grid, left to right.
export const THEME_GRID_COLUMNS: readonly ThemeColor[] = [
  "background1",
  "text1",
  "background2",
  "text2",
  "accent1",
  "accent2",
  "accent3",
  "accent4",
  "accent5",
  "accent6",
];

// Word picks the five variations from the base colour's lightness: very dark
// colours get only lighter variants, very light ones only darker, the rest
// three lighter and two darker. Percentages as Word's tooltips show them.
const PURE_BLACK = [50, 35, 25, 15, 5];
const DARK = [90, 75, 50, 25, 10];
const PURE_WHITE = [-5, -15, -25, -35, -50];
const LIGHT = [-10, -25, -50, -75, -90];
const MID = [80, 60, 40, -25, -50];
const DARK_LIMIT = 0.2;
const LIGHT_LIMIT = 0.8;

function variations(base: string): readonly number[] {
  const l = lightness(base);
  if (l === 0) return PURE_BLACK;
  if (l === 1) return PURE_WHITE;
  if (l < DARK_LIMIT) return DARK;
  if (l > LIGHT_LIMIT) return LIGHT;
  return MID;
}

/**
 * Word's Theme Colors grid: one column per {@link THEME_GRID_COLUMNS} entry,
 * the base colour on top and five lighter / darker variations below.
 */
export function themeColorGrid(palette: ThemePalette): ThemeSwatch[][] {
  return THEME_GRID_COLUMNS.map((themeColor) => {
    const slot = SLOT_OF[themeColor] ?? "dk1";
    const base = palette[slot];
    const swatches: ThemeSwatch[] = [{ rgb: base, themeColor, percent: 0 }];
    for (const percent of variations(base)) {
      const amount = Math.round(UCHAR * (1 - Math.abs(percent) / 100));
      swatches.push(
        percent > 0
          ? { rgb: applyTint(base, amount), themeColor, themeTint: amount, percent }
          : { rgb: applyShade(base, amount), themeColor, themeShade: amount, percent },
      );
    }
    return swatches;
  });
}

/**
 * The RGB a colour reference shows: the theme slot with its tint / shade when
 * there is a usable theme reference, otherwise the stored value.
 */
export function resolveThemeColor(
  palette: ThemePalette,
  stored: string | undefined,
  themeColor: string | undefined,
  tint: string | undefined,
  shade: string | undefined,
): string | undefined {
  const slot = themeColor ? SLOT_OF[themeColor as ThemeColor] : undefined;
  if (!slot) return stored;
  let rgb = palette[slot];
  const t = tint === undefined ? Number.NaN : Number.parseInt(tint, 16);
  const s = shade === undefined ? Number.NaN : Number.parseInt(shade, 16);
  if (Number.isInteger(t)) rgb = applyTint(rgb, t);
  if (Number.isInteger(s)) rgb = applyShade(rgb, s);
  return rgb;
}
