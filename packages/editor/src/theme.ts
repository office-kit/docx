/**
 * The document's theme as the editor uses it: colours for theme-colour
 * references (`w:themeColor`) and the heading / body fonts. A document
 * without a theme part behaves as Word's: theme references fall back to the
 * Office theme.
 */

import {
  type Docx,
  getTheme,
  THEMES,
  themeColorValue,
  type ThemeColorScheme,
  type ThemeInfo,
} from "@office-kit/docx";

function officeTheme(): ThemeInfo {
  const office = THEMES[0];
  if (!office) throw new Error("The built-in theme list is empty.");
  return {
    name: office.name,
    colors: office.colors,
    fonts: office.fonts,
    effects: office.effects.name,
  };
}

/** The document's theme, or Word's Office theme when it has none. */
export function resolveTheme(doc: Docx): ThemeInfo {
  return getTheme(doc) ?? officeTheme();
}

/** The theme colour slots Word's colour pickers show, in palette column order. */
export const THEME_PALETTE_COLUMNS = [
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
] as const;

export type PaletteThemeColor = (typeof THEME_PALETTE_COLUMNS)[number];

/** One cell of Word's theme colour grid: RGB plus how to write it back. */
export interface PaletteColor {
  readonly rgb: string;
  readonly themeColor: PaletteThemeColor;
  readonly themeShade?: string;
  readonly themeTint?: string;
}

// Word's five lighter / darker rows under each theme colour. Very dark or
// very light colours get different steps, as in Word's palette.
const ROWS_NORMAL: ReadonlyArray<{ tint?: string; shade?: string }> = [
  { tint: "33" },
  { tint: "66" },
  { tint: "99" },
  { shade: "BF" },
  { shade: "80" },
];
const ROWS_DARK: ReadonlyArray<{ tint?: string; shade?: string }> = [
  { tint: "80" },
  { tint: "A6" },
  { tint: "BF" },
  { tint: "D9" },
  { tint: "F2" },
];
const ROWS_LIGHT: ReadonlyArray<{ tint?: string; shade?: string }> = [
  { shade: "F2" },
  { shade: "D9" },
  { shade: "BF" },
  { shade: "A6" },
  { shade: "80" },
];

function luminance(hex: string): number {
  const n = Number.parseInt(hex, 16);
  return (0.299 * ((n >> 16) & 0xff) + 0.587 * ((n >> 8) & 0xff) + 0.114 * (n & 0xff)) / 255;
}

// Word's thresholds for switching the shade/tint rows (by perceived brightness).
const DARK_THRESHOLD = 0.2;
const LIGHT_THRESHOLD = 0.8;

/**
 * Word's theme colour grid: a row of the ten theme colours, then five rows
 * of lighter / darker variants of each.
 */
export function themeSchemePalette(colors: ThemeColorScheme): PaletteColor[][] {
  const base = THEME_PALETTE_COLUMNS.map((themeColor) => ({
    themeColor,
    rgb: themeColorValue(colors, themeColor) ?? "000000",
  }));
  const rows: PaletteColor[][] = [base];
  for (let r = 0; r < ROWS_NORMAL.length; r++) {
    rows.push(
      base.map(({ themeColor, rgb }) => {
        const lum = luminance(rgb);
        const step =
          (lum < DARK_THRESHOLD ? ROWS_DARK : lum > LIGHT_THRESHOLD ? ROWS_LIGHT : ROWS_NORMAL)[
            r
          ] ?? {};
        const cell: { -readonly [K in keyof PaletteColor]: PaletteColor[K] } = {
          rgb: themeColorValue(colors, themeColor, step) ?? rgb,
          themeColor,
        };
        if (step.shade) cell.themeShade = step.shade;
        if (step.tint) cell.themeTint = step.tint;
        return cell;
      }),
    );
  }
  return rows;
}
