/**
 * Word's built-in table styles (Table Grid, Plain Tables, Grid Tables, List
 * Tables) as `<w:style w:type="table">` definitions with conditional
 * formatting (§17.7.6). Word adds a built-in style to styles.xml the first
 * time a table uses it; these generators produce that definition.
 *
 * The recipes follow the Word 2013+ definitions: each family is one layout of
 * borders, shading and emphasis, drawn in a theme color (`text1` for the plain
 * variant, `accent1`…`accent6` for the accent variants) at fixed tints. Colors
 * are written both as hex and as theme references (`w:themeColor`,
 * `w:themeTint` / `w:themeShade`), as Word does, so a theme change recolors
 * the table. Hex values come from the document's theme so the two agree.
 */

import type { XmlElement } from "../xml/index.js";
import { wEl } from "./table-xml.js";

/** The colors of a theme's color scheme the table styles use (hex RGB). */
export interface ThemeColors {
  readonly dk1: string;
  readonly lt1: string;
  readonly accent1: string;
  readonly accent2: string;
  readonly accent3: string;
  readonly accent4: string;
  readonly accent5: string;
  readonly accent6: string;
}

/** The Office theme (2013–2022), what Word uses for a document without a theme part. */
export const OFFICE_THEME_COLORS: ThemeColors = {
  dk1: "000000",
  lt1: "FFFFFF",
  accent1: "4472C4",
  accent2: "ED7D31",
  accent3: "A5A5A5",
  accent4: "FFC000",
  accent5: "5B9BD5",
  accent6: "70AD47",
};

export type TableStyleCategory = "plain" | "grid" | "list";

export interface BuiltInTableStyle {
  readonly styleId: string;
  /** The style name Word shows (`w:name`). */
  readonly name: string;
  readonly category: TableStyleCategory;
}

/** Conditional formatting regions (ST_TblStyleOverrideType, §17.18.89). */
export type TableStyleRegion =
  | "wholeTable"
  | "firstRow"
  | "lastRow"
  | "firstCol"
  | "lastCol"
  | "band1Vert"
  | "band2Vert"
  | "band1Horz"
  | "band2Horz"
  | "neCell"
  | "nwCell"
  | "seCell"
  | "swCell";

interface ColorRef {
  readonly hex: string;
  readonly theme: string;
  readonly tint?: string;
  readonly shade?: string;
}

type Edge = { readonly val: string; readonly sz: number; readonly color: ColorRef } | "nil";
type Side = "top" | "left" | "bottom" | "right" | "insideH" | "insideV";

interface Region {
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly caps?: boolean;
  readonly color?: ColorRef;
  /** Half-points. */
  readonly size?: number;
  readonly jc?: string;
  readonly fill?: ColorRef;
  readonly borders?: Partial<Record<Side, Edge>>;
}

interface Recipe {
  readonly borders?: Partial<Record<Side, Edge>>;
  readonly fill?: ColorRef;
  readonly color?: ColorRef;
  readonly regions: Partial<Record<Exclude<TableStyleRegion, "wholeTable">, Region>>;
}

const ACCENTS = [1, 2, 3, 4, 5, 6] as const;
const TINT_20 = "33";
const TINT_40 = "66";
const TINT_60 = "99";
const SHADE_75 = "BF";
// Word draws Table Grid Light and Plain Table 1 in background1 shaded 25%
// (BFBFBF) and Plain Tables 2/3/5 in text1 lightened 50% (7F7F7F).
const SHADE_GREY = "BF";
const TINT_GREY = "80";
const BAND_GREY = "F2";

const HEX = /^[0-9A-Fa-f]{6}$/;

function channels(hex: string): number[] {
  const value = HEX.test(hex) ? hex : "000000";
  return [0, 2, 4].map((i) => Number.parseInt(value.slice(i, i + 2), 16) / 255);
}

// Word truncates the scaled channels; a hair of slack keeps exact values
// (e.g. 0.6 × 255) from flooring one step low.
const FLOOR_SLACK = 1e-6;

function toHex(rgb: readonly number[]): string {
  return rgb
    .map((c) =>
      Math.floor(Math.min(1, Math.max(0, c)) * 255 + FLOOR_SLACK)
        .toString(16)
        .padStart(2, "0"),
    )
    .join("")
    .toUpperCase();
}

/**
 * Scale a color's HSL luminance, keeping hue and saturation — how
 * DrawingML's lumMod/lumOff (and so Word's theme tints and shades) work.
 */
function withLuminance(hex: string, scale: (l: number) => number): string {
  const [r = 0, g = 0, b = 0] = channels(hex);
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  const d = max - min;
  const s = d === 0 ? 0 : d / (1 - Math.abs(2 * l - 1));
  let h = 0;
  if (d !== 0) {
    if (max === r) h = ((g - b) / d) % 6;
    else if (max === g) h = (b - r) / d + 2;
    else h = (r - g) / d + 4;
  }
  const l2 = Math.min(1, Math.max(0, scale(l)));
  const c = (1 - Math.abs(2 * l2 - 1)) * s;
  const x = c * (1 - Math.abs((((h % 6) + 6) % 6 % 2) - 1));
  const m = l2 - c / 2;
  const sector = Math.floor((((h % 6) + 6) % 6));
  const [r1, g1, b1] = [
    [c, x, 0],
    [x, c, 0],
    [0, c, x],
    [0, x, c],
    [x, 0, c],
    [c, 0, x],
  ][sector] ?? [0, 0, 0];
  return toHex([(r1 ?? 0) + m, (g1 ?? 0) + m, (b1 ?? 0) + m]);
}

/** `w:themeTint` (§17.3.2.6): luminance moved toward white, `tint`/255 of it kept. */
export function tintColor(hex: string, tint: string): string {
  const keep = Number.parseInt(tint, 16) / 255;
  return withLuminance(hex, (l) => l * keep + (1 - keep));
}

/** `w:themeShade`: luminance scaled toward black, `shade`/255 of it kept. */
export function shadeColor(hex: string, shade: string): string {
  const keep = Number.parseInt(shade, 16) / 255;
  return withLuminance(hex, (l) => l * keep);
}

function ref(hex: string, theme: string, mod?: { tint?: string; shade?: string }): ColorRef {
  if (mod?.tint) return { hex: tintColor(hex, mod.tint), theme, tint: mod.tint };
  if (mod?.shade) return { hex: shadeColor(hex, mod.shade), theme, shade: mod.shade };
  return { hex, theme };
}

const single = (sz: number, color: ColorRef): Edge => ({ val: "single", sz, color });
const double = (sz: number, color: ColorRef): Edge => ({ val: "double", sz, color });
const box = (edge: Edge, sides: readonly Side[]): Partial<Record<Side, Edge>> =>
  Object.fromEntries(sides.map((s) => [s, edge]));
const ALL_SIDES: readonly Side[] = ["top", "left", "bottom", "right", "insideH", "insideV"];
const OUTER: readonly Side[] = ["top", "left", "bottom", "right"];

/** The palette one style variant is drawn in. */
interface Palette {
  readonly full: ColorRef;
  readonly t60: ColorRef;
  readonly t40: ColorRef;
  readonly t20: ColorRef;
  readonly dark: ColorRef;
  readonly white: ColorRef;
}

function palette(colors: ThemeColors, accent: number | undefined): Palette {
  const theme = accent === undefined ? "text1" : `accent${accent}`;
  const base =
    accent === undefined ? colors.dk1 : colors[`accent${accent}` as keyof ThemeColors];
  return {
    full: ref(base, theme),
    t60: ref(base, theme, { tint: TINT_60 }),
    t40: ref(base, theme, { tint: TINT_40 }),
    t20: ref(base, theme, { tint: TINT_20 }),
    // The plain variant's "colorful" text stays text1 (black), not a shade.
    dark: accent === undefined ? ref(base, theme) : ref(base, theme, { shade: SHADE_75 }),
    white: ref(colors.lt1, "background1"),
  };
}

const BOLD: Region = { bold: true };

function bands(fill: ColorRef): Recipe["regions"] {
  return { band1Vert: { fill }, band1Horz: { fill } };
}

const GRID_RECIPES: Readonly<Record<number, (p: Palette) => Recipe>> = {
  1: (p) => ({
    borders: box(single(4, p.t40), ALL_SIDES),
    regions: {
      firstRow: { bold: true, borders: { bottom: single(12, p.t60) } },
      lastRow: { bold: true, borders: { top: double(2, p.t60) } },
      firstCol: BOLD,
      lastCol: BOLD,
    },
  }),
  2: (p) => ({
    borders: {
      top: single(2, p.t60),
      bottom: single(2, p.t60),
      insideH: single(2, p.t60),
      insideV: single(2, p.t60),
    },
    regions: {
      firstRow: {
        bold: true,
        fill: p.white,
        borders: { top: "nil", bottom: single(12, p.t60), insideH: "nil", insideV: "nil" },
      },
      lastRow: {
        bold: true,
        fill: p.white,
        borders: { top: double(2, p.t60), bottom: "nil", insideH: "nil", insideV: "nil" },
      },
      firstCol: BOLD,
      lastCol: BOLD,
      ...bands(p.t20),
    },
  }),
  3: (p) => ({
    borders: box(single(4, p.t60), ALL_SIDES),
    regions: {
      firstRow: { bold: true, borders: { top: "nil", left: "nil", right: "nil", insideV: "nil" } },
      lastRow: {
        bold: true,
        borders: { bottom: "nil", left: "nil", right: "nil", insideV: "nil" },
      },
      firstCol: {
        italic: true,
        jc: "right",
        fill: p.white,
        borders: { top: "nil", left: "nil", bottom: "nil", insideH: "nil" },
      },
      lastCol: {
        italic: true,
        jc: "left",
        fill: p.white,
        borders: { top: "nil", right: "nil", bottom: "nil", insideH: "nil" },
      },
      ...bands(p.t20),
      nwCell: { borders: { bottom: single(4, p.t60) } },
      neCell: { borders: { bottom: single(4, p.t60) } },
      swCell: { borders: { top: single(4, p.t60) } },
      seCell: { borders: { top: single(4, p.t60) } },
    },
  }),
  4: (p) => ({
    borders: box(single(4, p.t60), ALL_SIDES),
    regions: {
      firstRow: {
        bold: true,
        color: p.white,
        fill: p.full,
        borders: { ...box(single(4, p.full), OUTER), insideH: "nil", insideV: "nil" },
      },
      lastRow: { bold: true, borders: { top: double(4, p.full) } },
      firstCol: BOLD,
      lastCol: BOLD,
      ...bands(p.t20),
    },
  }),
  5: (p) => ({
    borders: box(single(4, p.white), ALL_SIDES),
    fill: p.t20,
    regions: {
      firstRow: {
        bold: true,
        color: p.white,
        fill: p.full,
        borders: { top: "nil", left: "nil", right: "nil", insideV: "nil" },
      },
      lastRow: {
        bold: true,
        color: p.white,
        fill: p.full,
        borders: { left: "nil", bottom: "nil", right: "nil", insideV: "nil" },
      },
      firstCol: {
        bold: true,
        color: p.white,
        fill: p.full,
        borders: { top: "nil", left: "nil", bottom: "nil", insideH: "nil" },
      },
      lastCol: {
        bold: true,
        color: p.white,
        fill: p.full,
        borders: { top: "nil", bottom: "nil", right: "nil", insideH: "nil" },
      },
      ...bands(p.t40),
    },
  }),
  6: (p) => ({
    borders: box(single(4, p.t60), ALL_SIDES),
    color: p.dark,
    regions: {
      firstRow: { bold: true, borders: { bottom: single(12, p.t60) } },
      lastRow: { bold: true, borders: { top: double(4, p.t60) } },
      firstCol: BOLD,
      lastCol: BOLD,
      ...bands(p.t20),
    },
  }),
  7: (p) => {
    const g3 = GRID_RECIPES[3]!(p);
    return { ...g3, color: p.dark };
  },
};

const LIST_RECIPES: Readonly<Record<number, (p: Palette) => Recipe>> = {
  1: (p) => ({
    regions: {
      firstRow: { bold: true, borders: { bottom: single(4, p.t60) } },
      lastRow: { bold: true, borders: { top: single(4, p.t60) } },
      firstCol: BOLD,
      lastCol: BOLD,
      ...bands(p.t20),
    },
  }),
  2: (p) => ({
    borders: { top: single(4, p.t60), bottom: single(4, p.t60), insideH: single(4, p.t60) },
    regions: { firstRow: BOLD, lastRow: BOLD, firstCol: BOLD, lastCol: BOLD, ...bands(p.t20) },
  }),
  3: (p) => ({
    borders: box(single(4, p.full), OUTER),
    regions: {
      firstRow: { bold: true, color: p.white, fill: p.full },
      lastRow: { bold: true, borders: { top: double(4, p.full) } },
      firstCol: BOLD,
      lastCol: BOLD,
      band1Vert: { borders: { left: single(4, p.full), right: single(4, p.full) } },
      band1Horz: { borders: { top: single(4, p.full), bottom: single(4, p.full) } },
    },
  }),
  4: (p) => ({
    borders: {
      ...box(single(4, p.t60), OUTER),
      insideH: single(4, p.t60),
    },
    regions: {
      firstRow: {
        bold: true,
        color: p.white,
        fill: p.full,
        borders: { ...box(single(4, p.full), OUTER), insideH: "nil" },
      },
      lastRow: { bold: true, borders: { top: double(4, p.full) } },
      firstCol: BOLD,
      lastCol: BOLD,
      ...bands(p.t20),
    },
  }),
  5: (p) => ({
    borders: box(single(24, p.full), OUTER),
    fill: p.full,
    color: p.white,
    regions: {
      firstRow: { bold: true, borders: { bottom: single(18, p.white) } },
      lastRow: { bold: true, borders: { top: single(4, p.white) } },
      firstCol: { bold: true, borders: { right: single(4, p.white) } },
      lastCol: { bold: true, borders: { left: single(4, p.white) } },
      band1Vert: { borders: { left: single(4, p.white), right: single(4, p.white) } },
      band1Horz: { borders: { top: single(4, p.white), bottom: single(4, p.white) } },
    },
  }),
  6: (p) => ({
    borders: { top: single(4, p.full), bottom: single(4, p.full) },
    color: p.dark,
    regions: {
      firstRow: { bold: true, borders: { bottom: single(4, p.full) } },
      lastRow: { bold: true, borders: { top: double(4, p.full) } },
      firstCol: BOLD,
      lastCol: BOLD,
      ...bands(p.t20),
    },
  }),
  7: (p) => {
    // Word's List Table 7 sets its headers in 13 pt italics, like Plain Table 5.
    const header: Region = { italic: true, size: 26 };
    return {
      color: p.dark,
      regions: {
        firstRow: { ...header, fill: p.white, borders: { bottom: single(4, p.full) } },
        lastRow: { ...header, fill: p.white, borders: { top: single(4, p.full) } },
        firstCol: { ...header, jc: "right", fill: p.white, borders: { right: single(4, p.full) } },
        lastCol: { ...header, fill: p.white, borders: { left: single(4, p.full) } },
        ...bands(p.t20),
        nwCell: { borders: { bottom: "nil" } },
      },
    };
  },
};

function plainRecipes(colors: ThemeColors): Readonly<Record<number, Recipe>> {
  const grey = ref(colors.lt1, "background1", { shade: SHADE_GREY });
  const mid = ref(colors.dk1, "text1", { tint: TINT_GREY });
  const band = ref(colors.lt1, "background1", { shade: BAND_GREY });
  const white = ref(colors.lt1, "background1");
  const header: Region = { italic: true, size: 26 };
  return {
    1: {
      borders: box(single(4, grey), ALL_SIDES),
      regions: {
        firstRow: BOLD,
        lastRow: { bold: true, borders: { top: double(4, grey) } },
        firstCol: BOLD,
        lastCol: BOLD,
        ...bands(band),
      },
    },
    2: {
      borders: { top: single(4, mid), bottom: single(4, mid) },
      regions: {
        firstRow: { bold: true, borders: { bottom: single(4, mid) } },
        lastRow: { bold: true, borders: { top: single(4, mid) } },
        firstCol: BOLD,
        lastCol: BOLD,
        band1Vert: { borders: { left: single(4, mid), right: single(4, mid) } },
        band1Horz: { borders: { top: single(4, mid), bottom: single(4, mid) } },
      },
    },
    3: {
      regions: {
        firstRow: { bold: true, caps: true, borders: { bottom: single(4, mid) } },
        lastRow: { bold: true, caps: true },
        firstCol: { bold: true, caps: true, borders: { right: single(4, mid) } },
        lastCol: { bold: true, caps: true },
        ...bands(band),
        nwCell: { borders: { right: "nil" } },
      },
    },
    4: {
      regions: { firstRow: BOLD, lastRow: BOLD, firstCol: BOLD, lastCol: BOLD, ...bands(band) },
    },
    5: {
      regions: {
        firstRow: { ...header, fill: white, borders: { bottom: single(4, mid) } },
        lastRow: { ...header, fill: white, borders: { top: single(4, mid) } },
        firstCol: { ...header, jc: "right", fill: white, borders: { right: single(4, mid) } },
        lastCol: { ...header, fill: white, borders: { left: single(4, mid) } },
        ...bands(band),
        nwCell: { borders: { bottom: "nil" } },
      },
    },
  };
}

const FAMILY_NAMES: Readonly<Record<number, string>> = {
  1: "1 Light",
  2: "2",
  3: "3",
  4: "4",
  5: "5 Dark",
  6: "6 Colorful",
  7: "7 Colorful",
};
const FAMILY_IDS: Readonly<Record<number, string>> = {
  1: "1Light",
  2: "2",
  3: "3",
  4: "4",
  5: "5Dark",
  6: "6Colorful",
  7: "7Colorful",
};
const FAMILIES = [1, 2, 3, 4, 5, 6, 7] as const;
const PLAIN_COUNT = 5;

// Word's uiPriority values, which order the styles in its gallery.
const PRIORITY_TABLE_GRID = 39;
const PRIORITY_GRID_LIGHT = 40;
const PRIORITY_PLAIN_1 = 41;
const PRIORITY_FAMILY_1 = 46;

type Spec = BuiltInTableStyle & {
  readonly priority: number;
  readonly recipe: (colors: ThemeColors) => Recipe;
};

function buildSpecs(): Spec[] {
  const specs: Spec[] = [
    {
      styleId: "TableGrid",
      name: "Table Grid",
      category: "plain",
      priority: PRIORITY_TABLE_GRID,
      recipe: () => ({ borders: box(single(4, { hex: "auto", theme: "" }), ALL_SIDES), regions: {} }),
    },
    {
      styleId: "TableGridLight",
      name: "Grid Table Light",
      category: "plain",
      priority: PRIORITY_GRID_LIGHT,
      recipe: (c) => ({
        borders: box(single(4, ref(c.lt1, "background1", { shade: SHADE_GREY })), ALL_SIDES),
        regions: {},
      }),
    },
  ];
  for (let n = 1; n <= PLAIN_COUNT; n++) {
    specs.push({
      styleId: `PlainTable${n}`,
      name: `Plain Table ${n}`,
      category: "plain",
      priority: PRIORITY_PLAIN_1 + n - 1,
      recipe: (c) => plainRecipes(c)[n]!,
    });
  }
  for (const [category, label, recipes] of [
    ["grid", "Grid", GRID_RECIPES],
    ["list", "List", LIST_RECIPES],
  ] as const) {
    for (const family of FAMILIES) {
      const make = recipes[family]!;
      const priority = PRIORITY_FAMILY_1 + family - 1;
      const baseId = `${label}Table${FAMILY_IDS[family]}`;
      const baseName = `${label} Table ${FAMILY_NAMES[family]}`;
      specs.push({
        styleId: baseId,
        name: baseName,
        category,
        priority,
        recipe: (c) => make(palette(c, undefined)),
      });
      for (const accent of ACCENTS) {
        specs.push({
          styleId: `${baseId}-Accent${accent}`,
          name: `${baseName} Accent ${accent}`,
          category,
          priority,
          recipe: (c) => make(palette(c, accent)),
        });
      }
    }
  }
  return specs;
}

const SPECS = buildSpecs();
const SPEC_BY_ID = new Map(SPECS.map((s) => [s.styleId, s]));

/** Every built-in table style, in Word's gallery order. */
export const BUILT_IN_TABLE_STYLE_LIST: readonly BuiltInTableStyle[] = SPECS.map(
  ({ styleId, name, category }) => ({ styleId, name, category }),
);

export function isBuiltInTableStyle(styleId: string): boolean {
  return SPEC_BY_ID.has(styleId);
}

function colorAttrs(
  color: ColorRef,
  names: { val: string; theme: string; tint: string; shade: string },
): Record<string, string | undefined> {
  return {
    [names.val]: color.hex,
    [names.theme]: color.theme || undefined,
    [names.tint]: color.tint,
    [names.shade]: color.shade,
  };
}

function borderEl(side: Side, edge: Edge): XmlElement {
  if (edge === "nil") return wEl(side, { val: "nil" });
  return wEl(side, {
    val: edge.val,
    sz: String(edge.sz),
    space: "0",
    ...colorAttrs(edge.color, {
      val: "color",
      theme: "themeColor",
      tint: "themeTint",
      shade: "themeShade",
    }),
  });
}

const SIDE_ORDER: readonly Side[] = ["top", "left", "bottom", "right", "insideH", "insideV"];

function bordersEl(local: string, borders: Partial<Record<Side, Edge>>): XmlElement | undefined {
  const children = SIDE_ORDER.flatMap((s) => {
    const edge = borders[s];
    return edge ? [borderEl(s, edge)] : [];
  });
  return children.length ? wEl(local, {}, children) : undefined;
}

function shdEl(fill: ColorRef): XmlElement {
  return wEl("shd", {
    val: "clear",
    color: "auto",
    ...colorAttrs(fill, {
      val: "fill",
      theme: "themeFill",
      tint: "themeFillTint",
      shade: "themeFillShade",
    }),
  });
}

function rPrEl(region: Region): XmlElement | undefined {
  const children: XmlElement[] = [];
  if (region.bold) children.push(wEl("b"), wEl("bCs"));
  if (region.italic) children.push(wEl("i"), wEl("iCs"));
  if (region.caps) children.push(wEl("caps"));
  if (region.color) {
    children.push(
      wEl(
        "color",
        colorAttrs(region.color, {
          val: "val",
          theme: "themeColor",
          tint: "themeTint",
          shade: "themeShade",
        }),
      ),
    );
  }
  if (region.size !== undefined) {
    children.push(wEl("sz", { val: String(region.size) }), wEl("szCs", { val: String(region.size) }));
  }
  return children.length ? wEl("rPr", {}, children) : undefined;
}

function regionEl(type: string, region: Region): XmlElement {
  const children: XmlElement[] = [];
  if (region.jc) children.push(wEl("pPr", {}, [wEl("jc", { val: region.jc })]));
  const rPr = rPrEl(region);
  if (rPr) children.push(rPr);
  const tcChildren: XmlElement[] = [];
  const borders = region.borders && bordersEl("tcBorders", region.borders);
  if (borders) tcChildren.push(borders);
  if (region.fill) tcChildren.push(shdEl(region.fill));
  if (tcChildren.length) children.push(wEl("tcPr", {}, tcChildren));
  return wEl("tblStylePr", { type }, children);
}

// The cell margins of Word's Normal Table, which every built-in style keeps.
const DEFAULT_CELL_MARGIN_LR = "108";

/** The `<w:style w:type="table" w:default="1">` Word names "Normal Table". */
export function normalTableStyle(): XmlElement {
  return wEl("style", { type: "table", default: "1", styleId: "TableNormal" }, [
    wEl("name", { val: "Normal Table" }),
    wEl("uiPriority", { val: "99" }),
    wEl("semiHidden"),
    wEl("unhideWhenUsed"),
    wEl("tblPr", {}, [
      wEl("tblInd", { w: "0", type: "dxa" }),
      wEl("tblCellMar", {}, [
        wEl("top", { w: "0", type: "dxa" }),
        wEl("left", { w: DEFAULT_CELL_MARGIN_LR, type: "dxa" }),
        wEl("bottom", { w: "0", type: "dxa" }),
        wEl("right", { w: DEFAULT_CELL_MARGIN_LR, type: "dxa" }),
      ]),
    ]),
  ]);
}

/** Build a built-in table style's definition, or undefined for an unknown id. */
export function builtInTableStyleElement(
  styleId: string,
  colors: ThemeColors,
  basedOn: string,
): XmlElement | undefined {
  const spec = SPEC_BY_ID.get(styleId);
  if (!spec) return undefined;
  const recipe = spec.recipe(colors);
  const children: XmlElement[] = [
    wEl("name", { val: spec.name }),
    wEl("basedOn", { val: basedOn }),
    wEl("uiPriority", { val: String(spec.priority) }),
    // Word sets single spacing with no space after inside every modern table style.
    wEl("pPr", {}, [wEl("spacing", { after: "0", line: "240", lineRule: "auto" })]),
  ];
  const rPr = recipe.color && rPrEl({ color: recipe.color });
  if (rPr) children.push(rPr);
  const tblPr: XmlElement[] = [];
  if (Object.keys(recipe.regions).length > 0) {
    tblPr.push(wEl("tblStyleRowBandSize", { val: "1" }), wEl("tblStyleColBandSize", { val: "1" }));
  }
  const borders = recipe.borders && bordersEl("tblBorders", recipe.borders);
  if (borders) tblPr.push(borders);
  if (tblPr.length) children.push(wEl("tblPr", {}, tblPr));
  if (recipe.fill) children.push(wEl("tcPr", {}, [shdEl(recipe.fill)]));
  for (const [type, region] of Object.entries(recipe.regions)) {
    if (region) children.push(regionEl(type, region));
  }
  return wEl("style", { type: "table", styleId }, children);
}
