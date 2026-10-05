/**
 * The colour, line and preset galleries of Shape Format and the Draw tab, in
 * Word's order: theme colour grid, standard colours, outline weights and
 * dashes, arrowheads, shadow presets, Shape Styles and WordArt Styles.
 */

import type {
  ShapeArrow,
  ShapeCategory,
  ShapeDash,
  ShapeFill,
  ShapePreset,
  ShapeShadow,
  ShapeStroke,
} from "@office-kit/docx";
import { SHAPE_PRESETS } from "@office-kit/docx";

// Word's "Standard Colors" row.
export const STANDARD_COLORS = [
  "C00000",
  "FF0000",
  "FFC000",
  "FFFF00",
  "92D050",
  "00B050",
  "00B0F0",
  "0070C0",
  "002060",
  "7030A0",
] as const;

/** The ten theme colours across the top of Word's palette, in its order. */
export const THEME_SLOTS = [
  "lt1",
  "dk1",
  "lt2",
  "dk2",
  "accent1",
  "accent2",
  "accent3",
  "accent4",
  "accent5",
  "accent6",
] as const;
const ACCENTS = ["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"] as const;

function channels(hex: string): [number, number, number] {
  const n = Number.parseInt(hex, 16);
  return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
}

function toHex([r, g, b]: readonly [number, number, number]): string {
  return [r, g, b]
    .map((c) =>
      Math.round(Math.max(0, Math.min(255, c)))
        .toString(16)
        .padStart(2, "0"),
    )
    .join("")
    .toUpperCase();
}

/** Mix toward white (`amount` > 0) or black (`amount` < 0). */
export function tint(hex: string, amount: number): string {
  const c = channels(hex);
  return toHex(
    amount >= 0
      ? [c[0] + (255 - c[0]) * amount, c[1] + (255 - c[1]) * amount, c[2] + (255 - c[2]) * amount]
      : [c[0] * (1 + amount), c[1] * (1 + amount), c[2] * (1 + amount)],
  );
}

/**
 * Word's theme colour grid: each theme colour over five variations (lighter
 * 80/60/40 %, darker 25/50 %; the light background colour gets darker
 * 5/15/25/35/50 % instead).
 */
export function themeGrid(theme: Readonly<Record<string, string>>): string[][] {
  const base = THEME_SLOTS.map((slot) => theme[slot] ?? "000000");
  const rows: string[][] = [base];
  const lighter = [0.8, 0.6, 0.4, -0.25, -0.5];
  const fromWhite = [-0.05, -0.15, -0.25, -0.35, -0.5];
  for (let r = 0; r < 5; r++) {
    rows.push(
      base.map((hex, i) => {
        if (i === 0) return tint(hex, fromWhite[r] ?? 0);
        const isDark = channels(hex).reduce((a, b) => a + b, 0) < 120;
        // Black's variations are lighter shades of grey.
        return tint(hex, isDark ? ([0.5, 0.35, 0.25, 0.15, 0.05][r] ?? 0) : (lighter[r] ?? 0));
      }),
    );
  }
  return rows;
}

export const OUTLINE_WEIGHTS = [0.25, 0.5, 0.75, 1, 1.5, 2.25, 3, 4.5, 6] as const;

export const DASHES = [
  "solid",
  "shortdot",
  "dot",
  "dash",
  "dashdot",
  "longdash",
  "longdashdot",
  "longdashdotdot",
] as const satisfies readonly ShapeDash[];

/** Arrows menu: (start, end) pairs in Word's order. */
export const ARROW_STYLES: ReadonlyArray<readonly [ShapeArrow, ShapeArrow]> = [
  ["none", "none"],
  ["none", "block"],
  ["block", "none"],
  ["block", "block"],
  ["none", "open"],
  ["open", "none"],
  ["open", "open"],
  ["oval", "oval"],
  ["diamond", "diamond"],
];

/** Shape Effects ▸ Shadow: Outer presets VML can express (offset shadows). */
export const SHADOW_PRESETS = [
    { id: "offsetBottomRight", shadow: { color: "000000", offsetX: 3, offsetY: 3, opacity: 0.4 } },
    { id: "offsetBottom", shadow: { color: "000000", offsetX: 0, offsetY: 3, opacity: 0.4 } },
    { id: "offsetBottomLeft", shadow: { color: "000000", offsetX: -3, offsetY: 3, opacity: 0.4 } },
    { id: "offsetRight", shadow: { color: "000000", offsetX: 3, offsetY: 0, opacity: 0.4 } },
    { id: "offsetCenter", shadow: { color: "000000", offsetX: 0, offsetY: 0, opacity: 0.4 } },
    { id: "offsetLeft", shadow: { color: "000000", offsetX: -3, offsetY: 0, opacity: 0.4 } },
    { id: "offsetTopRight", shadow: { color: "000000", offsetX: 3, offsetY: -3, opacity: 0.4 } },
    { id: "offsetTop", shadow: { color: "000000", offsetX: 0, offsetY: -3, opacity: 0.4 } },
    { id: "offsetTopLeft", shadow: { color: "000000", offsetX: -3, offsetY: -3, opacity: 0.4 } },
  ] as const satisfies ReadonlyArray<{ readonly id: string; readonly shadow: ShapeShadow }>;

/** Gradient ▸ Light / Dark Variations, as (from, to, angle) over the current colour. */
export function gradientVariations(color: string): ShapeFill[] {
  const light = tint(color, 0.6);
  const dark = tint(color, -0.4);
  return [
    { type: "gradient", color: light, color2: color, angle: 0 },
    { type: "gradient", color: light, color2: color, angle: 45 },
    { type: "gradient", color: light, color2: color, angle: 90 },
    { type: "gradient", color: light, color2: color, angle: 135 },
    { type: "gradient", color: light, color2: color, style: "radial" },
    { type: "gradient", color, color2: dark, angle: 0 },
    { type: "gradient", color, color2: dark, angle: 45 },
    { type: "gradient", color, color2: dark, angle: 90 },
    { type: "gradient", color, color2: dark, angle: 135 },
    { type: "gradient", color, color2: dark, style: "radial" },
  ];
}

export interface ShapeStylePreset {
  readonly fill: ShapeFill;
  readonly stroke: ShapeStroke | null;
  readonly shadow: ShapeShadow | null;
  /** Text colour on the preset (shown in the gallery's "Abc"). */
  readonly text: string;
}

/**
 * Shape Styles ▸ Theme Styles: six rows (Colored Outline, Colored Fill,
 * Light 1 Outline Colored Fill, Subtle Effect, Moderate Effect, Intense
 * Effect) over Dark 1 and the six accents, expressed with VML's fill, line and
 * shadow.
 */
export function shapeStyles(theme: Readonly<Record<string, string>>): ShapeStylePreset[][] {
  const columns = [theme.dk1 ?? "000000", ...ACCENTS.map((a) => theme[a] ?? "4472C4")];
  const white = theme.lt1 ?? "FFFFFF";
  const soft: ShapeShadow = { color: "000000", offsetX: 0, offsetY: 2, opacity: 0.3 };
  const strong: ShapeShadow = { color: "000000", offsetX: 0, offsetY: 3, opacity: 0.5 };
  return [
    columns.map((c) => ({
      fill: { type: "solid", color: white },
      stroke: { color: c, weight: 1 },
      shadow: null,
      text: c,
    })),
    columns.map((c) => ({
      fill: { type: "solid", color: c },
      stroke: { color: tint(c, -0.5), weight: 1 },
      shadow: null,
      text: white,
    })),
    columns.map((c) => ({
      fill: { type: "solid", color: c },
      stroke: { color: white, weight: 3 },
      shadow: soft,
      text: white,
    })),
    columns.map((c) => ({
      fill: { type: "solid", color: tint(c, 0.8) },
      stroke: { color: c, weight: 1 },
      shadow: null,
      text: "000000",
    })),
    columns.map((c) => ({
      fill: { type: "gradient", color: tint(c, 0.3), color2: c, angle: 0 },
      stroke: null,
      shadow: soft,
      text: white,
    })),
    columns.map((c) => ({
      fill: { type: "gradient", color: tint(c, 0.2), color2: tint(c, -0.25), angle: 0 },
      stroke: null,
      shadow: strong,
      text: white,
    })),
  ];
}

/** WordArt Styles: fill, outline and shadow over the theme's text and accent colours. */
export function wordArtStyles(theme: Readonly<Record<string, string>>): ShapeStylePreset[] {
  const dark = theme.dk1 ?? "000000";
  const white = theme.lt1 ?? "FFFFFF";
  const a1 = theme.accent1 ?? "4472C4";
  const a2 = theme.accent2 ?? "ED7D31";
  const a4 = theme.accent4 ?? "FFC000";
  const shadow: ShapeShadow = { color: "000000", offsetX: 2, offsetY: 2, opacity: 0.35 };
  const solid = (color: string): ShapeFill => ({ type: "solid", color });
  return [
    { fill: solid(dark), stroke: null, shadow: null, text: dark },
    { fill: solid(a1), stroke: null, shadow, text: a1 },
    { fill: solid(a2), stroke: null, shadow, text: a2 },
    { fill: solid(white), stroke: { color: a1, weight: 1 }, shadow, text: white },
    { fill: solid(a4), stroke: null, shadow, text: a4 },
    {
      fill: { type: "gradient", color: tint(a1, 0.4), color2: tint(a1, -0.3), angle: 0 },
      stroke: null,
      shadow: null,
      text: a1,
    },
    { fill: solid(tint(dark, 0.5)), stroke: null, shadow: null, text: tint(dark, 0.5) },
    { fill: solid(white), stroke: { color: dark, weight: 0.75 }, shadow: null, text: white },
    { fill: solid(a2), stroke: { color: white, weight: 0.75 }, shadow, text: a2 },
    {
      fill: { type: "gradient", color: tint(a4, 0.3), color2: a2, angle: 0 },
      stroke: null,
      shadow,
      text: a2,
    },
    { fill: solid(tint(a1, -0.5)), stroke: null, shadow, text: tint(a1, -0.5) },
    { fill: { type: "none" }, stroke: { color: a1, weight: 1 }, shadow: null, text: white },
  ];
}

/** The gallery's presets per category, in the order of the preset table. */
export function presetsByCategory(): ReadonlyMap<ShapeCategory, ShapePreset[]> {
  const map = new Map<ShapeCategory, ShapePreset[]>();
  for (const [name, def] of Object.entries(SHAPE_PRESETS)) {
    if (name === "textBox" || !isPreset(name)) continue;
    const list = map.get(def.category) ?? [];
    list.push(name);
    map.set(def.category, list);
  }
  return map;
}

function isPreset(name: string): name is ShapePreset {
  return Object.hasOwn(SHAPE_PRESETS, name);
}

/** CSS for a fill (gallery swatches). */
export function fillCss(fill: ShapeFill): string {
  switch (fill.type) {
    case "none":
      return "transparent";
    case "solid":
      return `#${fill.color}`;
    case "gradient":
      return fill.style === "radial"
        ? `radial-gradient(#${fill.color}, #${fill.color2})`
        : `linear-gradient(${180 + (fill.angle ?? 0)}deg, #${fill.color}, #${fill.color2})`;
    default:
      return "#ccc";
  }
}
