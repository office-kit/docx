/**
 * State shared by the Table Design / Table Layout tabs, the Insert ▸ Table
 * button and the canvas: the pen (Line Style / Line Weight / Pen Color), the
 * pointer tool the canvas is in (Draw Table, Eraser, Border Painter), and
 * View Gridlines. Module state rather than session fields, so the canvas only
 * needs one import and the session stays free of table specifics.
 */

import type { TableBorder, TableBorderEdges } from "@office-kit/docx";
import { locale } from "../../i18n/index.svelte";

export type TablePointerTool = "draw" | "eraser" | "painter";

// Word's default pen: a ½ pt automatic single line.
const DEFAULT_PEN: TableBorder = { style: "single", size: 4, color: "auto" };

export const tableTool = $state<{
  mode: TablePointerTool | null;
  pen: TableBorder;
  /** Word shows table gridlines by default. */
  gridlines: boolean;
  /** The Shading button's face color (its last choice). */
  shading: string;
  /** The Borders button's face (its last choice). */
  borders: TableBorderEdges;
}>({ mode: null, pen: DEFAULT_PEN, gridlines: true, shading: "auto", borders: "bottom" });

/** Toggle a pointer tool (pressing the active tool's button turns it off, as in Word). */
export function toggleTableTool(mode: TablePointerTool): void {
  tableTool.mode = tableTool.mode === mode ? null : mode;
}

// --- Measurement units ----------------------------------------------------------

export type LengthUnit = "in" | "cm" | "mm";

const TWIPS_PER_INCH = 1440;
const CM_PER_INCH = 2.54;
const TWIPS_PER: Record<LengthUnit, number> = {
  in: TWIPS_PER_INCH,
  cm: TWIPS_PER_INCH / CM_PER_INCH,
  mm: TWIPS_PER_INCH / CM_PER_INCH / 10,
};
const DECIMALS: Record<LengthUnit, number> = { in: 2, cm: 2, mm: 1 };

/** The unit Word's ribbon spinners use: inches for US English, millimetres in Japan, centimetres elsewhere. */
export function lengthUnit(): LengthUnit {
  const id = locale();
  return id === "en" ? "in" : id === "ja" ? "mm" : "cm";
}

export function twipsToUnit(twips: number, unit: LengthUnit): number {
  const factor = 10 ** DECIMALS[unit];
  return Math.round((twips / TWIPS_PER[unit]) * factor) / factor;
}

export function unitToTwips(value: number, unit: LengthUnit): number {
  return Math.round(value * TWIPS_PER[unit]);
}

/**
 * The spinner step: the displayed precision (0.01″, 0.01 cm, 0.1 mm). A
 * coarser step would make the browser reject a converted value such as
 * 0.25″ as a step mismatch and block the dialog's form.
 */
export function unitStep(unit: LengthUnit): number {
  return 10 ** -DECIMALS[unit];
}

// --- Colors -----------------------------------------------------------------------

/** Word's Standard Colors row. */
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

// The Office theme's ten palette columns (lt1, dk1, lt2, dk2, accent1–6).
const THEME_BASE = [
  "FFFFFF",
  "000000",
  "E7E6E6",
  "44546A",
  "4472C4",
  "ED7D31",
  "A5A5A5",
  "FFC000",
  "5B9BD5",
  "70AD47",
] as const;

// The variations under each theme color, as Word's palette lists them:
// positive = lighter by that fraction, negative = darker.
const WHITE_STEPS = [-0.05, -0.15, -0.25, -0.35, -0.5];
const BLACK_STEPS = [0.5, 0.35, 0.25, 0.15, 0.05];
const LIGHT_STEPS = [-0.1, -0.25, -0.5, -0.75, -0.9];
const DEFAULT_STEPS = [0.8, 0.6, 0.4, -0.25, -0.5];
// Colors at least this light take the darkening variations.
const LIGHT_LUMINANCE = 0.8;

function hsl(hex: string): [number, number, number] {
  const channel = (i: number): number => Number.parseInt(hex.slice(i, i + 2), 16) / 255;
  const r = channel(0);
  const g = channel(2);
  const b = channel(4);
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

function hex([h, s, l]: [number, number, number]): string {
  const channel = (n: number): number => {
    const k = (n + h * 12) % 12;
    const a = s * Math.min(l, 1 - l);
    return l - a * Math.max(-1, Math.min(k - 3, 9 - k, 1));
  };
  return [0, 8, 4]
    .map((n) =>
      Math.round(channel(n) * 255)
        .toString(16)
        .padStart(2, "0"),
    )
    .join("")
    .toUpperCase();
}

function vary(base: string, step: number): string {
  const [h, s, l] = hsl(base);
  return hex([h, s, step > 0 ? l + (1 - l) * step : l * (1 + step)]);
}

/** The theme palette: one column per theme color, the base on top and its five variations below. */
export function themePalette(): string[][] {
  return THEME_BASE.map((base) => {
    const l = hsl(base)[2];
    const steps =
      base === "FFFFFF"
        ? WHITE_STEPS
        : base === "000000"
          ? BLACK_STEPS
          : l >= LIGHT_LUMINANCE
            ? LIGHT_STEPS
            : DEFAULT_STEPS;
    const column: string[] = [base];
    for (const step of steps) column.push(vary(base, step));
    return column;
  });
}

/** CSS for a border pen (`auto` draws black). */
export function penCss(pen: TableBorder): string {
  const color = pen.color === "auto" ? "#000" : `#${pen.color}`;
  const width = Math.max(1, pen.size / 8);
  const style =
    pen.style === "double" || pen.style === "triple" || pen.style.startsWith("thick")
      ? "double"
      : pen.style === "dotted"
        ? "dotted"
        : pen.style.startsWith("dash") || pen.style.startsWith("dot")
          ? "dashed"
          : "solid";
  return `${pen.style === "double" ? Math.max(3, width) : width}pt ${style} ${color}`;
}
