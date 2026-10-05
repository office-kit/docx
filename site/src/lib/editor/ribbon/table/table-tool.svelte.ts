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
