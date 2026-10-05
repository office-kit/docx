/**
 * Thumbnails for the Table Styles gallery: how a small table looks in each
 * style with the current Table Style Options, resolved with the same
 * conditional-formatting rules the canvas uses.
 */

import {
  addTable,
  BUILT_IN_TABLE_STYLES,
  builtInTableStyle,
  createDocx,
  type Docx,
  getElementAttr,
  setTableLook,
  setTableStyle,
  stylesPart,
  type TableLook,
} from "@office-kit/docx";
import { createStyleResolver } from "./resolve.js";
import { type BorderSpec, resolveTable } from "./table-format.js";

export interface TablePreviewCell {
  /** CSS background color, if the cell is shaded. */
  readonly background?: string | undefined;
  /** CSS `border-*` values, `none` where there is no border. */
  readonly top: string;
  readonly right: string;
  readonly bottom: string;
  readonly left: string;
  /** CSS text color of the cell's text (the gallery draws it as a short bar). */
  readonly text: string;
  readonly bold: boolean;
}

const PREVIEW_ROWS = 5;
const PREVIEW_COLS = 5;
const NO_BORDER: ReadonlySet<string> = new Set(["none", "nil"]);
const HEX = /^[0-9A-Fa-f]{6}$/;

function borderCss(spec: BorderSpec | undefined): string {
  if (!spec || NO_BORDER.has(spec.style)) return "none";
  const color = HEX.test(spec.color) ? `#${spec.color}` : "#000";
  // Thumbnails are tiny: any line is one pixel, a double line three.
  return spec.style === "double" ? `3px double ${color}` : `1px solid ${color}`;
}

/**
 * Preview cells (`[row][col]`, 5 × 5) for each table style id, as the
 * document would show them. Built-in styles not yet in the document are
 * previewed from their definitions colored with the document's theme.
 */
export function tableStylePreviews(
  doc: Docx,
  styleIds: readonly string[],
  look: TableLook,
): Map<string, TablePreviewCell[][]> {
  const scratch = createDocx({ paragraphs: [] });
  const target = stylesPart(scratch);
  if (!target) throw new Error("A new document always has styles.");
  const present = new Map<string, number>();
  target.styles.forEach((s, i) => {
    const id = getElementAttr(s, "styleId");
    if (id !== undefined) present.set(id, i);
  });
  // The document's own table styles (custom ones, or modified built-ins) win.
  for (const style of stylesPart(doc)?.styles ?? []) {
    if (getElementAttr(style, "type") !== "table") continue;
    const id = getElementAttr(style, "styleId");
    if (id === undefined) continue;
    const index = present.get(id);
    if (index === undefined) {
      present.set(id, target.styles.length);
      target.styles.push(style);
    } else {
      target.styles[index] = style;
    }
  }
  const builtIn = new Set(BUILT_IN_TABLE_STYLES.map((s) => s.styleId));
  for (const id of styleIds) {
    if (!present.has(id) && builtIn.has(id)) {
      present.set(id, target.styles.length);
      target.styles.push(builtInTableStyle(doc, id));
    }
  }
  const table = addTable(
    scratch,
    Array.from({ length: PREVIEW_ROWS }, () => Array.from({ length: PREVIEW_COLS }, () => "x")),
  );
  setTableLook(table, look);
  const styles = createStyleResolver(scratch);
  const para = table.rows[0]?.cells[0]?.paragraphs[0];
  const out = new Map<string, TablePreviewCell[][]>();
  for (const id of styleIds) {
    setTableStyle(table, id);
    const format = resolveTable(scratch, table);
    out.set(
      id,
      format.cells.map((row) =>
        row.map((cell) => {
          const run = para ? styles.run(para, undefined, cell.text) : undefined;
          const color = run?.color && HEX.test(run.color) ? `#${run.color}` : "#000";
          return {
            background: cell.background,
            top: borderCss(cell.borders.top),
            right: borderCss(cell.borders.right),
            bottom: borderCss(cell.borders.bottom),
            left: borderCss(cell.borders.left),
            text: color,
            bold: run?.bold ?? false,
          };
        }),
      ),
    );
  }
  return out;
}
