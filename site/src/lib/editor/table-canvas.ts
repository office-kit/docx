/**
 * Table behavior on the canvas, kept out of EditorCanvas.svelte so the canvas
 * only forwards events: Tab / Shift+Tab between cells, the cell-selection
 * highlight, View Gridlines, and the pointer tools (Draw Table, Eraser,
 * Border Painter).
 */

import {
  adjacentCellPosition,
  caretAt,
  cellStart,
  type Command,
  commands,
  type EditorModel,
  tableSelection,
} from "@office-kit/docx-editor";
import { tableColumnCount, type TableBorderEdges, type TableCellRange } from "@office-kit/docx";
import { tableTool } from "./ribbon/table/table-tool.svelte";

/** Runs a command the way the canvas does (reporting failures); returns whether it applied. */
export type CanvasExec = <P>(cmd: Command<P>, params: P) => boolean;

const SELECTED = "wk-cell-sel";
const MULTI = "wk-multicell";
// CSS px are 1/96″, twips 1/1440″.
const TWIPS_PER_PX = 15;
// A drag shorter than this (px) is a click, not a drawn line.
const MIN_STROKE = 6;
// Word's table gridlines: a light dashed line where a cell has no border.
const GRIDLINE = "1px dashed #a5c3e6";

/**
 * Tab / Shift+Tab in a table move to the next / previous cell, and Tab in
 * the last cell adds a row, as in Word; Escape drops a pointer tool. Returns
 * whether the key was handled (the caller re-renders and restores the caret).
 */
export function handleTableKey(e: KeyboardEvent, model: EditorModel, exec: CanvasExec): boolean {
  if (e.key === "Escape" && tableTool.mode) {
    tableTool.mode = null;
    return true;
  }
  if (e.key !== "Tab" || e.ctrlKey || e.metaKey || e.altKey) return false;
  const ts = tableSelection(model);
  if (!ts) return false;
  e.preventDefault();
  const next = adjacentCellPosition(model, e.shiftKey ? -1 : 1);
  if (next) {
    model.setSelection(caretAt(next));
    return true;
  }
  if (e.shiftKey) return true;
  // Past the last cell: a new row, with the caret in its first cell.
  if (exec(commands.insertRowsCommand, { where: "below" })) {
    model.setSelection(caretAt(cellStart(ts.block, ts.table.rows.length - 1, 0)));
  }
  return true;
}

/** Highlight the selected cells when the selection spans several (Word shades them). */
export function decorateTableSelection(canvas: HTMLElement, model: EditorModel): void {
  for (const td of canvas.querySelectorAll(`.${SELECTED}`)) td.classList.remove(SELECTED);
  const ts = tableSelection(model);
  const multi = !!ts?.multiCell;
  canvas.classList.toggle(MULTI, multi);
  if (!ts || !multi) return;
  const { range } = ts;
  for (const row of ts.placements) {
    for (const p of row) {
      const inside =
        p.rowSpan > 0 &&
        p.row >= range.firstRow &&
        p.row <= range.lastRow &&
        p.gridStart >= range.firstColumn &&
        p.gridStart + p.gridSpan - 1 <= range.lastColumn;
      if (!inside) continue;
      canvas
        .querySelector(`td[data-wk-block="${ts.block}"][data-wk-cell="${p.row},${p.cell}"]`)
        ?.classList.add(SELECTED);
    }
  }
}

/** Inline style for the page: View Gridlines draws borderless cell edges. */
export function tableCanvasStyle(): string {
  return tableTool.gridlines ? `--wk-gridline:${GRIDLINE}` : "";
}

/** The `data-table-tool` attribute (pointer cursor for the active tool). */
export function tableToolAttr(): string | undefined {
  return tableTool.mode ?? undefined;
}

/** The average laid-out height (twips) of rows `first`…`last` of the table at `block`. */
export function measuredRowHeight(block: number, first: number, last: number): number | undefined {
  const table = document.querySelector<HTMLTableElement>(
    `table.wk-table[data-wk-block="${block}"]`,
  );
  const rows = table ? Array.from(table.tBodies[0]?.rows ?? []).slice(first, last + 1) : [];
  if (rows.length === 0) return undefined;
  const total = rows.reduce((sum, tr) => sum + tr.offsetHeight, 0);
  return Math.round((total / rows.length) * TWIPS_PER_PX);
}

// --- Pointer tools ------------------------------------------------------------------------

interface CellHit {
  readonly td: HTMLTableCellElement;
  readonly block: number;
  readonly row: number;
  readonly cell: number;
  readonly range: TableCellRange;
}

function cellHit(target: EventTarget | null): CellHit | undefined {
  if (!(target instanceof Element)) return undefined;
  const td = target.closest<HTMLTableCellElement>("td[data-wk-cell]");
  // Cells of nested tables carry no anchors and are not editable here.
  if (!td) return undefined;
  const [row = Number.NaN, cell = Number.NaN] = (td.dataset.wkCell ?? "").split(",").map(Number);
  const block = Number(td.dataset.wkBlock);
  const grid = Number(td.dataset.wkGrid);
  const span = Number(td.dataset.wkSpan);
  if (![row, cell, block, grid, span].every(Number.isInteger)) return undefined;
  return {
    td,
    block,
    row,
    cell,
    range: {
      firstRow: row,
      lastRow: row + td.rowSpan - 1,
      firstColumn: grid,
      lastColumn: grid + span - 1,
    },
  };
}

type Side = "top" | "bottom" | "left" | "right";

function nearestSide(td: Element, x: number, y: number): Side {
  const r = td.getBoundingClientRect();
  const distances: [Side, number][] = [
    ["top", y - r.top],
    ["bottom", r.bottom - y],
    ["left", x - r.left],
    ["right", r.right - x],
  ];
  return distances.reduce((best, d) => (d[1] < best[1] ? d : best))[0];
}

let stroke: { x: number; y: number; hit: CellHit | undefined } | null = null;

/**
 * Pointer down on the canvas with a table tool active. Returns whether the
 * tool took the event (the caller then skips its own handling).
 */
export function tablePointerDown(e: PointerEvent, model: EditorModel, exec: CanvasExec): boolean {
  const mode = tableTool.mode;
  if (!mode) return false;
  const hit = cellHit(e.target);
  if (mode === "draw") {
    stroke = { x: e.clientX, y: e.clientY, hit };
    // Outside a table the browser still places the caret where the new table goes.
    if (hit) e.preventDefault();
    return true;
  }
  if (!hit) return false;
  e.preventDefault();
  model.setSelection(caretAt(cellStart(hit.block, hit.row, hit.cell)));
  const side = nearestSide(hit.td, e.clientX, e.clientY);
  const edges: TableBorderEdges = side;
  if (mode === "painter") {
    exec(commands.rangeBordersCommand, { edges, border: tableTool.pen, range: hit.range });
    return true;
  }
  // Eraser: an edge between two cells merges them; an outer edge loses its border.
  const ts = tableSelection(model);
  const lastColumn = ts ? tableColumnCount(ts.table) - 1 : 0;
  const lastRow = (ts?.table.rows.length ?? 1) - 1;
  const r = hit.range;
  const merged: TableCellRange | undefined =
    side === "top" && r.firstRow > 0
      ? { ...r, firstRow: r.firstRow - 1 }
      : side === "bottom" && r.lastRow < lastRow
        ? { ...r, lastRow: r.lastRow + 1 }
        : side === "left" && r.firstColumn > 0
          ? { ...r, firstColumn: r.firstColumn - 1 }
          : side === "right" && r.lastColumn < lastColumn
            ? { ...r, lastColumn: r.lastColumn + 1 }
            : undefined;
  if (merged) exec(commands.mergeCellsCommand, { range: merged });
  else exec(commands.rangeBordersCommand, { edges, range: r });
  return true;
}

/** Pointer up after a Draw Table stroke: split the cell along the line, or draw a new table. */
export function tablePointerUp(e: PointerEvent, model: EditorModel, exec: CanvasExec): void {
  const start = stroke;
  stroke = null;
  if (!start || tableTool.mode !== "draw") return;
  const dx = Math.abs(e.clientX - start.x);
  const dy = Math.abs(e.clientY - start.y);
  if (Math.max(dx, dy) < MIN_STROKE) return;
  const hit = start.hit;
  if (hit) {
    model.setSelection(caretAt(cellStart(hit.block, hit.row, hit.cell)));
    // A mostly horizontal line splits the cell into rows, a vertical one into columns.
    exec(commands.splitCellsCommand, dx > dy ? { columns: 1, rows: 2 } : { columns: 2, rows: 1 });
    return;
  }
  // A rectangle drawn in the text: a one-cell table as wide as the rectangle.
  exec(commands.insertTableCommand, {
    rows: 1,
    cols: 1,
    columnWidthTwips: Math.round(dx * TWIPS_PER_PX),
  });
}
