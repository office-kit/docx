/**
 * The selection seen as table cells, as Word's table commands see it: a caret
 * in a cell is that cell; a selection from one cell to another of the same
 * table is the rectangle between them on the table grid, grown until no
 * merged cell is cut in two.
 */

import {
  type TableCellPlacement,
  type TableCellRange,
  tableCellPlacements,
  tableColumnCount,
  type WmlTable,
} from "@office-kit/docx";
import { asTable, blockAt } from "./doc-access.js";
import type { EditorModel } from "./model.js";
import { caretAt, type DocPosition, type Selection } from "./selection.js";

export interface TableSelection {
  /** The table's index in the body. */
  readonly block: number;
  readonly table: WmlTable;
  /** The selected rectangle (rows, and grid columns). */
  readonly range: TableCellRange;
  /** The cell holding the selection's focus. */
  readonly focus: { readonly row: number; readonly cell: number };
  readonly placements: readonly (readonly TableCellPlacement[])[];
  /** More than one cell selected. */
  readonly multiCell: boolean;
}

/** The rows a placement covers on screen (a continuation reports its merge's rows). */
function rowExtent(
  placements: readonly (readonly TableCellPlacement[])[],
  p: TableCellPlacement,
): [number, number] {
  if (p.rowSpan > 0) return [p.row, p.row + p.rowSpan - 1];
  let top = p.row;
  while (top > 0) {
    const above = placements[top - 1]?.find((q) => q.gridStart === p.gridStart);
    if (!above) break;
    top--;
    if (above.rowSpan > 0) return [top, top + above.rowSpan - 1];
  }
  return [p.row, p.row];
}

/** Grow a rectangle until it cuts no merged cell. */
function expand(
  placements: readonly (readonly TableCellPlacement[])[],
  start: TableCellRange,
): TableCellRange {
  let range = start;
  for (let changed = true; changed; ) {
    changed = false;
    for (let r = range.firstRow; r <= range.lastRow; r++) {
      for (const p of placements[r] ?? []) {
        const right = p.gridStart + p.gridSpan - 1;
        if (right < range.firstColumn || p.gridStart > range.lastColumn) continue;
        const [top, bottom] = rowExtent(placements, p);
        const next = {
          firstRow: Math.min(range.firstRow, top),
          lastRow: Math.max(range.lastRow, bottom),
          firstColumn: Math.min(range.firstColumn, p.gridStart),
          lastColumn: Math.max(range.lastColumn, right),
        };
        if (
          next.firstRow !== range.firstRow ||
          next.lastRow !== range.lastRow ||
          next.firstColumn !== range.firstColumn ||
          next.lastColumn !== range.lastColumn
        ) {
          range = next;
          changed = true;
        }
      }
    }
  }
  return range;
}

function cellRange(p: TableCellPlacement): TableCellRange {
  return {
    firstRow: p.row,
    lastRow: p.row,
    firstColumn: p.gridStart,
    lastColumn: p.gridStart + p.gridSpan - 1,
  };
}

/** The table cells the selection covers, or undefined when the focus is not in a table. */
export function tableSelection(model: EditorModel): TableSelection | undefined {
  const sel = model.selection;
  const focusCell = sel?.focus.cell;
  if (!sel || !focusCell) return undefined;
  const table = asTable(blockAt(model.doc, sel.focus.block));
  if (!table) return undefined;
  const placements = tableCellPlacements(table);
  const focus = placements[focusCell.row]?.[focusCell.col];
  if (!focus) return undefined;
  const anchorCell = sel.anchor.block === sel.focus.block ? sel.anchor.cell : undefined;
  const anchor = (anchorCell && placements[anchorCell.row]?.[anchorCell.col]) ?? focus;
  const a = cellRange(anchor);
  const f = cellRange(focus);
  const range = expand(placements, {
    firstRow: Math.min(a.firstRow, f.firstRow),
    lastRow: Math.max(a.lastRow, f.lastRow),
    firstColumn: Math.min(a.firstColumn, f.firstColumn),
    lastColumn: Math.max(a.lastColumn, f.lastColumn),
  });
  return {
    block: sel.focus.block,
    table,
    range,
    focus: { row: focusCell.row, cell: focusCell.col },
    placements,
    multiCell: anchor !== focus,
  };
}

/** Every cell (continuations of merges included) that lies in the selected rectangle. */
export function selectedPlacements(ts: TableSelection): TableCellPlacement[] {
  const out: TableCellPlacement[] = [];
  for (let r = ts.range.firstRow; r <= ts.range.lastRow; r++) {
    for (const p of ts.placements[r] ?? []) {
      if (
        p.gridStart + p.gridSpan - 1 >= ts.range.firstColumn &&
        p.gridStart <= ts.range.lastColumn
      ) {
        out.push(p);
      }
    }
  }
  return out;
}

/** A caret at the start of a cell. */
export function cellStart(block: number, row: number, cell: number): DocPosition {
  return { block, cell: { row, col: cell }, para: 0, inline: 0, offset: 0 };
}

/** The cell index covering a grid column in a row, clamped to the row's cells. */
export function cellIndexAt(
  placements: readonly (readonly TableCellPlacement[])[],
  row: number,
  column: number,
): number {
  const placed = placements[row] ?? [];
  const hit = placed.find((p) => p.gridStart <= column && column < p.gridStart + p.gridSpan);
  return hit?.cell ?? Math.max(0, placed.length - 1);
}

export type TableSelectTarget = "cell" | "column" | "row" | "table";

/** The selection Word's Table Layout ▸ Select makes, from the current cell. */
export function selectInTable(model: EditorModel, what: TableSelectTarget): Selection | undefined {
  const ts = tableSelection(model);
  if (!ts) return undefined;
  const { range, placements, block } = ts;
  const lastRow = ts.table.rows.length - 1;
  const lastColumn = tableColumnCount(ts.table) - 1;
  const target =
    what === "table"
      ? { firstRow: 0, lastRow, firstColumn: 0, lastColumn }
      : what === "row"
        ? { ...range, firstColumn: 0, lastColumn }
        : what === "column"
          ? { ...range, firstRow: 0, lastRow }
          : range;
  const first = cellIndexAt(placements, target.firstRow, target.firstColumn);
  const last = cellIndexAt(placements, target.lastRow, target.lastColumn);
  const anchor = cellStart(block, target.firstRow, first);
  const focus = cellStart(block, target.lastRow, last);
  return anchor.cell?.row === focus.cell?.row && first === last
    ? caretAt(anchor)
    : { anchor, focus };
}

/**
 * The start of the next (or previous) cell in reading order, skipping the
 * hidden rows of vertical merges — where Tab / Shift+Tab move in Word.
 * Undefined past the last (or before the first) cell.
 */
export function adjacentCellPosition(
  model: EditorModel,
  direction: 1 | -1,
): DocPosition | undefined {
  const ts = tableSelection(model);
  if (!ts) return undefined;
  const order = ts.placements.flat().filter((p) => p.rowSpan > 0);
  const here = order.findIndex((p) => p.row === ts.focus.row && p.cell === ts.focus.cell);
  const next = order[here + direction];
  return next ? cellStart(ts.block, next.row, next.cell) : undefined;
}
