/**
 * Table editing tools: the operations behind Word's Table Layout and Table
 * Design tabs — insert/delete rows and columns anywhere, merge and split
 * cells, split a table, table/row/cell properties, borders over a cell range,
 * sorting, converting text to a table, and evaluating table formulas.
 *
 * Positions are given on the table *grid* (`<w:tblGrid>`): once a cell spans
 * columns (`w:gridSpan`) or rows (`w:vMerge`), its index in `row.cells` no
 * longer says which column it is in. {@link tableCellPlacements} maps one to
 * the other.
 */

import type { XmlElement } from "../internal/xml/index.js";
import {
  type CellNode,
  cellNodes,
  cellPlacements,
  emptyCellLike,
  emptyParagraphLike,
  ensureTcPr,
  gridAfterOf,
  gridBeforeOf,
  gridColumnCount,
  gridSpanOf,
  gridWidths,
  isEmptyCell,
  normalizeVerticalMerges,
  placementAt,
  setCellNodes,
  setGridSpan,
  setGridWidths,
  setRowGridSkip,
  setVMerge,
  syncCellWidth,
  separatorChar,
  type TableCellPlacement,
  type TableTextSeparator,
  vMergeOf,
} from "../internal/wordprocessingml/table-grid.js";
import {
  cellNumber,
  evaluateFormula,
  formatFieldNumber as formatNumber,
} from "../internal/wordprocessingml/table-formula.js";
import {
  BORDER_ORDER,
  MARGIN_ORDER,
  TBL_PR_ORDER,
  TC_PR_ORDER,
  borderAttrs,
  removeWChild,
  upsertWChild,
  wAttrOf,
  wChild,
  wEl,
} from "../internal/wordprocessingml/table-xml.js";
import { parseTableElement } from "../internal/wordprocessingml/parser.js";
import type {
  WmlParagraph,
  WmlRun,
  WmlRunPiece,
  WmlTable,
  WmlTableCell,
  WmlTableRow,
} from "../internal/wordprocessingml/index.js";
import { type Docx, addTable } from "./docx.js";

export type { TableCellPlacement, TableTextSeparator };

/** A rectangle of cells: rows of `table.rows`, columns of the table grid (both inclusive). */
export interface TableCellRange {
  readonly firstRow: number;
  readonly lastRow: number;
  readonly firstColumn: number;
  readonly lastColumn: number;
}

/** Every cell's position on the table grid, row by row. */
export function tableCellPlacements(table: WmlTable): TableCellPlacement[][] {
  return cellPlacements(table);
}

/** How many grid columns the table has. */
export function tableColumnCount(table: WmlTable): number {
  return gridColumnCount(table);
}

/** Grid column widths in twips (`<w:gridCol w:w>`). */
export function tableColumnWidths(table: WmlTable): number[] {
  return gridWidths(table);
}

function assertIndex(name: string, value: number, max: number): void {
  if (!Number.isInteger(value) || value < 0 || value > max) {
    throw new RangeError(`${name} must be an integer from 0 to ${max}, got ${value}.`);
  }
}

function assertRange(table: WmlTable, range: TableCellRange): void {
  const columns = gridColumnCount(table);
  assertIndex("firstRow", range.firstRow, table.rows.length - 1);
  assertIndex("lastRow", range.lastRow, table.rows.length - 1);
  assertIndex("firstColumn", range.firstColumn, columns - 1);
  assertIndex("lastColumn", range.lastColumn, columns - 1);
  if (range.firstRow > range.lastRow || range.firstColumn > range.lastColumn) {
    throw new RangeError("A cell range must not end before it starts.");
  }
}

/** Make the grid as wide as the widest row (a short `tblGrid` is common in generated files). */
function ensureGrid(table: WmlTable): number[] {
  const widths = gridWidths(table);
  const count = gridColumnCount(table);
  if (widths.length < count) {
    const fill = widths.length ? Math.round(widths.reduce((a, b) => a + b, 0) / widths.length) : 0;
    while (widths.length < count) widths.push(fill);
    setGridWidths(table, widths);
  }
  return widths;
}

function tblPrOf(table: WmlTable): XmlElement {
  table.tblPr ??= wEl("tblPr");
  return table.tblPr;
}

function cellAt(table: WmlTable, p: TableCellPlacement): WmlTableCell {
  const cell = table.rows[p.row]?.cells[p.cell];
  if (!cell) throw new Error(`No cell at row ${p.row}, cell ${p.cell}.`);
  return cell;
}

// --- Rows ---------------------------------------------------------------------

export interface InsertTableRowOptions {
  /**
   * The row whose formatting (row height, cell widths, shading, borders,
   * paragraph formatting) the new row copies. Defaults to the row above the
   * insertion point (the row below when inserting at the top), as Word does.
   */
  readonly formatFrom?: number;
}

/**
 * Insert an empty row at `index` (0 inserts above the first row,
 * `table.rows.length` appends). A vertical merge that spans the insertion
 * point grows to cover the new row, as in Word.
 */
export function insertTableRow(
  table: WmlTable,
  index: number,
  options: InsertTableRowOptions = {},
): WmlTableRow {
  assertIndex("index", index, table.rows.length);
  const refIndex = options.formatFrom ?? (index > 0 ? index - 1 : 0);
  const ref = table.rows[refIndex];
  if (!ref) throw new RangeError(`formatFrom ${refIndex} is not a row of the table.`);
  const placements = cellPlacements(table);
  const above = placements[index - 1] ?? [];
  const below = placements[index] ?? [];
  const trPr = ref.trPr ? structuredClone(ref.trPr) : undefined;
  // A repeated header row stays a header only when inserted inside the header block.
  if (trPr && !(table.rows[index]?.trPr && wChild(table.rows[index]?.trPr, "tblHeader"))) {
    removeWChild(trPr, "tblHeader");
  }
  const cells = ref.cells.map((cell, c) => {
    const fresh = emptyCellLike(cell);
    const start = placements[refIndex]?.[c]?.gridStart ?? -1;
    const up = above.find((p) => p.gridStart === start);
    const down = below.find((p) => p.gridStart === start);
    if (up?.vMerge && down?.vMerge === "continue") setVMerge(fresh, "continue");
    setGridSpan(fresh, gridSpanOf(cell));
    return fresh;
  });
  const row: WmlTableRow = { ...(trPr ? { trPr } : {}), cells, extras: [] };
  table.rows.splice(index, 0, row);
  normalizeVerticalMerges(table);
  return row;
}

/**
 * Remove rows `first`…`last` (inclusive). A vertical merge that started in a
 * removed row restarts in the first remaining row.
 */
export function deleteTableRows(table: WmlTable, first: number, last = first): void {
  assertIndex("first", first, table.rows.length - 1);
  assertIndex("last", last, table.rows.length - 1);
  if (last < first) throw new RangeError("last must not be before first.");
  // A removed merge start hands its content to the cell that continues it.
  const placements = cellPlacements(table);
  const following = placements[last + 1] ?? [];
  for (const p of following) {
    if (p.vMerge !== "continue") continue;
    const owner = placements
      .slice(first, last + 1)
      .flat()
      .find((q) => q.gridStart === p.gridStart && q.vMerge === "restart");
    if (owner) {
      const target = cellAt(table, p);
      if (isEmptyCell(target)) setCellNodes(target, cellNodes(cellAt(table, owner)));
    }
  }
  table.rows.splice(first, last - first + 1);
  normalizeVerticalMerges(table);
}

// --- Columns ------------------------------------------------------------------

export interface InsertTableColumnOptions {
  /** Width of the new grid column in twips; defaults to the width of the column it copies. */
  readonly widthTwips?: number;
  /** The grid column whose cell formatting the new cells copy; defaults to the one left of `index`. */
  readonly formatFrom?: number;
}

/**
 * Insert an empty grid column before grid column `index` (`tableColumnCount`
 * appends). A cell that spans across the insertion point widens instead of
 * being split. A table whose width is fixed in twips grows by the new
 * column's width.
 */
export function insertTableColumn(
  table: WmlTable,
  index: number,
  options: InsertTableColumnOptions = {},
): void {
  const count = gridColumnCount(table);
  assertIndex("index", index, count);
  const widths = ensureGrid(table);
  const refColumn = options.formatFrom ?? (index > 0 ? index - 1 : 0);
  const width = options.widthTwips ?? widths[refColumn] ?? 0;
  if (!Number.isFinite(width) || width < 0) throw new RangeError("widthTwips must be ≥ 0.");
  const placements = cellPlacements(table);
  table.rows.forEach((row, r) => {
    const placed = placements[r] ?? [];
    const before = gridBeforeOf(row);
    const end = placed.length ? placed[placed.length - 1]!.gridStart + placed[placed.length - 1]!.gridSpan : before;
    if (index < before) {
      setRowGridSkip(row, "gridBefore", before + 1);
      return;
    }
    if (index > end) {
      setRowGridSkip(row, "gridAfter", gridAfterOf(row) + 1);
      return;
    }
    const inside = placed.find((p) => p.gridStart < index && index < p.gridStart + p.gridSpan);
    if (inside) {
      setGridSpan(cellAt(table, inside), inside.gridSpan + 1);
      return;
    }
    const at = placed.filter((p) => p.gridStart < index).length;
    const ref = placementAt(placed, refColumn) ?? placed[Math.max(0, at - 1)];
    const fresh = emptyCellLike(ref ? cellAt(table, ref) : undefined);
    row.cells.splice(at, 0, fresh);
  });
  widths.splice(index, 0, width);
  setGridWidths(table, widths);
  syncAllCellWidths(table);
  growFixedTableWidth(table, width);
  normalizeVerticalMerges(table);
}

function growFixedTableWidth(table: WmlTable, delta: number): void {
  const tblW = wChild(table.tblPr, "tblW");
  if (!tblW || (wAttrOf(tblW, "type") ?? "dxa") !== "dxa") return;
  const current = Number(wAttrOf(tblW, "w") ?? 0);
  upsertWChild(
    tblPrOf(table),
    wEl("tblW", { w: String(Math.max(0, Math.round(current + delta))), type: "dxa" }),
    TBL_PR_ORDER,
  );
}

function syncAllCellWidths(table: WmlTable): void {
  const widths = gridWidths(table);
  for (const row of cellPlacements(table)) {
    for (const p of row) syncCellWidth(cellAt(table, p), widths, p.gridStart);
  }
}

/**
 * Remove grid columns `first`…`last` (inclusive). Cells inside the range are
 * removed; cells that span into it get narrower. Returns false when no column
 * is left (the caller should then remove the table, as Word does).
 */
export function deleteTableColumns(table: WmlTable, first: number, last = first): boolean {
  const count = gridColumnCount(table);
  assertIndex("first", first, count - 1);
  assertIndex("last", last, count - 1);
  if (last < first) throw new RangeError("last must not be before first.");
  const widths = ensureGrid(table);
  const removedWidth = widths.slice(first, last + 1).reduce((a, b) => a + b, 0);
  const placements = cellPlacements(table);
  table.rows.forEach((row, r) => {
    const before = gridBeforeOf(row);
    const overlapBefore = Math.max(0, Math.min(before, last + 1) - first);
    if (overlapBefore > 0) setRowGridSkip(row, "gridBefore", before - overlapBefore);
    const placed = placements[r] ?? [];
    const end = placed.length ? placed[placed.length - 1]!.gridStart + placed[placed.length - 1]!.gridSpan : before;
    const after = gridAfterOf(row);
    const overlapAfter = Math.max(0, Math.min(end + after, last + 1) - Math.max(end, first));
    if (overlapAfter > 0) setRowGridSkip(row, "gridAfter", after - overlapAfter);
    const keep: WmlTableCell[] = [];
    for (const p of placed) {
      const cell = cellAt(table, p);
      const from = Math.max(p.gridStart, first);
      const to = Math.min(p.gridStart + p.gridSpan - 1, last);
      const overlap = Math.max(0, to - from + 1);
      if (overlap === p.gridSpan) continue;
      if (overlap > 0) setGridSpan(cell, p.gridSpan - overlap);
      keep.push(cell);
    }
    row.cells = keep;
  });
  widths.splice(first, last - first + 1);
  setGridWidths(table, widths);
  table.rows = table.rows.filter((row) => row.cells.length > 0);
  if (widths.length === 0 || table.rows.length === 0) return false;
  syncAllCellWidths(table);
  growFixedTableWidth(table, -removedWidth);
  normalizeVerticalMerges(table);
  return true;
}

/**
 * Word's Delete Cells: remove the cells of `range` and shift the cells to
 * their right left (leaving the rows shorter), or shift the cells below up
 * (moving their content up and leaving empty cells at the bottom).
 */
export function deleteTableCells(
  table: WmlTable,
  range: TableCellRange,
  shift: "left" | "up",
): void {
  assertRange(table, range);
  const placements = cellPlacements(table);
  if (shift === "left") {
    for (let r = range.firstRow; r <= range.lastRow; r++) {
      const row = table.rows[r];
      if (!row) continue;
      const placed = placements[r] ?? [];
      const doomed = new Set(
        placed
          .filter((p) => p.gridStart >= range.firstColumn && p.gridStart <= range.lastColumn)
          .map((p) => p.cell),
      );
      const removedSpan = placed
        .filter((p) => doomed.has(p.cell))
        .reduce((a, p) => a + p.gridSpan, 0);
      row.cells = row.cells.filter((_, c) => !doomed.has(c));
      // The grid keeps its columns; the row now ends early (w:gridAfter).
      setRowGridSkip(row, "gridAfter", gridAfterOf(row) + removedSpan);
    }
    table.rows = table.rows.filter((row) => row.cells.length > 0);
    normalizeVerticalMerges(table);
    return;
  }
  const height = range.lastRow - range.firstRow + 1;
  for (let g = range.firstColumn; g <= range.lastColumn; g++) {
    const column = placements
      .map((row) => placementAt(row, g))
      .filter((p): p is TableCellPlacement => p !== undefined && p.gridStart === g);
    const cells = column.filter((p) => p.row >= range.firstRow).map((p) => cellAt(table, p));
    for (let i = 0; i < cells.length; i++) {
      const target = cells[i]!;
      const source = cells[i + height];
      setCellNodes(
        target,
        source
          ? cellNodes(source)
          : [{ kind: "paragraph", paragraph: emptyParagraphLike(target.paragraphs[0]) }],
      );
    }
  }
}

// --- Merge / split ------------------------------------------------------------

/**
 * Merge a rectangle of cells into one (Word's Merge Cells). The content of
 * every non-empty cell is kept, in reading order, in the merged cell. Throws
 * when the rectangle cuts through a cell that spans past its edge.
 */
export function mergeTableCells(table: WmlTable, range: TableCellRange): void {
  assertRange(table, range);
  const widths = ensureGrid(table);
  const placements = cellPlacements(table);
  const content: CellNode[] = [];
  const rows: TableCellPlacement[][] = [];
  for (let r = range.firstRow; r <= range.lastRow; r++) {
    const inRange = (placements[r] ?? []).filter(
      (p) => p.gridStart + p.gridSpan > range.firstColumn && p.gridStart <= range.lastColumn,
    );
    const first = inRange[0];
    const last = inRange[inRange.length - 1];
    if (
      !first ||
      !last ||
      first.gridStart !== range.firstColumn ||
      last.gridStart + last.gridSpan - 1 !== range.lastColumn
    ) {
      throw new RangeError("The cells to merge must form a rectangle.");
    }
    for (const p of inRange) {
      const cell = cellAt(table, p);
      if (!isEmptyCell(cell)) content.push(...cellNodes(cell));
    }
    rows.push(inRange);
  }
  // Merges that reach in from above or out below would be cut in two.
  const top = rows[0] ?? [];
  if (top.some((p) => p.vMerge === "continue" && p.rowSpan === 0)) {
    throw new RangeError("The cells to merge must form a rectangle.");
  }
  const below = placements[range.lastRow + 1] ?? [];
  if (
    below.some(
      (p) =>
        p.vMerge === "continue" &&
        p.gridStart >= range.firstColumn &&
        p.gridStart <= range.lastColumn,
    )
  ) {
    throw new RangeError("The cells to merge must form a rectangle.");
  }
  const span = range.lastColumn - range.firstColumn + 1;
  const multiRow = range.lastRow > range.firstRow;
  rows.forEach((inRange, i) => {
    const row = table.rows[range.firstRow + i];
    const first = inRange[0];
    if (!row || !first) return;
    const merged = cellAt(table, first);
    setGridSpan(merged, span);
    syncCellWidth(merged, widths, range.firstColumn);
    setVMerge(merged, multiRow ? (i === 0 ? "restart" : "continue") : undefined);
    if (i === 0) {
      setCellNodes(
        merged,
        content.length
          ? content
          : [{ kind: "paragraph", paragraph: emptyParagraphLike(merged.paragraphs[0]) }],
      );
    } else {
      setCellNodes(merged, [
        { kind: "paragraph", paragraph: emptyParagraphLike(merged.paragraphs[0]) },
      ]);
    }
    const drop = new Set(inRange.slice(1).map((p) => p.cell));
    row.cells = row.cells.filter((_, c) => !drop.has(c));
  });
  normalizeVerticalMerges(table);
}

export interface SplitTableCellOptions {
  /** Number of columns to split into (≥ 1). */
  readonly columns: number;
  /** Number of rows to split into (≥ 1). */
  readonly rows: number;
}

// Word's Split Cells dialog accepts up to 63 columns.
const MAX_SPLIT = 63;

/**
 * Split a cell into `columns` × `rows` cells (Word's Split Cells). The
 * original content stays in the first new cell. Splitting into more columns
 * than the cell spans adds grid columns, which the other rows' cells span so
 * their widths do not change. A vertically merged cell splits into rows by
 * dividing the merge.
 */
export function splitTableCell(
  table: WmlTable,
  row: number,
  cell: number,
  options: SplitTableCellOptions,
): void {
  const { columns, rows } = options;
  for (const [name, value] of [
    ["columns", columns],
    ["rows", rows],
  ] as const) {
    if (!Number.isInteger(value) || value < 1 || value > MAX_SPLIT) {
      throw new RangeError(`${name} must be an integer from 1 to ${MAX_SPLIT}, got ${value}.`);
    }
  }
  const start = cellPlacements(table)[row]?.[cell];
  if (!start) throw new RangeError(`No cell ${cell} in row ${row}.`);
  if (start.rowSpan === 0) throw new RangeError("Split the first cell of a merged cell.");
  if (columns > 1) splitColumns(table, start, columns);
  if (rows > 1) splitRows(table, row, start.gridStart, columns, rows);
  normalizeVerticalMerges(table);
}

/** Split a cell (and every row of its vertical merge) into `count` side-by-side cells. */
function splitColumns(table: WmlTable, original: TableCellPlacement, count: number): void {
  let widths = ensureGrid(table);
  let start = original;
  const firstGrid = start.gridStart;
  let span = start.gridSpan;
  if (span < count) {
    // Re-grid the cell's extent into equal parts; other rows' cells keep their
    // edges by spanning the new, narrower columns.
    const total = widths.slice(firstGrid, firstGrid + span).reduce((a, b) => a + b, 0);
    const edges = new Set<number>();
    let x = 0;
    for (let g = 0; g <= widths.length; g++) {
      edges.add(x);
      x += widths[g] ?? 0;
    }
    const left = widths.slice(0, firstGrid).reduce((a, b) => a + b, 0);
    for (let k = 1; k < count; k++) edges.add(Math.round(left + (total * k) / count));
    regrid(table, [...edges].toSorted((a, b) => a - b));
    widths = gridWidths(table);
    const placed = cellPlacements(table)[start.row]?.[start.cell];
    if (!placed) return;
    span = placed.gridSpan;
    start = placed;
  }
  const placements = cellPlacements(table);
  // Spans of the new cells: as even as the grid allows.
  const spans = Array.from(
    { length: count },
    (_, k) => Math.floor((span * (k + 1)) / count) - Math.floor((span * k) / count),
  );
  for (let r = start.row; r < start.row + start.rowSpan; r++) {
    const p = placements[r]?.find((q) => q.gridStart === start.gridStart);
    const row = table.rows[r];
    if (!p || !row) continue;
    const original = cellAt(table, p);
    let g = p.gridStart;
    const parts = spans.map((s, k) => {
      const part = k === 0 ? original : emptyCellLike(original);
      if (k > 0) setVMerge(part, vMergeOf(original));
      setGridSpan(part, s);
      syncCellWidth(part, widths, g);
      g += s;
      return part;
    });
    row.cells.splice(p.cell, 1, ...parts);
  }
}

/** Split the cells covering grid columns of the split cell into `count` rows. */
function splitRows(
  table: WmlTable,
  rowIndex: number,
  gridStart: number,
  columns: number,
  count: number,
): void {
  const placements = cellPlacements(table);
  const targets = (placements[rowIndex] ?? []).filter(
    (p) => p.gridStart >= gridStart,
  ).slice(0, columns);
  const height = targets[0]?.rowSpan ?? 1;
  if (height > 1) {
    if (count > height) {
      throw new RangeError(`A cell merged over ${height} rows splits into at most ${height} rows.`);
    }
    // Divide the merge into `count` merges of near-equal height.
    for (const t of targets) {
      for (let k = 0; k < count; k++) {
        const from = rowIndex + Math.floor((height * k) / count);
        const to = rowIndex + Math.floor((height * (k + 1)) / count) - 1;
        for (let r = from; r <= to; r++) {
          const p = placements[r]?.find((q) => q.gridStart === t.gridStart);
          if (p) setVMerge(cellAt(table, p), to > from ? (r === from ? "restart" : "continue") : undefined);
        }
      }
    }
    return;
  }
  const splitStarts = new Set(targets.map((t) => t.gridStart));
  const source = table.rows[rowIndex];
  if (!source) return;
  for (const p of placements[rowIndex] ?? []) {
    if (!splitStarts.has(p.gridStart) && !p.vMerge) setVMerge(cellAt(table, p), "restart");
  }
  for (let k = 1; k < count; k++) {
    const trPr = source.trPr ? structuredClone(source.trPr) : undefined;
    if (trPr) removeWChild(trPr, "tblHeader");
    const cells = source.cells.map((cell, c) => {
      const p = placements[rowIndex]?.[c];
      const fresh = emptyCellLike(cell);
      setGridSpan(fresh, gridSpanOf(cell));
      if (p && !splitStarts.has(p.gridStart)) setVMerge(fresh, "continue");
      return fresh;
    });
    table.rows.splice(rowIndex + k, 0, { ...(trPr ? { trPr } : {}), cells, extras: [] });
  }
}

/**
 * Replace the grid with columns between the given x positions (twips from the
 * table's left edge, sorted, starting at 0); every cell keeps its edges.
 */
function regrid(table: WmlTable, edges: readonly number[]): void {
  const old = gridWidths(table);
  const oldEdges = [0];
  for (const w of old) oldEdges.push(oldEdges[oldEdges.length - 1]! + w);
  const index = new Map(edges.map((x, i) => [x, i]));
  const at = (g: number): number => index.get(oldEdges[Math.min(g, oldEdges.length - 1)]!) ?? 0;
  const placements = cellPlacements(table);
  table.rows.forEach((row, r) => {
    const before = gridBeforeOf(row);
    const placed = placements[r] ?? [];
    for (const p of placed) {
      setGridSpan(cellAt(table, p), at(p.gridStart + p.gridSpan) - at(p.gridStart));
    }
    if (before) setRowGridSkip(row, "gridBefore", at(before));
    const end = placed.length ? placed[placed.length - 1]!.gridStart + placed[placed.length - 1]!.gridSpan : before;
    const after = gridAfterOf(row);
    if (after) setRowGridSkip(row, "gridAfter", at(end + after) - at(end));
  });
  setGridWidths(
    table,
    edges.slice(1).map((x, i) => x - edges[i]!),
  );
}

/**
 * Split a body table into two before row `row` (Word's Split Table), with an
 * empty paragraph between them. Splitting before the first row inserts the
 * paragraph above the table instead. Returns the table that holds `row`.
 */
export function splitTable(doc: Docx, table: WmlTable, row: number): WmlTable {
  const blocks = doc.document.body.blocks;
  const at = blocks.indexOf(table);
  if (at < 0) throw new Error("splitTable: the table is not in the document body.");
  assertIndex("row", row, table.rows.length - 1);
  const gap: WmlParagraph = { kind: "paragraph", children: [], extras: [] };
  doc.dirty = true;
  if (row === 0) {
    blocks.splice(at, 0, gap);
    return table;
  }
  const second: WmlTable = {
    kind: "table",
    ...(table.tblPr ? { tblPr: structuredClone(table.tblPr) } : {}),
    ...(table.tblGrid ? { tblGrid: structuredClone(table.tblGrid) } : {}),
    rows: table.rows.splice(row),
    extras: [],
  };
  normalizeVerticalMerges(table);
  normalizeVerticalMerges(second);
  blocks.splice(at + 1, 0, gap, second);
  return second;
}

// --- Table properties ---------------------------------------------------------

/**
 * A preferred width (ST_TblWidth, §17.18.90). `pct` is in fiftieths of a
 * percent (5000 = 100 %), `dxa` in twips; `auto` and `nil` ignore `value`.
 */
export interface TableWidth {
  readonly type: "auto" | "dxa" | "pct" | "nil";
  readonly value: number;
}

function widthAttrs(width: TableWidth): Record<string, string> {
  if (!Number.isFinite(width.value) || width.value < 0) {
    throw new RangeError(`A width must be ≥ 0, got ${width.value}.`);
  }
  const value = width.type === "auto" || width.type === "nil" ? 0 : Math.round(width.value);
  return { w: String(value), type: width.type };
}

/** The table's preferred width (`w:tblW`); undefined removes it. */
export function setTableWidth(table: WmlTable, width: TableWidth | undefined): void {
  if (width === undefined) removeWChild(tblPrOf(table), "tblW");
  else upsertWChild(tblPrOf(table), wEl("tblW", widthAttrs(width)), TBL_PR_ORDER);
}

export type TableAlignment = "left" | "center" | "right";

/** How the table sits between the margins (`w:jc`); undefined is left (the default). */
export function setTableAlignment(table: WmlTable, alignment: TableAlignment | undefined): void {
  if (alignment === undefined) removeWChild(tblPrOf(table), "jc");
  else upsertWChild(tblPrOf(table), wEl("jc", { val: alignment }), TBL_PR_ORDER);
}

/** Indent from the leading margin in twips (`w:tblInd`); undefined removes it. */
export function setTableIndent(table: WmlTable, twips: number | undefined): void {
  if (twips === undefined) {
    removeWChild(tblPrOf(table), "tblInd");
    return;
  }
  if (!Number.isFinite(twips)) throw new RangeError("The indent must be a number of twips.");
  upsertWChild(
    tblPrOf(table),
    wEl("tblInd", { w: String(Math.round(twips)), type: "dxa" }),
    TBL_PR_ORDER,
  );
}

/** `fixed` keeps the column widths; `autofit` lets them follow the content (`w:tblLayout`). */
export function setTableLayout(table: WmlTable, layout: "fixed" | "autofit"): void {
  // autofit is the default (§17.4.53), so Word writes nothing for it.
  if (layout === "autofit") removeWChild(tblPrOf(table), "tblLayout");
  else upsertWChild(tblPrOf(table), wEl("tblLayout", { type: "fixed" }), TBL_PR_ORDER);
}

/**
 * Spacing between cells (`w:tblCellSpacing`) in twips; undefined or 0 removes
 * it. Word's "Allow spacing between cells" writes half the spacing it shows
 * here, because each cell's edge is inset by the value on every side.
 */
export function setTableCellSpacing(table: WmlTable, twips: number | undefined): void {
  if (!twips) {
    removeWChild(tblPrOf(table), "tblCellSpacing");
    return;
  }
  if (!Number.isFinite(twips) || twips < 0) throw new RangeError("Spacing must be ≥ 0.");
  upsertWChild(
    tblPrOf(table),
    wEl("tblCellSpacing", { w: String(Math.round(twips)), type: "dxa" }),
    TBL_PR_ORDER,
  );
}

/** Cell margins in twips; an omitted side is left out. */
export interface CellMargins {
  readonly top?: number;
  readonly left?: number;
  readonly bottom?: number;
  readonly right?: number;
}

function marginsEl(local: string, margins: CellMargins): XmlElement | undefined {
  const children: XmlElement[] = [];
  for (const side of MARGIN_ORDER) {
    if (side !== "top" && side !== "left" && side !== "bottom" && side !== "right") continue;
    const value = margins[side];
    if (value === undefined) continue;
    if (!Number.isFinite(value) || value < 0) {
      throw new RangeError(`The ${side} margin must be ≥ 0, got ${value}.`);
    }
    children.push(wEl(side, { w: String(Math.round(value)), type: "dxa" }));
  }
  return children.length ? wEl(local, {}, children) : undefined;
}

/** The table's default cell margins (`w:tblCellMar`); undefined removes them. */
export function setTableDefaultCellMargins(table: WmlTable, margins: CellMargins | undefined): void {
  const el = margins && marginsEl("tblCellMar", margins);
  if (el) upsertWChild(tblPrOf(table), el, TBL_PR_ORDER);
  else removeWChild(tblPrOf(table), "tblCellMar");
}

/**
 * Alternative text for the table (Table Properties ▸ Alt Text): its title
 * (`w:tblCaption`) and description (`w:tblDescription`), both ISO/IEC 29500
 * additions to `<w:tblPr>`. An empty or undefined value removes the element.
 */
export function setTableAltText(
  table: WmlTable,
  alt: { readonly title?: string | undefined; readonly description?: string | undefined },
): void {
  const tblPr = tblPrOf(table);
  for (const [local, value] of [
    ["tblCaption", alt.title],
    ["tblDescription", alt.description],
  ] as const) {
    if (value) upsertWChild(tblPr, wEl(local, { val: value }), TBL_PR_ORDER);
    else removeWChild(tblPr, local);
  }
}

/**
 * A floating table's position (`w:tblpPr`, §17.4.58) — Word's "Text wrapping:
 * Around". Distances are in twips; `x` / `y` are offsets from the anchors,
 * and `xAlign` / `yAlign` (when given) replace them with a relative position.
 */
export interface TablePosition {
  readonly horizontalAnchor?: "text" | "margin" | "page";
  readonly verticalAnchor?: "text" | "margin" | "page";
  readonly x?: number;
  readonly xAlign?: "left" | "center" | "right" | "inside" | "outside";
  readonly y?: number;
  readonly yAlign?: "top" | "center" | "bottom" | "inside" | "outside" | "inline";
  readonly leftFromText?: number;
  readonly rightFromText?: number;
  readonly topFromText?: number;
  readonly bottomFromText?: number;
  /** Whether other floating tables may overlap this one (`w:tblOverlap`). */
  readonly allowOverlap?: boolean;
}

// What Word writes when "Around" is first chosen: ⅛ in to the sides of the text.
const DEFAULT_FROM_TEXT = 180;

function twips(v: number | undefined): string | undefined {
  return v === undefined ? undefined : String(Math.round(v));
}

/** Make the table float with text wrapping around it, or (undefined) put it back in line. */
export function setTablePosition(table: WmlTable, position: TablePosition | undefined): void {
  const tblPr = tblPrOf(table);
  if (!position) {
    removeWChild(tblPr, "tblpPr");
    removeWChild(tblPr, "tblOverlap");
    return;
  }
  upsertWChild(
    tblPr,
    wEl("tblpPr", {
      leftFromText: twips(position.leftFromText ?? DEFAULT_FROM_TEXT),
      rightFromText: twips(position.rightFromText ?? DEFAULT_FROM_TEXT),
      topFromText: twips(position.topFromText),
      bottomFromText: twips(position.bottomFromText),
      vertAnchor: position.verticalAnchor ?? "text",
      horzAnchor: position.horizontalAnchor,
      tblpXSpec: position.xAlign,
      tblpX: twips(position.x),
      tblpYSpec: position.yAlign,
      tblpY: twips(position.y),
    }),
    TBL_PR_ORDER,
  );
  if (position.allowOverlap === undefined) removeWChild(tblPr, "tblOverlap");
  else
    upsertWChild(
      tblPr,
      wEl("tblOverlap", { val: position.allowOverlap ? "overlap" : "never" }),
      TBL_PR_ORDER,
    );
}

// --- Cell properties ----------------------------------------------------------

/** A cell's preferred width (`w:tcW`); undefined removes it. */
export function setTableCellWidth(cell: WmlTableCell, width: TableWidth | undefined): void {
  if (width === undefined) removeWChild(ensureTcPr(cell), "tcW");
  else upsertWChild(ensureTcPr(cell), wEl("tcW", widthAttrs(width)), TC_PR_ORDER);
}

/** A cell's own margins (`w:tcMar`), overriding the table's; undefined removes them. */
export function setTableCellMargins(cell: WmlTableCell, margins: CellMargins | undefined): void {
  const el = margins && marginsEl("tcMar", margins);
  if (el) upsertWChild(ensureTcPr(cell), el, TC_PR_ORDER);
  else removeWChild(ensureTcPr(cell), "tcMar");
}

/**
 * Text flow in a cell (`w:textDirection`, ST_TextDirection): `lrTb` is
 * horizontal (the default), `tbRl` reads top to bottom, `btLr` bottom to top;
 * the `V` variants keep East Asian characters upright.
 */
export type TableCellTextDirection = "lrTb" | "tbRl" | "btLr" | "lrTbV" | "tbRlV" | "tbLrV";

export function setTableCellTextDirection(
  cell: WmlTableCell,
  direction: TableCellTextDirection | undefined,
): void {
  if (direction === undefined || direction === "lrTb") removeWChild(ensureTcPr(cell), "textDirection");
  else upsertWChild(ensureTcPr(cell), wEl("textDirection", { val: direction }), TC_PR_ORDER);
}

// --- Column widths / AutoFit --------------------------------------------------

/** Set one grid column's width in twips; the cells over it follow. */
export function setTableColumnWidth(table: WmlTable, column: number, twips: number): void {
  const widths = ensureGrid(table);
  assertIndex("column", column, widths.length - 1);
  if (!Number.isFinite(twips) || twips < 0) throw new RangeError("A width must be ≥ 0.");
  const delta = Math.round(twips) - (widths[column] ?? 0);
  widths[column] = Math.round(twips);
  setGridWidths(table, widths);
  syncAllCellWidths(table);
  growFixedTableWidth(table, delta);
}

/** Give grid columns `first`…`last` the same width, keeping their total (Distribute Columns). */
export function distributeTableColumns(table: WmlTable, first: number, last: number): void {
  const widths = ensureGrid(table);
  assertIndex("first", first, widths.length - 1);
  assertIndex("last", last, widths.length - 1);
  if (last < first) throw new RangeError("last must not be before first.");
  const total = widths.slice(first, last + 1).reduce((a, b) => a + b, 0);
  const count = last - first + 1;
  for (let g = first; g <= last; g++) {
    widths[g] = Math.floor((total * (g - first + 1)) / count) - Math.floor((total * (g - first)) / count);
  }
  setGridWidths(table, widths);
  syncAllCellWidths(table);
}

// A full-width table in fiftieths of a percent.
const FULL_WIDTH_PCT = 5000;

/**
 * Word's AutoFit: `contents` sizes the columns to their content, `window`
 * stretches the table across the text width, `fixed` freezes the current
 * column widths.
 */
export function autoFitTable(table: WmlTable, mode: "contents" | "window" | "fixed"): void {
  const widths = ensureGrid(table);
  const total = widths.reduce((a, b) => a + b, 0);
  const placements = cellPlacements(table);
  if (mode === "contents") {
    setTableLayout(table, "autofit");
    setTableWidth(table, { type: "auto", value: 0 });
    for (const row of placements) {
      for (const p of row) setTableCellWidth(cellAt(table, p), { type: "auto", value: 0 });
    }
    return;
  }
  if (mode === "window") {
    setTableLayout(table, "autofit");
    setTableWidth(table, { type: "pct", value: FULL_WIDTH_PCT });
    for (const row of placements) {
      for (const p of row) {
        const share = widths.slice(p.gridStart, p.gridStart + p.gridSpan).reduce((a, b) => a + b, 0);
        setTableCellWidth(cellAt(table, p), {
          type: "pct",
          value: total ? Math.round((FULL_WIDTH_PCT * share) / total) : 0,
        });
      }
    }
    return;
  }
  setTableLayout(table, "fixed");
  const tblW = wChild(table.tblPr, "tblW");
  if (wAttrOf(tblW, "type") === "pct") setTableWidth(table, { type: "dxa", value: total });
  for (const row of placements) {
    for (const p of row) {
      const cell = cellAt(table, p);
      setTableCellWidth(cell, { type: "dxa", value: 0 });
      syncCellWidth(cell, widths, p.gridStart);
    }
  }
}

// --- Borders ------------------------------------------------------------------

/** One border line (`<w:top w:val w:sz w:color/>` …). */
export interface TableBorder {
  /** ST_Border line style, e.g. `single`, `double`, `dotted`, `thickThinSmallGap`. */
  readonly style: string;
  /** Width in eighths of a point (Word offers ¼ pt = 2 to 6 pt = 48). */
  readonly size: number;
  /** Hex RGB or `auto`. */
  readonly color: string;
}

/**
 * Which edges of a cell range a border command sets — the entries of Word's
 * Borders menu. `none` clears every edge of the range.
 */
export type TableBorderEdges =
  | "bottom"
  | "top"
  | "left"
  | "right"
  | "none"
  | "all"
  | "outside"
  | "inside"
  | "insideH"
  | "insideV"
  | "tl2br"
  | "tr2bl";

type CellSide = "top" | "left" | "bottom" | "right" | "tl2br" | "tr2bl";

function setCellBorder(cell: WmlTableCell, side: CellSide, attrs: Record<string, string>): void {
  const tcPr = ensureTcPr(cell);
  const borders = wChild(tcPr, "tcBorders") ?? upsertWChild(tcPr, wEl("tcBorders"), TC_PR_ORDER);
  upsertWChild(borders, wEl(side, attrs), BORDER_ORDER);
}

function setTableLevelBorder(table: WmlTable, side: string, attrs: Record<string, string>): void {
  const tblPr = tblPrOf(table);
  const borders = wChild(tblPr, "tblBorders") ?? upsertWChild(tblPr, wEl("tblBorders"), TBL_PR_ORDER);
  upsertWChild(borders, wEl(side, attrs), BORDER_ORDER);
}

function removeCellBorder(cell: WmlTableCell, side: CellSide): void {
  const borders = wChild(cell.tcPr, "tcBorders");
  if (!borders) return;
  removeWChild(borders, side);
  if (borders.children.length === 0 && cell.tcPr) removeWChild(cell.tcPr, "tcBorders");
}

/**
 * Apply a border (or, with `undefined`, no border) to edges of a cell range,
 * like Word's Borders menu. On the whole table the outer and inside borders
 * go to the table (`w:tblBorders`) and cell overrides of those edges are
 * dropped; on part of it they go to the cells (`w:tcBorders`), and the
 * neighbouring cells' facing edges follow so the shared line agrees.
 */
export function setTableRangeBorders(
  table: WmlTable,
  range: TableCellRange,
  edges: TableBorderEdges,
  border: TableBorder | undefined,
): void {
  assertRange(table, range);
  const attrs = borderAttrs(edges === "none" ? undefined : border);
  const columns = gridColumnCount(table);
  const whole =
    range.firstRow === 0 &&
    range.firstColumn === 0 &&
    range.lastRow === table.rows.length - 1 &&
    range.lastColumn === columns - 1;
  const wants = {
    top: edges === "top" || edges === "all" || edges === "outside" || edges === "none",
    bottom: edges === "bottom" || edges === "all" || edges === "outside" || edges === "none",
    left: edges === "left" || edges === "all" || edges === "outside" || edges === "none",
    right: edges === "right" || edges === "all" || edges === "outside" || edges === "none",
    insideH: edges === "insideH" || edges === "inside" || edges === "all" || edges === "none",
    insideV: edges === "insideV" || edges === "inside" || edges === "all" || edges === "none",
  };
  const placements = cellPlacements(table);
  if (whole && edges !== "tl2br" && edges !== "tr2bl") {
    for (const [side, on] of Object.entries(wants)) if (on) setTableLevelBorder(table, side, attrs);
    for (const row of placements) {
      for (const p of row) {
        const cell = cellAt(table, p);
        const last = p.gridStart + p.gridSpan - 1;
        if (wants.top && p.row === 0) removeCellBorder(cell, "top");
        if (wants.insideH && p.row > 0) removeCellBorder(cell, "top");
        if (wants.bottom && p.row === table.rows.length - 1) removeCellBorder(cell, "bottom");
        if (wants.insideH && p.row < table.rows.length - 1) removeCellBorder(cell, "bottom");
        if (wants.left && p.gridStart === 0) removeCellBorder(cell, "left");
        if (wants.insideV && p.gridStart > 0) removeCellBorder(cell, "left");
        if (wants.right && last === columns - 1) removeCellBorder(cell, "right");
        if (wants.insideV && last < columns - 1) removeCellBorder(cell, "right");
      }
    }
    return;
  }
  const inRange = (p: TableCellPlacement): boolean =>
    p.row >= range.firstRow &&
    p.row <= range.lastRow &&
    p.gridStart + p.gridSpan > range.firstColumn &&
    p.gridStart <= range.lastColumn;
  for (const row of placements) {
    for (const p of row) {
      const cell = cellAt(table, p);
      const last = p.gridStart + p.gridSpan - 1;
      if (inRange(p)) {
        if (edges === "tl2br" || edges === "tr2bl") {
          setCellBorder(cell, edges, attrs);
          continue;
        }
        const top = p.row === range.firstRow ? wants.top : wants.insideH;
        const bottom = p.row === range.lastRow ? wants.bottom : wants.insideH;
        const left = p.gridStart <= range.firstColumn ? wants.left : wants.insideV;
        const right = last >= range.lastColumn ? wants.right : wants.insideV;
        if (top) setCellBorder(cell, "top", attrs);
        if (bottom) setCellBorder(cell, "bottom", attrs);
        if (left) setCellBorder(cell, "left", attrs);
        if (right) setCellBorder(cell, "right", attrs);
        continue;
      }
      if (edges === "tl2br" || edges === "tr2bl") continue;
      // Neighbours outside the range share the range's outer edges.
      const overlapsColumns = p.gridStart <= range.lastColumn && last >= range.firstColumn;
      const overlapsRows = p.row >= range.firstRow && p.row <= range.lastRow;
      if (wants.top && overlapsColumns && p.row === range.firstRow - 1) setCellBorder(cell, "bottom", attrs);
      if (wants.bottom && overlapsColumns && p.row === range.lastRow + 1) setCellBorder(cell, "top", attrs);
      if (wants.left && overlapsRows && last === range.firstColumn - 1) setCellBorder(cell, "right", attrs);
      if (wants.right && overlapsRows && p.gridStart === range.lastColumn + 1) setCellBorder(cell, "left", attrs);
    }
  }
}

// --- Sort ---------------------------------------------------------------------

export interface TableSortKey {
  /** Grid column to sort by. */
  readonly column: number;
  readonly type?: "text" | "number" | "date";
  readonly order?: "ascending" | "descending";
}

export interface SortTableOptions {
  /** Up to three keys, as in Word's Sort dialog (Sort by / Then by / Then by). */
  readonly keys: readonly TableSortKey[];
  /** Keep the first row in place ("My list has a header row"). */
  readonly headerRow?: boolean;
  readonly caseSensitive?: boolean;
}

function cellTextOf(cell: WmlTableCell | undefined): string {
  if (!cell) return "";
  return cell.paragraphs
    .map((p) =>
      p.children
        .map((c) =>
          c.kind === "run"
            ? c.pieces.map((piece) => (piece.kind === "text" ? piece.value : piece.kind === "tab" ? "\t" : "")).join("")
            : "",
        )
        .join(""),
    )
    .join("\n");
}

/** Sort the table's rows by up to three columns, as Word's Sort does. */
export function sortTableRows(table: WmlTable, options: SortTableOptions): void {
  const columns = gridColumnCount(table);
  if (options.keys.length === 0) throw new RangeError("Give at least one sort key.");
  for (const key of options.keys) assertIndex("column", key.column, columns - 1);
  const first = options.headerRow ? 1 : 0;
  const placements = cellPlacements(table);
  if (placements.slice(first).some((row) => row.some((p) => p.vMerge))) {
    throw new RangeError("Tables with vertically merged cells cannot be sorted.");
  }
  const collator = new Intl.Collator(undefined, {
    sensitivity: options.caseSensitive ? "case" : "base",
    numeric: false,
  });
  const keyOf = (r: number, key: TableSortKey): string | number => {
    const p = placementAt(placements[r] ?? [], key.column);
    const text = cellTextOf(p ? table.rows[r]?.cells[p.cell] : undefined).trim();
    if (key.type === "number") return cellNumber(text) ?? Number.NEGATIVE_INFINITY;
    if (key.type === "date") {
      const time = Date.parse(text);
      return Number.isNaN(time) ? Number.NEGATIVE_INFINITY : time;
    }
    return text;
  };
  const indexes = table.rows.map((_, r) => r).slice(first);
  const keyed = indexes.map((r) => ({ r, keys: options.keys.map((k) => keyOf(r, k)) }));
  keyed.sort((a, b) => {
    for (let i = 0; i < options.keys.length; i++) {
      const x = a.keys[i]!;
      const y = b.keys[i]!;
      const cmp =
        typeof x === "number" && typeof y === "number"
          ? x - y
          : collator.compare(String(x), String(y));
      if (cmp !== 0) return options.keys[i]?.order === "descending" ? -cmp : cmp;
    }
    return a.r - b.r;
  });
  const sorted = keyed.map((k) => table.rows[k.r]!);
  table.rows.splice(first, sorted.length, ...sorted);
}

// --- Text ↔ table -------------------------------------------------------------

/** Split a paragraph's runs at a separator character, keeping run formatting. */
function splitParagraphAt(para: WmlParagraph, sep: string): WmlParagraph[] {
  const out: WmlParagraph[] = [];
  let current: WmlParagraph = emptyParagraphLike(para);
  const push = (): void => {
    out.push(current);
    current = emptyParagraphLike(para);
  };
  for (const child of para.children) {
    if (child.kind !== "run") {
      current.children.push(child);
      continue;
    }
    let pieces: WmlRunPiece[] = [];
    const flush = (): void => {
      if (pieces.length === 0) return;
      const run: WmlRun = {
        kind: "run",
        ...(child.rPr ? { rPr: structuredClone(child.rPr) } : {}),
        pieces,
        extras: [],
      };
      current.children.push(run);
      pieces = [];
    };
    for (const piece of child.pieces) {
      if (piece.kind === "tab" && sep === "\t") {
        flush();
        push();
      } else if (piece.kind === "text" && piece.value.includes(sep)) {
        const parts = piece.value.split(sep);
        parts.forEach((part, i) => {
          if (i > 0) {
            flush();
            push();
          }
          if (part) pieces.push({ kind: "text", value: part, preserveSpace: /^\s|\s$/.test(part) });
        });
      } else {
        pieces.push(piece);
      }
    }
    flush();
  }
  out.push(current);
  return out;
}

export interface ConvertTextToTableOptions {
  readonly separator: TableTextSeparator;
  /** Number of columns; defaults to the most fields in one paragraph (1 for `paragraph`). */
  readonly columns?: number;
  /** Total table width in twips, split evenly between the columns. */
  readonly widthTwips?: number;
}

// Word caps a converted table at 63 columns, like Insert Table.
const MAX_COLUMNS = 63;

/**
 * Turn body paragraphs `first`…`last` into a table (Word's Convert Text to
 * Table). With a character separator each paragraph becomes a row; with
 * `paragraph` each paragraph becomes a cell, filling rows of `columns`
 * cells. Run formatting is kept. Returns the new table, which replaces the
 * paragraphs in the body.
 */
export function convertTextToTable(
  doc: Docx,
  first: number,
  last: number,
  options: ConvertTextToTableOptions,
): WmlTable {
  const blocks = doc.document.body.blocks;
  assertIndex("first", first, blocks.length - 1);
  assertIndex("last", last, blocks.length - 1);
  if (last < first) throw new RangeError("last must not be before first.");
  const paragraphs = blocks.slice(first, last + 1).map((b) => {
    if (b.kind !== "paragraph") throw new RangeError("Only paragraphs can be converted to a table.");
    return b;
  });
  const sep = separatorChar(options.separator);
  let rows: WmlParagraph[][];
  if (sep === undefined) {
    const columns = options.columns ?? 1;
    assertIndex("columns", columns, MAX_COLUMNS);
    if (columns < 1) throw new RangeError("columns must be at least 1.");
    rows = [];
    for (let i = 0; i < paragraphs.length; i += columns) rows.push(paragraphs.slice(i, i + columns));
  } else {
    rows = paragraphs.map((p) => splitParagraphAt(p, sep));
  }
  const columns = Math.max(options.columns ?? 0, ...rows.map((r) => r.length));
  if (columns > MAX_COLUMNS) throw new RangeError(`A table can have at most ${MAX_COLUMNS} columns.`);
  const texts = rows.map(() => Array.from({ length: columns }, () => ""));
  const table = addTable(doc, texts, options.widthTwips ? { totalWidthTwips: options.widthTwips } : {});
  blocks.pop();
  table.rows.forEach((row, r) => {
    row.cells.forEach((cell, c) => {
      const para = rows[r]?.[c];
      if (para) cell.paragraphs = [para];
    });
  });
  blocks.splice(first, last - first + 1, table);
  doc.dirty = true;
  return table;
}

// --- Formulas -----------------------------------------------------------------

/**
 * Evaluate a table formula (the `=` field: `SUM(ABOVE)`, `AVERAGE(B2:B5)`,
 * `A1*2` …) for the cell at `row` / `cell`, reading the numbers the table's
 * cells hold. Throws a SyntaxError on a malformed formula.
 */
export function evaluateTableFormula(
  table: WmlTable,
  row: number,
  cell: number,
  formula: string,
): number {
  const placements = cellPlacements(table);
  const at = placements[row]?.[cell];
  if (!at) throw new RangeError(`No cell ${cell} in row ${row}.`);
  return evaluateFormula(formula, {
    rows: table.rows.length,
    columns: gridColumnCount(table),
    at: { row, column: at.gridStart },
    value: (r, c) => {
      const p = placementAt(placements[r] ?? [], c);
      if (!p || p.rowSpan === 0) return undefined;
      return cellNumber(cellTextOf(table.rows[r]?.cells[p.cell]));
    },
  });
}

/**
 * Format a number with a field numeric picture (`\# "#,##0.00"`, §17.16.4.2);
 * without a picture, Word's general format.
 */
export function formatFieldNumber(value: number, picture?: string): string {
  return formatNumber(value, picture);
}

// --- Cell content -------------------------------------------------------------

/** A block inside a table cell, in document order. */
export type TableCellBlock =
  | WmlParagraph
  | WmlTable
  | { readonly kind: "raw"; readonly node: XmlElement };

/**
 * A cell's blocks in document order, nested tables included. The cell's
 * paragraphs are the live objects; a nested table is parsed from the cell's
 * XML, so it is a read-only snapshot (edits to it are not saved).
 */
export function tableCellBlocks(cell: WmlTableCell): TableCellBlock[] {
  return cellNodes(cell).flatMap((n): TableCellBlock[] => {
    if (n.kind === "paragraph") return [n.paragraph];
    if (n.node.kind !== "element") return [];
    if (n.node.name.local === "tbl") return [parseTableElement(n.node)];
    return [{ kind: "raw", node: n.node }];
  });
}
