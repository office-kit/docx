/**
 * Table commands — Insert ▸ Table, and the Table Design / Table Layout tabs.
 * Commands act on the cells the selection covers (see ../table-selection.ts):
 * a caret is its cell, a selection across cells the rectangle between them.
 */

import {
  addBuiltInTableStyle,
  addStyle,
  addTable,
  appendField,
  appendTableRow,
  autoFitTable,
  type CellMargins,
  convertTextToTable,
  deleteTableCells,
  deleteTableColumns,
  deleteTableRows,
  distributeTableColumns,
  evaluateTableFormula,
  formatFieldNumber,
  getTableLook,
  getTableStyle,
  insertTableColumn,
  insertTableRow,
  listStyles,
  mergeTableCells,
  type SortTableOptions,
  setParagraphAlignment,
  setTableAlignment,
  setTableAltText,
  setTableBorders,
  setTableCellMargins,
  setTableCellShading,
  setTableCellSpacing,
  setTableCellText,
  setTableCellTextDirection,
  setTableCellVerticalAlign,
  setTableCellWidth,
  setTableColumnWidth,
  setTableDefaultCellMargins,
  setTableIndent,
  setTableLayout,
  setTableLook,
  setTablePosition,
  setTableRangeBorders,
  setTableRowAsHeader,
  setTableRowHeight,
  setTableStyle,
  setTableStyleFormatting,
  setTableWidth,
  sortTableRows,
  splitTable,
  splitTableCell,
  type TableAlignment,
  type TableBorder,
  type TableBorderEdges,
  type TableBordersOptions,
  type TableCellRange,
  type TableCellShadingOptions,
  type TableCellTextDirection,
  type TableCellVerticalAlign,
  type TableLook,
  type TablePosition,
  type TableRowHeightRule,
  type TableStyleFormatting,
  type TableStyleRegion,
  type TableTextSeparator,
  type TableWidth,
  tableColumnCount,
  unwrapTable,
  type WmlTable,
  setElementOnOff,
  makePropsElement,
  BUILT_IN_TABLE_STYLES,
} from "@office-kit/docx";
import { asTable, blockAt } from "../doc-access.js";
import type { EditorModel } from "../model.js";
import { pageGeometry } from "../resolve.js";
import { caretAt, orderSelection } from "../selection.js";
import {
  cellIndexAt,
  cellStart,
  selectedPlacements,
  type TableSelection,
  tableSelection,
} from "../table-selection.js";
import { caretBlockIndex, moveLastBlockAfter } from "./insert-util.js";
import type { Command } from "./types.js";

/** The table the caret is inside, or undefined. */
function currentTable(model: EditorModel): WmlTable | undefined {
  const block = model.selection?.focus.block;
  if (block === undefined) return undefined;
  return asTable(blockAt(model.doc, block));
}

const inTable = (model: EditorModel): boolean => !!tableSelection(model);

function requireSelection(model: EditorModel): TableSelection {
  const ts = tableSelection(model);
  if (!ts) throw new Error("Put the insertion point in a table first.");
  return ts;
}

function cellsOf(ts: TableSelection) {
  return selectedPlacements(ts).flatMap((p) => {
    const cell = ts.table.rows[p.row]?.cells[p.cell];
    return cell ? [cell] : [];
  });
}

function rowsOf(ts: TableSelection) {
  return ts.table.rows.slice(ts.range.firstRow, ts.range.lastRow + 1);
}

// Word's own Insert Table dialog caps a new table at 63 columns; rows are
// uncapped there, but an unbounded count from a typo would freeze the page.
const MAX_TABLE_COLS = 63;
const MAX_TABLE_ROWS = 1000;
// Word applies its Table Grid style to a table inserted from the ribbon.
export const DEFAULT_TABLE_STYLE = "TableGrid";

function assertGridSize(name: string, value: number, max: number): void {
  if (!Number.isInteger(value) || value < 1 || value > max) {
    throw new RangeError(`Table ${name} must be an integer from 1 to ${max}, got ${value}.`);
  }
}

/** Apply a table style, adding a built-in style's definition the first time it is used. */
function applyStyle(model: EditorModel, table: WmlTable, styleId: string | undefined): void {
  if (styleId !== undefined && BUILT_IN_TABLE_STYLES.some((s) => s.styleId === styleId)) {
    addBuiltInTableStyle(model.doc, styleId);
  }
  setTableStyle(table, styleId);
}

export interface InsertTableParams {
  readonly rows: number;
  readonly cols: number;
  /** Word's Insert Table ▸ AutoFit behavior; `fixed` (the default) keeps equal columns. */
  readonly autoFit?: "fixed" | "contents" | "window";
  /** Fixed column width in twips; by default the columns share the text width. */
  readonly columnWidthTwips?: number;
  /** Table style; defaults to Table Grid, `null` for none. */
  readonly styleId?: string | null;
  /** Initial cell text (Quick Tables), row-major. */
  readonly cells?: ReadonlyArray<ReadonlyArray<string>>;
}

export const insertTableCommand: Command<InsertTableParams> = {
  id: "table.insert",
  group: "table",
  label: "Insert table",
  run(model, { rows, cols, autoFit = "fixed", columnWidthTwips, styleId, cells }) {
    assertGridSize("rows", rows, MAX_TABLE_ROWS);
    assertGridSize("cols", cols, MAX_TABLE_COLS);
    const at = caretBlockIndex(model.doc, model.selection?.focus.block);
    const page = pageGeometry(model.doc);
    const textWidth = page.width - page.left - page.right;
    const width = columnWidthTwips === undefined ? textWidth : columnWidthTwips * cols;
    const grid: string[][] = Array.from({ length: rows }, (_, r) =>
      Array.from({ length: cols }, (_, c) => cells?.[r]?.[c] ?? ""),
    );
    const table = addTable(model.doc, grid, { totalWidthTwips: width });
    // Word's Insert Table leaves the table's width to its columns.
    setTableWidth(table, { type: "auto", value: 0 });
    applyStyle(model, table, styleId === undefined ? DEFAULT_TABLE_STYLE : (styleId ?? undefined));
    if (autoFit !== "fixed") autoFitTable(table, autoFit);
    else if (columnWidthTwips !== undefined) setTableLayout(table, "fixed");
    moveLastBlockAfter(model.doc, at);
    model.setSelection(caretAt(cellStart(at + 1, 0, 0)));
  },
};

export const addRowCommand: Command<{ texts?: readonly string[] }> = {
  id: "table.addRow",
  group: "table",
  label: "Add row",
  run(model, { texts = [] }) {
    const table = currentTable(model);
    if (!table) return;
    appendTableRow(table, texts);
  },
  isEnabled: (model) => !!currentTable(model),
};

/** Insert Above / Below: as many rows as are selected, like Word. */
export const insertRowsCommand: Command<{ where: "above" | "below" }> = {
  id: "table.insertRows",
  group: "table",
  label: "Insert rows",
  run(model, { where }) {
    const ts = requireSelection(model);
    const count = ts.range.lastRow - ts.range.firstRow + 1;
    const at = where === "above" ? ts.range.firstRow : ts.range.lastRow + 1;
    const formatFrom = where === "above" ? ts.range.firstRow : ts.range.lastRow;
    for (let i = 0; i < count; i++)
      insertTableRow(ts.table, at, { formatFrom: formatFrom + (where === "above" ? i : 0) });
    const placements = tableSelectionAfter(model, ts);
    model.setSelection(
      caretAt(cellStart(ts.block, at, cellIndexAt(placements, at, ts.range.firstColumn))),
    );
  },
  isEnabled: inTable,
};

function tableSelectionAfter(model: EditorModel, ts: TableSelection) {
  return tableSelection(model)?.placements ?? ts.placements;
}

/** Insert Left / Right: as many columns as are selected. */
export const insertColumnsCommand: Command<{ where: "left" | "right" }> = {
  id: "table.insertColumns",
  group: "table",
  label: "Insert columns",
  run(model, { where }) {
    const ts = requireSelection(model);
    const count = ts.range.lastColumn - ts.range.firstColumn + 1;
    const at = where === "left" ? ts.range.firstColumn : ts.range.lastColumn + 1;
    const formatFrom = where === "left" ? ts.range.firstColumn : ts.range.lastColumn;
    for (let i = 0; i < count; i++) {
      // Each new column to the left pushes the column it copies one step right.
      insertTableColumn(ts.table, at, {
        formatFrom: where === "left" ? formatFrom + i : formatFrom,
      });
    }
    const placements = tableSelectionAfter(model, ts);
    model.setSelection(
      caretAt(
        cellStart(ts.block, ts.range.firstRow, cellIndexAt(placements, ts.range.firstRow, at)),
      ),
    );
  },
  isEnabled: inTable,
};

function removeTableBlock(model: EditorModel, block: number): void {
  const blocks = model.doc.document.body.blocks;
  blocks.splice(block, 1);
  // A body must keep a paragraph to put the caret in.
  if (blocks.length === 0) blocks.push({ kind: "paragraph", children: [], extras: [] });
  const at = Math.min(block, blocks.length - 1);
  model.setSelection(
    blocks[at]?.kind === "paragraph" ? caretAt({ block: at, inline: 0, offset: 0 }) : null,
  );
}

/** Delete Rows: the selected rows (the whole table when every row goes). */
export const deleteRowCommand: Command<{ row?: number }> = {
  id: "table.deleteRow",
  group: "table",
  label: "Delete rows",
  run(model, { row }) {
    const ts = requireSelection(model);
    const first = row ?? ts.range.firstRow;
    const last = row ?? ts.range.lastRow;
    if (last - first + 1 >= ts.table.rows.length) {
      removeTableBlock(model, ts.block);
      return;
    }
    deleteTableRows(ts.table, first, last);
    const target = Math.min(first, ts.table.rows.length - 1);
    model.setSelection(caretAt(cellStart(ts.block, target, 0)));
  },
  isEnabled: inTable,
};

export const deleteColumnsCommand: Command<void> = {
  id: "table.deleteColumns",
  group: "table",
  label: "Delete columns",
  run(model) {
    const ts = requireSelection(model);
    if (!deleteTableColumns(ts.table, ts.range.firstColumn, ts.range.lastColumn)) {
      removeTableBlock(model, ts.block);
      return;
    }
    model.setSelection(caretAt(cellStart(ts.block, ts.range.firstRow, 0)));
  },
  isEnabled: inTable,
};

export const deleteTableCommand: Command<void> = {
  id: "table.deleteTable",
  group: "table",
  label: "Delete table",
  run(model) {
    removeTableBlock(model, requireSelection(model).block);
  },
  isEnabled: inTable,
};

/** Word's Delete Cells dialog: shift cells left / up, or delete entire rows / columns. */
export const deleteCellsCommand: Command<{ shift: "left" | "up" | "row" | "column" }> = {
  id: "table.deleteCells",
  group: "table",
  label: "Delete cells",
  run(model, { shift }) {
    if (shift === "row") {
      deleteRowCommand.run(model, {});
      return;
    }
    if (shift === "column") {
      deleteColumnsCommand.run(model);
      return;
    }
    const ts = requireSelection(model);
    deleteTableCells(ts.table, ts.range, shift);
    if (ts.table.rows.length === 0) {
      removeTableBlock(model, ts.block);
      return;
    }
    model.setSelection(
      caretAt(cellStart(ts.block, Math.min(ts.range.firstRow, ts.table.rows.length - 1), 0)),
    );
  },
  isEnabled: inTable,
};

/** Merge Cells: the selection, or (for the Eraser) an explicit range. */
export const mergeCellsCommand: Command<{ range?: TableCellRange }> = {
  id: "table.mergeCells",
  group: "table",
  label: "Merge cells",
  run(model, { range }) {
    const ts = requireSelection(model);
    const target = range ?? ts.range;
    mergeTableCells(ts.table, target);
    const placements = tableSelectionAfter(model, ts);
    model.setSelection(
      caretAt(
        cellStart(
          ts.block,
          target.firstRow,
          cellIndexAt(placements, target.firstRow, target.firstColumn),
        ),
      ),
    );
  },
  isEnabled: (model) => !!tableSelection(model)?.multiCell,
};

/** Split Cells: the focus cell, or (for Draw Table) an explicit one. */
export const splitCellsCommand: Command<{
  columns: number;
  rows: number;
  row?: number;
  cell?: number;
}> = {
  id: "table.splitCells",
  group: "table",
  label: "Split cells",
  run(model, { columns, rows, row, cell }) {
    const ts = requireSelection(model);
    const r = row ?? ts.focus.row;
    const c = cell ?? ts.focus.cell;
    splitTableCell(ts.table, r, c, { columns, rows });
    model.setSelection(caretAt(cellStart(ts.block, r, c)));
  },
  isEnabled: inTable,
};

export const splitTableCommand: Command<void> = {
  id: "table.splitTable",
  group: "table",
  label: "Split table",
  run(model) {
    const ts = requireSelection(model);
    const row = ts.focus.row;
    splitTable(model.doc, ts.table, row);
    const block = row === 0 ? ts.block + 1 : ts.block + 2;
    model.setSelection(caretAt(cellStart(block, 0, ts.focus.cell)));
  },
  isEnabled: inTable,
};

// --- Table Design ---------------------------------------------------------------

export const setTableStyleCommand: Command<{ styleId: string | undefined }> = {
  id: "table.setStyle",
  group: "table",
  label: "Table style",
  run(model, { styleId }) {
    applyStyle(model, requireSelection(model).table, styleId);
  },
  isEnabled: inTable,
};

/** The table style's id at the selection, for the gallery's highlight. */
export function currentTableStyle(model: EditorModel): string | undefined {
  const table = tableSelection(model)?.table;
  return table && getTableStyle(table);
}

/** The Table Style Options at the selection. */
export function currentTableLook(model: EditorModel): TableLook | undefined {
  const table = tableSelection(model)?.table;
  return table && getTableLook(table);
}

export const setTableLookCommand: Command<Partial<TableLook>> = {
  id: "table.setLook",
  group: "table",
  label: "Table style options",
  run(model, change) {
    const table = requireSelection(model).table;
    setTableLook(table, { ...getTableLook(table), ...change });
  },
  isEnabled: inTable,
};

/** New Table Style: define a table style (based on another) and apply it. */
export const newTableStyleCommand: Command<{ name: string; basedOn?: string }, string> = {
  id: "table.newStyle",
  group: "table",
  label: "New table style",
  run(model, { name, basedOn = DEFAULT_TABLE_STYLE }) {
    const trimmed = name.trim();
    if (!trimmed) throw new RangeError("Give the style a name.");
    const base = trimmed.replace(/[^A-Za-z0-9]/g, "") || "TableStyle";
    const taken = new Set(listStyles(model.doc).map((s) => s.styleId));
    let styleId = base;
    for (let n = 2; taken.has(styleId); n++) styleId = `${base}${n}`;
    if (BUILT_IN_TABLE_STYLES.some((s) => s.styleId === basedOn))
      addBuiltInTableStyle(model.doc, basedOn);
    addStyle(model.doc, { type: "table", styleId, name: trimmed, basedOn, customStyle: true });
    const table = tableSelection(model)?.table;
    if (table) setTableStyle(table, styleId);
    return styleId;
  },
};

/** Modify Table Style: change one region's formatting of a style. */
export const modifyTableStyleCommand: Command<{
  styleId: string;
  region: TableStyleRegion;
  formatting: TableStyleFormatting;
}> = {
  id: "table.modifyStyle",
  group: "table",
  label: "Modify table style",
  run(model, { styleId, region, formatting }) {
    if (BUILT_IN_TABLE_STYLES.some((s) => s.styleId === styleId))
      addBuiltInTableStyle(model.doc, styleId);
    setTableStyleFormatting(model.doc, styleId, region, formatting);
  },
};

/** Shading: fill every selected cell (`auto` removes the fill). */
export const shadeCellsCommand: Command<{ fill: string }> = {
  id: "table.shading",
  group: "table",
  label: "Shading",
  run(model, { fill }) {
    for (const cell of cellsOf(requireSelection(model))) setTableCellShading(cell, { fill });
  },
  isEnabled: inTable,
};

/** Borders menu and Border Painter: a border on edges of the selection (or a given range). */
export const rangeBordersCommand: Command<{
  edges: TableBorderEdges;
  border?: TableBorder;
  range?: TableCellRange;
}> = {
  id: "table.rangeBorders",
  group: "table",
  label: "Borders",
  run(model, { edges, border, range }) {
    const ts = requireSelection(model);
    setTableRangeBorders(ts.table, range ?? ts.range, edges, border);
  },
  isEnabled: inTable,
};

/** A cell edge the Borders and Shading dialog sets or clears. */
export type TableBorderSide = (typeof TABLE_BORDER_SIDES)[number];
const TABLE_BORDER_SIDES = [
  "top",
  "bottom",
  "left",
  "right",
  "insideH",
  "insideV",
  "tl2br",
  "tr2bl",
] as const;

/**
 * The Borders and Shading dialog: every edge it lists (a border, or `null`
 * for none) and optionally the cells' fill, applied as one undo step.
 */
export const bordersAndShadingCommand: Command<{
  range: TableCellRange;
  borders: Partial<Record<TableBorderSide, TableBorder | null>>;
  fill?: string;
}> = {
  id: "table.bordersShading",
  group: "table",
  label: "Borders and shading",
  run(model, { range, borders, fill }) {
    const ts = requireSelection(model);
    for (const side of TABLE_BORDER_SIDES) {
      const border = borders[side];
      if (border !== undefined) setTableRangeBorders(ts.table, range, side, border ?? undefined);
    }
    if (fill === undefined) return;
    for (const p of ts.placements.flat()) {
      if (p.rowSpan === 0) continue;
      const inRange =
        p.row >= range.firstRow &&
        p.row <= range.lastRow &&
        p.gridStart >= range.firstColumn &&
        p.gridStart + p.gridSpan - 1 <= range.lastColumn;
      const cell = ts.table.rows[p.row]?.cells[p.cell];
      if (inRange && cell) setTableCellShading(cell, { fill });
    }
  },
  isEnabled: inTable,
};

// --- Table Layout ---------------------------------------------------------------

/** Everything the Table Properties dialog sets, applied as one undo step. */
export interface TablePropertiesParams {
  readonly table?: {
    readonly width?: TableWidth | null;
    readonly alignment?: TableAlignment;
    readonly indentTwips?: number;
    /** `null` puts a floating table back in line with the text. */
    readonly position?: TablePosition | null;
  };
  readonly rows?: {
    readonly height?: { readonly twips: number; readonly rule: TableRowHeightRule } | null;
    readonly cantSplit?: boolean;
    readonly header?: boolean;
  };
  readonly columnWidthTwips?: number;
  readonly cells?: {
    readonly width?: TableWidth | null;
    readonly verticalAlign?: TableCellVerticalAlign;
    /** `null` goes back to the table's default margins. */
    readonly margins?: CellMargins | null;
    readonly noWrap?: boolean;
    readonly fitText?: boolean;
  };
  readonly options?: {
    readonly defaultMargins?: CellMargins;
    readonly cellSpacingTwips?: number;
    /** "Automatically resize to fit contents". */
    readonly autoResize?: boolean;
  };
  readonly altText?: { readonly title?: string; readonly description?: string };
}

export const tablePropertiesCommand: Command<TablePropertiesParams> = {
  id: "table.properties",
  group: "table",
  label: "Table properties",
  run(model, params) {
    const ts = requireSelection(model);
    const { table } = ts;
    const t = params.table;
    if (t?.width !== undefined) setTableWidth(table, t.width ?? undefined);
    if (t?.alignment !== undefined)
      setTableAlignment(table, t.alignment === "left" ? undefined : t.alignment);
    if (t?.indentTwips !== undefined) setTableIndent(table, t.indentTwips);
    if (t?.position !== undefined) setTablePosition(table, t.position ?? undefined);
    const r = params.rows;
    for (const row of r ? rowsOf(ts) : []) {
      if (r?.height !== undefined) {
        if (r.height === null) setTableRowHeight(row, 0, "auto");
        else setTableRowHeight(row, r.height.twips, r.height.rule);
      }
      if (r?.cantSplit !== undefined) {
        row.trPr ??= makePropsElement("trPr");
        setElementOnOff(row.trPr, "cantSplit", r.cantSplit);
      }
      if (r?.header !== undefined) setTableRowAsHeader(row, r.header);
    }
    if (params.columnWidthTwips !== undefined) {
      for (let g = ts.range.firstColumn; g <= ts.range.lastColumn; g++) {
        setTableColumnWidth(table, g, params.columnWidthTwips);
      }
    }
    const c = params.cells;
    for (const cell of c ? cellsOf(ts) : []) {
      if (c?.width !== undefined) setTableCellWidth(cell, c.width ?? undefined);
      if (c?.verticalAlign !== undefined) setTableCellVerticalAlign(cell, c.verticalAlign);
      if (c?.margins !== undefined) setTableCellMargins(cell, c.margins ?? undefined);
      cell.tcPr ??= makePropsElement("tcPr");
      if (c?.noWrap !== undefined) setElementOnOff(cell.tcPr, "noWrap", c.noWrap);
      if (c?.fitText !== undefined) setElementOnOff(cell.tcPr, "tcFitText", c.fitText);
    }
    const o = params.options;
    if (o?.defaultMargins) setTableDefaultCellMargins(table, o.defaultMargins);
    if (o?.cellSpacingTwips !== undefined) setTableCellSpacing(table, o.cellSpacingTwips);
    if (o?.autoResize !== undefined) setTableLayout(table, o.autoResize ? "autofit" : "fixed");
    if (params.altText) setTableAltText(table, params.altText);
  },
  isEnabled: inTable,
};

/** Table Layout ▸ Height: the selected rows' minimum height. */
export const rowHeightSelectionCommand: Command<{ twips: number }> = {
  id: "table.rowHeightSel",
  group: "table",
  label: "Row height",
  run(model, { twips }) {
    if (!Number.isFinite(twips) || twips < 0) throw new RangeError("Height must be ≥ 0.");
    for (const row of rowsOf(requireSelection(model))) {
      setTableRowHeight(row, Math.round(twips), twips === 0 ? "auto" : "atLeast");
    }
  },
  isEnabled: inTable,
};

/** Table Layout ▸ Width: the selected columns' width. */
export const columnWidthSelectionCommand: Command<{ twips: number }> = {
  id: "table.columnWidthSel",
  group: "table",
  label: "Column width",
  run(model, { twips }) {
    const ts = requireSelection(model);
    for (let g = ts.range.firstColumn; g <= ts.range.lastColumn; g++)
      setTableColumnWidth(ts.table, g, twips);
  },
  isEnabled: inTable,
};

/** Distribute Rows: the selected rows get the same (measured) height. */
export const distributeRowsCommand: Command<{ heightTwips: number }> = {
  id: "table.distributeRows",
  group: "table",
  label: "Distribute rows",
  run(model, { heightTwips }) {
    for (const row of rowsOf(requireSelection(model)))
      setTableRowHeight(row, Math.round(heightTwips), "atLeast");
  },
  isEnabled: inTable,
};

/** Distribute Columns: the selected columns (all, for one cell) share their width equally. */
export const distributeColumnsCommand: Command<void> = {
  id: "table.distributeColumns",
  group: "table",
  label: "Distribute columns",
  run(model) {
    const ts = requireSelection(model);
    const single = ts.range.firstColumn === ts.range.lastColumn;
    distributeTableColumns(
      ts.table,
      single ? 0 : ts.range.firstColumn,
      single ? tableColumnCount(ts.table) - 1 : ts.range.lastColumn,
    );
  },
  isEnabled: inTable,
};

export const autoFitCommand: Command<{ mode: "contents" | "window" | "fixed" }> = {
  id: "table.autoFit",
  group: "table",
  label: "AutoFit",
  run(model, { mode }) {
    autoFitTable(requireSelection(model).table, mode);
  },
  isEnabled: inTable,
};

/** The Alignment grid: paragraph alignment and cell vertical alignment of the selected cells. */
export const cellAlignmentCommand: Command<{
  horizontal: "left" | "center" | "right";
  vertical: TableCellVerticalAlign;
}> = {
  id: "table.cellAlign",
  group: "table",
  label: "Cell alignment",
  run(model, { horizontal, vertical }) {
    for (const cell of cellsOf(requireSelection(model))) {
      setTableCellVerticalAlign(cell, vertical);
      for (const p of cell.paragraphs) setParagraphAlignment(p, horizontal);
    }
  },
  isEnabled: inTable,
};

export const cellTextDirectionCommand: Command<{ direction: TableCellTextDirection }> = {
  id: "table.textDirection",
  group: "table",
  label: "Text direction",
  run(model, { direction }) {
    for (const cell of cellsOf(requireSelection(model))) setTableCellTextDirection(cell, direction);
  },
  isEnabled: inTable,
};

/** Repeat Header Rows: toggle the selected rows (from the top) as repeated header rows. */
export const repeatHeaderRowsCommand: Command<void> = {
  id: "table.repeatHeader",
  group: "table",
  label: "Repeat header rows",
  run(model) {
    const ts = requireSelection(model);
    const on = !isHeaderRow(ts);
    // Word marks every row from the first through the selection.
    for (const row of ts.table.rows.slice(0, ts.range.lastRow + 1)) setTableRowAsHeader(row, on);
  },
  isEnabled: inTable,
  isActive: (model) => {
    const ts = tableSelection(model);
    return !!ts && isHeaderRow(ts);
  },
};

function isHeaderRow(ts: TableSelection): boolean {
  const trPr = ts.table.rows[ts.range.firstRow]?.trPr;
  return !!trPr?.children.some((c) => c.kind === "element" && c.name.local === "tblHeader");
}

export const sortTableCommand: Command<SortTableOptions> = {
  id: "table.sort",
  group: "table",
  label: "Sort",
  run(model, options) {
    sortTableRows(requireSelection(model).table, options);
  },
  isEnabled: inTable,
};

function tableIndexOf(model: EditorModel, block: number): number {
  return model.doc.document.body.blocks.slice(0, block).filter((b) => b.kind === "table").length;
}

export const convertTableToTextCommand: Command<{ separator: TableTextSeparator }> = {
  id: "table.convertToText",
  group: "table",
  label: "Convert to text",
  run(model, { separator }) {
    const ts = requireSelection(model);
    unwrapTable(model.doc, tableIndexOf(model, ts.block), { separator });
    model.setSelection(caretAt({ block: ts.block, inline: 0, offset: 0 }));
  },
  isEnabled: inTable,
};

export const convertTextToTableCommand: Command<{
  separator: TableTextSeparator;
  columns?: number;
}> = {
  id: "table.convertTextToTable",
  group: "table",
  label: "Convert text to table",
  run(model, { separator, columns }) {
    const sel = model.selection;
    if (!sel) throw new Error("Select the paragraphs to convert.");
    const { start, end } = orderSelection(sel);
    const page = pageGeometry(model.doc);
    const table = convertTextToTable(model.doc, start.block, end.block, {
      separator,
      ...(columns === undefined ? {} : { columns }),
      widthTwips: page.width - page.left - page.right,
    });
    setTableWidth(table, { type: "auto", value: 0 });
    applyStyle(model, table, DEFAULT_TABLE_STYLE);
    model.setSelection(caretAt(cellStart(start.block, 0, 0)));
  },
  isEnabled: (model) => {
    const sel = model.selection;
    if (!sel || sel.focus.cell || sel.anchor.cell) return false;
    const { start, end } = orderSelection(sel);
    return model.doc.document.body.blocks
      .slice(start.block, end.block + 1)
      .every((b) => b.kind === "paragraph");
  },
};

/** Formula: insert an `=` field with its computed result at the end of the caret's paragraph. */
export const insertFormulaCommand: Command<{ formula: string; numberFormat?: string }, number> = {
  id: "table.formula",
  group: "table",
  label: "Formula",
  run(model, { formula, numberFormat }) {
    const ts = requireSelection(model);
    const expression = formula.trim().replace(/^=/, "");
    const value = evaluateTableFormula(ts.table, ts.focus.row, ts.focus.cell, expression);
    const cell = ts.table.rows[ts.focus.row]?.cells[ts.focus.cell];
    const para = cell?.paragraphs[model.selection?.focus.para ?? 0] ?? cell?.paragraphs[0];
    if (!para) throw new Error("The cell has no paragraph.");
    const picture = numberFormat?.trim();
    // A picture with spaces or quotes must itself be quoted (§17.16.4.2).
    const instruction = picture
      ? `=${expression} \\# "${picture.replace(/"/g, "")}"`
      : `=${expression}`;
    appendField(model.doc, para, instruction, formatFieldNumber(value, picture || undefined));
    return value;
  },
  isEnabled: inTable,
};

// --- Direct cell / row / table formatting kept from the generic table API ---------------

export const setCellTextCommand: Command<{ row: number; col: number; text: string }> = {
  id: "table.setCellText",
  group: "table",
  label: "Set cell text",
  run(model, { row, col, text }) {
    const table = currentTable(model);
    if (!table) return;
    setTableCellText(table, row, col, text);
  },
  isEnabled: (model) => !!currentTable(model),
};

export const setCellVerticalAlignCommand: Command<{
  row: number;
  col: number;
  align: TableCellVerticalAlign;
}> = {
  id: "table.cellVAlign",
  group: "table",
  label: "Cell vertical alignment",
  run(model, { row, col, align }) {
    const cell = currentTable(model)?.rows[row]?.cells[col];
    if (cell) setTableCellVerticalAlign(cell, align);
  },
  isEnabled: (model) => !!currentTable(model),
};

export const setRowHeaderCommand: Command<{ row: number; isHeader?: boolean }> = {
  id: "table.rowHeader",
  group: "table",
  label: "Header row",
  run(model, { row, isHeader = true }) {
    const r = currentTable(model)?.rows[row];
    if (r) setTableRowAsHeader(r, isHeader);
  },
  isEnabled: (model) => !!currentTable(model),
};

export const setRowHeightCommand: Command<{
  row: number;
  heightTwips: number;
  rule?: TableRowHeightRule;
}> = {
  id: "table.rowHeight",
  group: "table",
  label: "Row height",
  run(model, { row, heightTwips, rule = "atLeast" }) {
    const r = currentTable(model)?.rows[row];
    if (r) setTableRowHeight(r, heightTwips, rule);
  },
  isEnabled: (model) => !!currentTable(model),
};

export const setTableBordersCommand: Command<TableBordersOptions> = {
  id: "table.borders",
  group: "table",
  label: "Table borders",
  run(model, options) {
    const table = currentTable(model);
    if (table) setTableBorders(table, options);
  },
  isEnabled: (model) => !!currentTable(model),
};

export const setCellShadingCommand: Command<{
  row: number;
  col: number;
  shading: TableCellShadingOptions;
}> = {
  id: "table.cellShading",
  group: "table",
  label: "Cell shading",
  run(model, { row, col, shading }) {
    const cell = currentTable(model)?.rows[row]?.cells[col];
    if (cell) setTableCellShading(cell, shading);
  },
  isEnabled: (model) => !!currentTable(model),
};

export const tableCommands = [
  insertTableCommand,
  addRowCommand,
  insertRowsCommand,
  insertColumnsCommand,
  deleteRowCommand,
  deleteColumnsCommand,
  deleteTableCommand,
  deleteCellsCommand,
  mergeCellsCommand,
  splitCellsCommand,
  splitTableCommand,
  setTableStyleCommand,
  setTableLookCommand,
  newTableStyleCommand,
  modifyTableStyleCommand,
  shadeCellsCommand,
  rangeBordersCommand,
  bordersAndShadingCommand,
  tablePropertiesCommand,
  rowHeightSelectionCommand,
  columnWidthSelectionCommand,
  distributeRowsCommand,
  distributeColumnsCommand,
  autoFitCommand,
  cellAlignmentCommand,
  cellTextDirectionCommand,
  repeatHeaderRowsCommand,
  sortTableCommand,
  convertTableToTextCommand,
  convertTextToTableCommand,
  insertFormulaCommand,
  setCellTextCommand,
  setCellVerticalAlignCommand,
  setRowHeaderCommand,
  setRowHeightCommand,
  setTableBordersCommand,
  setCellShadingCommand,
];
