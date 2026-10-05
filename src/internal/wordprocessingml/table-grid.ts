/**
 * The table grid model (ECMA-376 §17.4.49 tblGrid, §17.4.17 gridSpan,
 * §17.4.85 vMerge): where each `<w:tc>` sits on the shared column grid and how
 * many rows a vertically merged cell covers. Every structural table edit
 * (insert/delete columns, merge, split) works on this model, because a cell's
 * index within its row says nothing about its column once cells span.
 */

import type { XmlElement, XmlNode } from "../xml/index.js";
import {
  TC_PR_ORDER,
  TR_PR_ORDER,
  removeWChild,
  upsertWChild,
  wAttrOf,
  wChild,
  wChildren,
  wEl,
} from "./table-xml.js";
import type { WmlParagraph, WmlRunPiece, WmlTable, WmlTableCell, WmlTableRow } from "./types.js";

/** Where one `<w:tc>` sits on the table grid. */
export interface TableCellPlacement {
  /** Row index in `table.rows`. */
  readonly row: number;
  /** Cell index in `table.rows[row].cells`. */
  readonly cell: number;
  /** First grid column the cell covers (after the row's `gridBefore`). */
  readonly gridStart: number;
  /** Grid columns the cell covers (`w:gridSpan`, default 1). */
  readonly gridSpan: number;
  /** `w:vMerge`: `"restart"` starts a vertical merge, `"continue"` extends the one above. */
  readonly vMerge?: "restart" | "continue" | undefined;
  /**
   * Rows the cell covers on screen: 1 for a plain cell, the merge height for
   * the first cell of a vertical merge, and 0 for a cell that continues one.
   */
  readonly rowSpan: number;
}

type MutablePlacement = { -readonly [K in keyof TableCellPlacement]: TableCellPlacement[K] };

export function intAttr(el: XmlElement | undefined, local: string): number | undefined {
  const raw = wAttrOf(el, local);
  if (raw === undefined) return undefined;
  const n = Number(raw);
  return Number.isFinite(n) ? Math.trunc(n) : undefined;
}

export function gridSpanOf(cell: WmlTableCell): number {
  const span = intAttr(wChild(cell.tcPr, "gridSpan"), "val");
  return span !== undefined && span > 0 ? span : 1;
}

export function vMergeOf(cell: WmlTableCell): "restart" | "continue" | undefined {
  const el = wChild(cell.tcPr, "vMerge");
  if (!el) return undefined;
  // A bare <w:vMerge/> continues the merge above (§17.4.85).
  return wAttrOf(el, "val") === "restart" ? "restart" : "continue";
}

export function gridBeforeOf(row: WmlTableRow): number {
  return Math.max(0, intAttr(wChild(row.trPr, "gridBefore"), "val") ?? 0);
}

export function gridAfterOf(row: WmlTableRow): number {
  return Math.max(0, intAttr(wChild(row.trPr, "gridAfter"), "val") ?? 0);
}

/** Every cell's grid position, row by row. */
export function cellPlacements(table: WmlTable): TableCellPlacement[][] {
  const out: MutablePlacement[][] = [];
  // Grid column → the cell whose vertical merge is still open there.
  let open = new Map<number, MutablePlacement>();
  table.rows.forEach((row, r) => {
    let g = gridBeforeOf(row);
    const next = new Map<number, MutablePlacement>();
    const placed = row.cells.map((cell, c) => {
      const p: MutablePlacement = {
        row: r,
        cell: c,
        gridStart: g,
        gridSpan: gridSpanOf(cell),
        vMerge: vMergeOf(cell),
        rowSpan: 1,
      };
      g += p.gridSpan;
      if (p.vMerge === "continue") {
        const owner = open.get(p.gridStart);
        if (owner && owner.gridSpan === p.gridSpan) {
          owner.rowSpan++;
          p.rowSpan = 0;
          next.set(p.gridStart, owner);
          return p;
        }
        // A continuation with nothing above to continue: Word shows it as a
        // cell that starts its own merge.
      }
      if (p.vMerge) next.set(p.gridStart, p);
      return p;
    });
    open = next;
    out.push(placed);
  });
  return out;
}

/** The grid column count: `tblGrid`, or the widest row when the grid is short. */
export function gridColumnCount(table: WmlTable): number {
  let widest = 0;
  for (const row of table.rows) {
    let g = gridBeforeOf(row) + gridAfterOf(row);
    for (const cell of row.cells) g += gridSpanOf(cell);
    widest = Math.max(widest, g);
  }
  return Math.max(widest, gridWidths(table).length);
}

/** `<w:gridCol w:w>` values in twips. */
export function gridWidths(table: WmlTable): number[] {
  return wChildren(table.tblGrid, "gridCol").map((c) => intAttr(c, "w") ?? 0);
}

/** Replace `<w:tblGrid>` with the given column widths. */
export function setGridWidths(table: WmlTable, widths: readonly number[]): void {
  const grid = wEl(
    "tblGrid",
    {},
    widths.map((w) => wEl("gridCol", { w: String(Math.max(0, Math.round(w))) })),
  );
  if (table.tblGrid) {
    // Keep a tblGridChange (tracked change) if there is one.
    const change = wChild(table.tblGrid, "tblGridChange");
    if (change) (grid.children as XmlElement[]).push(change);
  }
  table.tblGrid = grid;
}

/** The placement covering a grid column in one row, if any. */
export function placementAt(
  row: readonly TableCellPlacement[],
  gridCol: number,
): TableCellPlacement | undefined {
  return row.find((p) => p.gridStart <= gridCol && gridCol < p.gridStart + p.gridSpan);
}

export function ensureTcPr(cell: WmlTableCell): XmlElement {
  cell.tcPr ??= wEl("tcPr");
  return cell.tcPr;
}

export function ensureTrPr(row: WmlTableRow): XmlElement {
  row.trPr ??= wEl("trPr");
  return row.trPr;
}

export function setGridSpan(cell: WmlTableCell, span: number): void {
  const tcPr = ensureTcPr(cell);
  if (span <= 1) removeWChild(tcPr, "gridSpan");
  else upsertWChild(tcPr, wEl("gridSpan", { val: String(span) }), TC_PR_ORDER);
}

export function setVMerge(cell: WmlTableCell, merge: "restart" | "continue" | undefined): void {
  const tcPr = ensureTcPr(cell);
  if (merge === undefined) removeWChild(tcPr, "vMerge");
  else
    upsertWChild(tcPr, wEl("vMerge", merge === "restart" ? { val: "restart" } : {}), TC_PR_ORDER);
}

/** Set (or with 0 remove) a row's `w:gridBefore` / `w:gridAfter`. */
export function setRowGridSkip(
  row: WmlTableRow,
  which: "gridBefore" | "gridAfter",
  count: number,
): void {
  const trPr = ensureTrPr(row);
  if (count <= 0) removeWChild(trPr, which);
  else upsertWChild(trPr, wEl(which, { val: String(count) }), TR_PR_ORDER);
}

/** Set the cell's preferred width to its grid columns' total, when it is in twips. */
export function syncCellWidth(cell: WmlTableCell, widths: readonly number[], start: number): void {
  const tcW = wChild(cell.tcPr, "tcW");
  const type = wAttrOf(tcW, "type") ?? "dxa";
  if (tcW && type !== "dxa") return;
  let total = 0;
  for (let g = start; g < start + gridSpanOf(cell); g++) total += widths[g] ?? 0;
  upsertWChild(ensureTcPr(cell), wEl("tcW", { w: String(total), type: "dxa" }), TC_PR_ORDER);
}

/**
 * Repair vertical merges after rows or columns moved: a continuation with no
 * merge above becomes a merge start, and a merge start that covers a single
 * row loses its `w:vMerge`. Word reports a file with dangling merges as
 * unreadable in some versions, so every structural edit ends with this.
 */
export function normalizeVerticalMerges(table: WmlTable): void {
  let above = new Map<number, number>();
  for (const row of table.rows) {
    const here = new Map<number, number>();
    let g = gridBeforeOf(row);
    for (const cell of row.cells) {
      const span = gridSpanOf(cell);
      const merge = vMergeOf(cell);
      if (merge === "continue" && above.get(g) !== span) setVMerge(cell, "restart");
      if (merge) here.set(g, span);
      g += span;
    }
    above = here;
  }
  for (const row of cellPlacements(table)) {
    for (const p of row) {
      if (p.vMerge && p.rowSpan === 1) {
        const cell = table.rows[p.row]?.cells[p.cell];
        if (cell) setVMerge(cell, undefined);
      }
    }
  }
}

/** One child of a `<w:tc>` in document order: a paragraph or a pass-through node. */
export type CellNode =
  | { readonly kind: "paragraph"; readonly paragraph: WmlParagraph }
  | { readonly kind: "other"; readonly node: XmlNode };

/** A cell's content in document order (paragraphs interleaved with nested tables etc.). */
export function cellNodes(cell: WmlTableCell): CellNode[] {
  const recognized: CellNode[] = cell.paragraphs.map((paragraph) => ({
    kind: "paragraph",
    paragraph,
  }));
  if (cell.extras.length === 0) return recognized;
  // Mirror the writer: extras keep their slot, recognized children fill the
  // gaps in order; the tcPr (if any) is the first recognized child.
  const offset = cell.tcPr ? 1 : 0;
  const total = recognized.length + offset + cell.extras.length;
  const slots: (CellNode | undefined)[] = Array.from({ length: total });
  for (const e of cell.extras) {
    if (e.slot >= 0 && e.slot < total) slots[e.slot] = { kind: "other", node: e.node };
  }
  let r = -offset;
  for (let i = 0; i < total; i++) {
    if (slots[i] !== undefined) continue;
    if (r >= 0 && r < recognized.length) slots[i] = recognized[r];
    r++;
  }
  const out = slots.filter((n): n is CellNode => n !== undefined);
  while (r < recognized.length) {
    const node = recognized[r++];
    if (node) out.push(node);
  }
  return out;
}

/** Replace a cell's content with nodes in document order. */
export function setCellNodes(cell: WmlTableCell, nodes: readonly CellNode[]): void {
  // Slots count the tcPr, so make sure it exists before computing them.
  ensureTcPr(cell);
  cell.paragraphs = [];
  cell.extras = [];
  nodes.forEach((n, i) => {
    if (n.kind === "paragraph") cell.paragraphs.push(n.paragraph);
    else cell.extras.push({ slot: i + 1, node: n.node });
  });
}

/** Whether a cell holds nothing but one empty paragraph. */
export function isEmptyCell(cell: WmlTableCell): boolean {
  if (cell.extras.some((e) => e.node.kind === "element")) return false;
  return cell.paragraphs.every((p) =>
    p.children.every(
      (c) =>
        c.kind === "run" && c.pieces.every((piece) => piece.kind === "text" && piece.value === ""),
    ),
  );
}

/** An empty paragraph that keeps another paragraph's properties (Word keeps the mark's format). */
export function emptyParagraphLike(source: WmlParagraph | undefined): WmlParagraph {
  const pPr = source?.pPr;
  return {
    kind: "paragraph",
    ...(pPr ? { pPr: structuredClone(pPr) } : {}),
    children: [],
    extras: [],
  };
}

/** A new empty cell carrying a copy of `source`'s formatting (no merge, no span). */
export function emptyCellLike(source: WmlTableCell | undefined): WmlTableCell {
  const tcPr = source?.tcPr ? structuredClone(source.tcPr) : wEl("tcPr");
  removeWChild(tcPr, "gridSpan");
  removeWChild(tcPr, "vMerge");
  removeWChild(tcPr, "hMerge");
  return { tcPr, paragraphs: [emptyParagraphLike(source?.paragraphs[0])], extras: [] };
}

/** A row's cell paragraphs joined into paragraphs with a separator between cells. */
export function joinRowCells(row: WmlTableRow, separator: string): WmlParagraph[] {
  const out: WmlParagraph[] = [];
  let current = emptyParagraphLike(row.cells[0]?.paragraphs[0]);
  row.cells.forEach((cell, c) => {
    cell.paragraphs.forEach((p, j) => {
      if (j > 0) {
        out.push(current);
        current = emptyParagraphLike(p);
      }
      current.children.push(...p.children);
    });
    if (c < row.cells.length - 1) {
      const piece: WmlRunPiece =
        separator === "\t"
          ? { kind: "tab" }
          : { kind: "text", value: separator, preserveSpace: false };
      current.children.push({ kind: "run", pieces: [piece], extras: [] });
    }
  });
  out.push(current);
  return out;
}

/** What separates the cells of a row: Word's Convert Text to Table / Table to Text choices. */
export type TableTextSeparator = "paragraph" | "tab" | "comma" | { readonly other: string };

export function separatorChar(separator: TableTextSeparator): string | undefined {
  if (separator === "paragraph") return undefined;
  if (separator === "tab") return "\t";
  if (separator === "comma") return ",";
  if (separator.other.length !== 1) throw new RangeError("The separator must be one character.");
  return separator.other;
}
