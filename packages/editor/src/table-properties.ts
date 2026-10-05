/**
 * What the Table Properties dialog shows when it opens: the selected table's,
 * row's, column's and cell's current settings, read from their properties
 * (`w:tblPr`, `w:trPr`, `w:tcPr`) — the values the dialog's fields edit.
 */

import {
  type CellMargins,
  childElementsOf,
  getElementAttr,
  type TableAlignment,
  type TableCellVerticalAlign,
  tableColumnWidths,
  type TableRowHeightRule,
  type TableWidth,
  type XmlElement,
} from "@office-kit/docx";
import type { EditorModel } from "./model.js";
import { tableSelection } from "./table-selection.js";

export interface TablePropertiesSnapshot {
  readonly table: {
    /** Undefined when the table has no preferred width (`auto`). */
    readonly width?: TableWidth;
    readonly alignment: TableAlignment;
    readonly indentTwips: number;
  };
  readonly row: {
    readonly index: number;
    readonly height?: { readonly twips: number; readonly rule: TableRowHeightRule };
    readonly cantSplit: boolean;
    readonly header: boolean;
  };
  readonly column: { readonly index: number; readonly widthTwips: number };
  readonly cell: {
    readonly width?: TableWidth;
    readonly verticalAlign: TableCellVerticalAlign;
    /** The cell's own margins; undefined when it uses the table's. */
    readonly margins?: CellMargins;
    readonly noWrap: boolean;
    readonly fitText: boolean;
  };
  readonly options: {
    readonly defaultMargins: CellMargins;
    readonly cellSpacingTwips: number;
    readonly autoResize: boolean;
  };
  readonly altText: { readonly title: string; readonly description: string };
}

const OFF: ReadonlySet<string> = new Set(["0", "false", "off"]);
// Word's default left / right cell margins (0.08″) when a table sets none.
const WORD_DEFAULT_SIDE_MARGIN = 108;
const MARGIN_SIDES = ["top", "left", "bottom", "right"] as const;

function child(el: XmlElement | undefined, local: string): XmlElement | undefined {
  return el && childElementsOf(el).find((c) => c.name.local === local);
}

function attr(el: XmlElement | undefined, name: string): string | undefined {
  return el && getElementAttr(el, name);
}

function num(el: XmlElement | undefined, name: string): number | undefined {
  const raw = attr(el, name);
  const n = raw === undefined ? Number.NaN : Number(raw);
  return Number.isFinite(n) ? n : undefined;
}

function onOff(el: XmlElement | undefined, local: string): boolean {
  const flag = child(el, local);
  return !!flag && !OFF.has(getElementAttr(flag, "val") ?? "true");
}

function width(el: XmlElement | undefined): TableWidth | undefined {
  const type = attr(el, "type");
  if (type !== "dxa" && type !== "pct") return undefined;
  return { type, value: num(el, "w") ?? 0 };
}

function margins(container: XmlElement | undefined): CellMargins | undefined {
  if (!container) return undefined;
  const out: { top?: number; left?: number; bottom?: number; right?: number } = {};
  for (const side of MARGIN_SIDES) {
    const el =
      child(container, side) ??
      child(container, side === "left" ? "start" : side === "right" ? "end" : side);
    const w = num(el, "w");
    if (w !== undefined) out[side] = w;
  }
  return out;
}

/** The current settings of the table, row, column and cell holding the selection. */
export function tablePropertiesSnapshot(model: EditorModel): TablePropertiesSnapshot | undefined {
  const ts = tableSelection(model);
  if (!ts) return undefined;
  const { table, focus } = ts;
  const tblPr = table.tblPr;
  const row = table.rows[focus.row];
  const trPr = row?.trPr;
  const cell = row?.cells[focus.cell];
  const tcPr = cell?.tcPr;
  const placement = ts.placements[focus.row]?.[focus.cell];
  const column = placement?.gridStart ?? 0;

  const jc = attr(child(tblPr, "jc"), "val");
  const trHeight = child(trPr, "trHeight");
  const heightTwips = num(trHeight, "val");
  const rule = attr(trHeight, "hRule") ?? "atLeast";
  const vAlign = attr(child(tcPr, "vAlign"), "val");
  const layout = attr(child(tblPr, "tblLayout"), "type");
  const tableMargins = margins(child(tblPr, "tblCellMar"));
  const tableWidth = width(child(tblPr, "tblW"));
  const cellWidth = width(child(tcPr, "tcW"));
  const cellMargins = margins(child(tcPr, "tcMar"));
  const caption = attr(child(tblPr, "tblCaption"), "val");
  const description = attr(child(tblPr, "tblDescription"), "val");

  return {
    table: {
      ...(tableWidth ? { width: tableWidth } : {}),
      alignment: jc === "center" || jc === "right" ? jc : "left",
      indentTwips: num(child(tblPr, "tblInd"), "w") ?? 0,
    },
    row: {
      index: focus.row,
      ...(heightTwips === undefined || rule === "auto"
        ? {}
        : { height: { twips: heightTwips, rule: rule === "exact" ? "exact" : "atLeast" } }),
      cantSplit: onOff(trPr, "cantSplit"),
      header: onOff(trPr, "tblHeader"),
    },
    column: { index: column, widthTwips: tableColumnWidths(table)[column] ?? 0 },
    cell: {
      ...(cellWidth ? { width: cellWidth } : {}),
      verticalAlign: vAlign === "center" || vAlign === "bottom" ? vAlign : "top",
      ...(cellMargins ? { margins: cellMargins } : {}),
      noWrap: onOff(tcPr, "noWrap"),
      fitText: onOff(tcPr, "tcFitText"),
    },
    options: {
      defaultMargins: {
        top: tableMargins?.top ?? 0,
        left: tableMargins?.left ?? WORD_DEFAULT_SIDE_MARGIN,
        bottom: tableMargins?.bottom ?? 0,
        right: tableMargins?.right ?? WORD_DEFAULT_SIDE_MARGIN,
      },
      cellSpacingTwips: num(child(tblPr, "tblCellSpacing"), "w") ?? 0,
      autoResize: layout !== "fixed",
    },
    altText: { title: caption ?? "", description: description ?? "" },
  };
}
