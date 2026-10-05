import {
  addTable,
  createDocx,
  getTableCellText,
  getTableLook,
  getTableStyle,
  openDocx,
  paragraphText,
  tableCellPlacements,
  tableColumnWidths,
  tables,
  toUint8Array,
  validatePackage,
  type WmlTable,
} from "@office-kit/docx";
import { describe, expect, it } from "vitest";
import { EditorModel } from "../model.js";
import { caretAt, type Selection } from "../selection.js";
import { adjacentCellPosition, cellStart, selectInTable, tableSelection } from "../table-selection.js";
import {
  autoFitCommand,
  cellAlignmentCommand,
  columnWidthSelectionCommand,
  convertTableToTextCommand,
  convertTextToTableCommand,
  deleteCellsCommand,
  deleteColumnsCommand,
  deleteRowCommand,
  deleteTableCommand,
  distributeColumnsCommand,
  distributeRowsCommand,
  insertColumnsCommand,
  insertFormulaCommand,
  insertRowsCommand,
  insertTableCommand,
  mergeCellsCommand,
  modifyTableStyleCommand,
  newTableStyleCommand,
  rangeBordersCommand,
  repeatHeaderRowsCommand,
  rowHeightSelectionCommand,
  setTableLookCommand,
  setTableStyleCommand,
  shadeCellsCommand,
  sortTableCommand,
  splitCellsCommand,
  splitTableCommand,
  tablePropertiesCommand,
  textDirectionCommand,
} from "./table.js";
import { runCommand } from "./types.js";

/** An editor over a document holding one table, caret in its first cell. */
function withTable(rows: string[][]): { model: EditorModel; table: () => WmlTable } {
  const doc = createDocx({ paragraphs: [] });
  addTable(doc, rows);
  const model = new EditorModel(doc);
  model.setSelection(caretAt(cellStart(0, 0, 0)));
  return { model, table: () => tables(model.doc)[0]! };
}

function select(model: EditorModel, from: [number, number], to: [number, number]): void {
  const sel: Selection = { anchor: cellStart(0, ...from), focus: cellStart(0, ...to) };
  model.setSelection(sel);
}

function texts(table: WmlTable): string[][] {
  return table.rows.map((r) => r.cells.map((c) => c.paragraphs.map(paragraphText).join("|")));
}

function xmlOf(model: EditorModel, part = "/word/document.xml"): string {
  const reopened = openDocx(toUint8Array(model.doc));
  expect(validatePackage(reopened.opc)).toEqual([]);
  return new TextDecoder().decode(reopened.opc.parts.get(part)?.data);
}

describe("insert table", () => {
  it("inserts a Table Grid table across the text width with the caret in A1", () => {
    const model = new EditorModel(createDocx({ paragraphs: ["x"] }));
    model.setSelection(caretAt({ block: 0, inline: 0, offset: 0 }));
    runCommand(model, insertTableCommand, { rows: 2, cols: 3 });
    const table = tables(model.doc)[0]!;
    expect(getTableStyle(table)).toBe("TableGrid");
    expect(model.selection?.focus.cell).toEqual({ row: 0, col: 0 });
    expect(xmlOf(model, "/word/styles.xml")).toContain('w:styleId="TableGrid"');
    // Letter paper with 1 in margins leaves 6.5 in = 9360 twips.
    expect(tableColumnWidths(table)).toEqual([3120, 3120, 3120]);
  });

  it("fills Quick Table content and applies the AutoFit choice", () => {
    const model = new EditorModel(createDocx({ paragraphs: [""] }));
    model.setSelection(caretAt({ block: 0 }));
    runCommand(model, insertTableCommand, {
      rows: 2,
      cols: 2,
      cells: [["Item", "Qty"], ["Pens", "3"]],
      autoFit: "window",
      styleId: "GridTable4-Accent1",
    });
    const table = tables(model.doc)[0]!;
    expect(getTableCellText(table, 1, 0)).toBe("Pens");
    expect(xmlOf(model)).toContain('<w:tblW w:w="5000" w:type="pct"/>');
  });
});

describe("rows and columns", () => {
  it("inserts above / below / left / right of the selection", () => {
    const { model, table } = withTable([["a", "b"], ["c", "d"]]);
    runCommand(model, insertRowsCommand, { where: "below" });
    runCommand(model, insertColumnsCommand, { where: "right" });
    expect(texts(table())).toEqual([["a", "", "b"], ["", "", ""], ["c", "", "d"]]);
    select(model, [0, 0], [2, 0]);
    runCommand(model, insertRowsCommand, { where: "above" });
    expect(table().rows).toHaveLength(6);
    xmlOf(model);
  });

  it("deletes rows, columns, cells and the table", () => {
    const { model, table } = withTable([["a", "b", "c"], ["d", "e", "f"], ["g", "h", "i"]]);
    select(model, [1, 0], [1, 0]);
    runCommand(model, deleteRowCommand, {});
    expect(texts(table()).map((r) => r[0])).toEqual(["a", "g"]);
    model.setSelection(caretAt(cellStart(0, 0, 1)));
    runCommand(model, deleteColumnsCommand, undefined);
    expect(texts(table())).toEqual([["a", "c"], ["g", "i"]]);
    runCommand(model, deleteCellsCommand, { shift: "up" });
    expect(texts(table()).map((r) => r[0])).toEqual(["g", ""]);
    runCommand(model, deleteTableCommand, undefined);
    expect(tables(model.doc)).toHaveLength(0);
    expect(model.doc.document.body.blocks).toHaveLength(1);
  });
});

describe("merge / split", () => {
  it("merges the selected cells and splits them again", () => {
    const { model, table } = withTable([["a", "b"], ["c", "d"]]);
    model.setSelection(caretAt(cellStart(0, 0, 0)));
    expect(mergeCellsCommand.isEnabled?.(model)).toBe(false);
    select(model, [0, 0], [0, 1]);
    runCommand(model, mergeCellsCommand, {});
    expect(texts(table())[0]).toEqual(["a|b"]);
    runCommand(model, splitCellsCommand, { columns: 2, rows: 1 });
    expect(table().rows[0]?.cells).toHaveLength(2);
    xmlOf(model);
  });

  it("splits the table at the caret row", () => {
    const { model } = withTable([["a"], ["b"], ["c"]]);
    model.setSelection(caretAt(cellStart(0, 1, 0)));
    runCommand(model, splitTableCommand, undefined);
    expect(tables(model.doc)).toHaveLength(2);
    expect(model.selection?.focus.block).toBe(2);
  });
});

describe("table design", () => {
  it("applies a built-in style and toggles style options", () => {
    const { model, table } = withTable([["a"]]);
    runCommand(model, setTableStyleCommand, { styleId: "ListTable3-Accent2" });
    runCommand(model, setTableLookCommand, { totalRow: true, bandedRows: false });
    expect(getTableStyle(table())).toBe("ListTable3-Accent2");
    expect(getTableLook(table())).toMatchObject({ headerRow: true, totalRow: true, bandedRows: false });
    expect(xmlOf(model, "/word/styles.xml")).toContain('w:styleId="ListTable3-Accent2"');
    runCommand(model, setTableStyleCommand, { styleId: undefined });
    expect(getTableStyle(table())).toBeUndefined();
  });

  it("creates and modifies a table style", () => {
    const { model, table } = withTable([["a"]]);
    const id = runCommand(model, newTableStyleCommand, { name: "My Grid" });
    expect(id).toBe("MyGrid");
    expect(getTableStyle(table())).toBe("MyGrid");
    runCommand(model, modifyTableStyleCommand, {
      styleId: "MyGrid",
      region: "firstRow",
      formatting: { bold: true, fill: "FFFF00" },
    });
    expect(xmlOf(model, "/word/styles.xml")).toContain('<w:tblStylePr w:type="firstRow">');
  });

  it("shades and borders the selected cells", () => {
    const { model, table } = withTable([["a", "b"], ["c", "d"]]);
    select(model, [0, 0], [1, 0]);
    runCommand(model, shadeCellsCommand, { fill: "FFC000" });
    runCommand(model, rangeBordersCommand, {
      edges: "outside",
      border: { style: "single", size: 12, color: "FF0000" },
    });
    const xml = xmlOf(model);
    expect(xml.match(/w:fill="FFC000"/g)).toHaveLength(2);
    expect(JSON.stringify(table().rows[0]?.cells[1]?.tcPr)).toContain('"left"');
  });
});

describe("table layout", () => {
  it("sets properties from the Table Properties dialog in one step", () => {
    const { model, table } = withTable([["a", "b"]]);
    runCommand(model, tablePropertiesCommand, {
      table: { width: { type: "pct", value: 2500 }, alignment: "center", indentTwips: 0 },
      rows: { height: { twips: 720, rule: "exact" }, cantSplit: true, header: true },
      columnWidthTwips: 2000,
      cells: { verticalAlign: "center", margins: { left: 0 }, noWrap: true },
      options: { defaultMargins: { left: 72, right: 72 }, cellSpacingTwips: 20, autoResize: false },
      altText: { title: "T", description: "D" },
    });
    const xml = xmlOf(model);
    for (const part of [
      '<w:jc w:val="center"/>',
      '<w:trHeight w:val="720" w:hRule="exact"/>',
      "<w:cantSplit/>",
      "<w:tblHeader/>",
      '<w:vAlign w:val="center"/>',
      "<w:noWrap/>",
      '<w:tblCellSpacing w:w="20" w:type="dxa"/>',
      '<w:tblLayout w:type="fixed"/>',
      '<w:tblCaption w:val="T"/>',
    ]) {
      expect(xml).toContain(part);
    }
    expect(tableColumnWidths(table())[0]).toBe(2000);
    expect(model.canUndo()).toBe(true);
  });

  it("height / width spinners, distribute, autofit, alignment and text direction", () => {
    const { model, table } = withTable([["a", "b"], ["c", "d"]]);
    runCommand(model, rowHeightSelectionCommand, { twips: 500 });
    runCommand(model, columnWidthSelectionCommand, { twips: 3000 });
    runCommand(model, distributeColumnsCommand, undefined);
    expect(new Set(tableColumnWidths(table())).size).toBe(1);
    select(model, [0, 0], [1, 0]);
    runCommand(model, distributeRowsCommand, { heightTwips: 400 });
    runCommand(model, cellAlignmentCommand, { horizontal: "right", vertical: "bottom" });
    runCommand(model, textDirectionCommand, { direction: "tbRl" });
    runCommand(model, autoFitCommand, { mode: "fixed" });
    const xml = xmlOf(model);
    expect(xml).toContain('<w:trHeight w:val="400" w:hRule="atLeast"/>');
    expect(xml).toContain('<w:jc w:val="right"/>');
    expect(xml).toContain('<w:vAlign w:val="bottom"/>');
    expect(xml).toContain('<w:textDirection w:val="tbRl"/>');
  });

  it("toggles Repeat Header Rows from the first row through the selection", () => {
    const { model, table } = withTable([["a"], ["b"], ["c"]]);
    model.setSelection(caretAt(cellStart(0, 1, 0)));
    runCommand(model, repeatHeaderRowsCommand, undefined);
    const header = (r: number): boolean => JSON.stringify(table().rows[r]?.trPr ?? {}).includes("tblHeader");
    expect([header(0), header(1), header(2)]).toEqual([true, true, false]);
    model.setSelection(caretAt(cellStart(0, 0, 0)));
    expect(repeatHeaderRowsCommand.isActive?.(model)).toBe(true);
  });

  it("sorts, converts to text and back", () => {
    const { model } = withTable([["b", "2"], ["a", "1"]]);
    runCommand(model, sortTableCommand, { keys: [{ column: 0 }] });
    expect(texts(tables(model.doc)[0]!)).toEqual([["a", "1"], ["b", "2"]]);
    runCommand(model, convertTableToTextCommand, { separator: "tab" });
    expect(tables(model.doc)).toHaveLength(0);
    model.setSelection({ anchor: { block: 0, inline: 0, offset: 0 }, focus: { block: 1, inline: 0, offset: 0 } });
    expect(convertTextToTableCommand.isEnabled?.(model)).toBe(true);
    runCommand(model, convertTextToTableCommand, { separator: "tab" });
    expect(texts(tables(model.doc)[0]!)).toEqual([["a", "1"], ["b", "2"]]);
    xmlOf(model);
  });

  it("inserts a formula field with its computed result", () => {
    const { model, table } = withTable([["1"], ["2"], [""]]);
    model.setSelection(caretAt(cellStart(0, 2, 0)));
    expect(runCommand(model, insertFormulaCommand, { formula: "=SUM(ABOVE)", numberFormat: "0.00" })).toBe(3);
    const xml = xmlOf(model);
    expect(xml).toContain('=SUM(ABOVE) \\# "0.00"');
    expect(xml).toContain("<w:t>3.00</w:t>");
    expect(table().rows).toHaveLength(3);
  });
});

describe("table selection", () => {
  it("expands a cross-cell selection to whole merged cells", () => {
    const { model } = withTable([["a", "b", "c"], ["d", "e", "f"]]);
    select(model, [0, 0], [0, 1]);
    runCommand(model, mergeCellsCommand, {});
    // Row 1 cell 1 ("e") sits under the merged A1:B1.
    select(model, [1, 1], [1, 1]);
    model.setSelection({ anchor: cellStart(0, 0, 0), focus: cellStart(0, 1, 2) });
    expect(tableSelection(model)?.range).toEqual({ firstRow: 0, lastRow: 1, firstColumn: 0, lastColumn: 2 });
  });

  it("selects the row / column / table and moves with Tab", () => {
    const { model } = withTable([["a", "b"], ["c", "d"]]);
    model.setSelection(caretAt(cellStart(0, 1, 1)));
    expect(selectInTable(model, "row")).toEqual({ anchor: cellStart(0, 1, 0), focus: cellStart(0, 1, 1) });
    expect(selectInTable(model, "column")?.anchor).toEqual(cellStart(0, 0, 1));
    expect(selectInTable(model, "table")?.focus).toEqual(cellStart(0, 1, 1));
    expect(adjacentCellPosition(model, 1)).toBeUndefined();
    expect(adjacentCellPosition(model, -1)).toEqual(cellStart(0, 1, 0));
  });

  it("skips merged continuation cells when tabbing", () => {
    const { model, table } = withTable([["a", "b"], ["c", "d"]]);
    select(model, [0, 0], [1, 0]);
    runCommand(model, mergeCellsCommand, {});
    expect(tableCellPlacements(table())[1]?.[0]?.rowSpan).toBe(0);
    model.setSelection(caretAt(cellStart(0, 0, 1)));
    expect(adjacentCellPosition(model, 1)).toEqual(cellStart(0, 1, 1));
  });
});
