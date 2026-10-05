import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { describe, expect, it } from "vitest";
import {
  addTable,
  autoFitTable,
  convertTextToTable,
  createDocx,
  deleteTableCells,
  deleteTableColumns,
  deleteTableRows,
  type Docx,
  distributeTableColumns,
  evaluateTableFormula,
  formatFieldNumber,
  getTableCellText,
  insertTableColumn,
  insertTableRow,
  mergeTableCells,
  openDocx,
  paragraphText,
  removeTableRow,
  setTableAlignment,
  setTableAltText,
  setTableCellMargins,
  setTableCellSpacing,
  setTableCellText,
  setTableCellTextDirection,
  setTableCellWidth,
  setTableColumnWidth,
  setTableDefaultCellMargins,
  setTableIndent,
  setTableLayout,
  setTablePosition,
  setTableRangeBorders,
  setTableWidth,
  sortTableRows,
  splitTable,
  splitTableCell,
  tableCellBlocks,
  tableCellPlacements,
  tableColumnCount,
  tableColumnWidths,
  tables,
  toUint8Array,
  unwrapTable,
  validatePackage,
  type WmlTable,
} from "./index.js";

function grid(rows: number, cols: number): { doc: Docx; table: WmlTable } {
  const doc = createDocx({ paragraphs: [] });
  const table = addTable(
    doc,
    Array.from({ length: rows }, (_, r) => Array.from({ length: cols }, (_, c) => `${r}${c}`)),
    { totalWidthTwips: cols * 1000 },
  );
  return { doc, table };
}

/** Save, reopen and return the reopened document plus its document.xml. */
function roundTrip(doc: Docx): { doc: Docx; xml: string; table: WmlTable } {
  const bytes = toUint8Array(doc);
  const reopened = openDocx(bytes);
  expect(validatePackage(reopened.opc)).toEqual([]);
  const xml = new TextDecoder().decode(reopened.opc.parts.get("/word/document.xml")?.data);
  const table = tables(reopened)[0];
  if (!table) throw new Error("no table after round trip");
  return { doc: reopened, xml, table };
}

function texts(table: WmlTable): string[][] {
  return table.rows.map((row) => row.cells.map((c) => c.paragraphs.map(paragraphText).join("|")));
}

describe("tableCellPlacements", () => {
  it("places spanned and vertically merged cells on the grid", () => {
    const { table } = grid(3, 3);
    mergeTableCells(table, { firstRow: 0, lastRow: 1, firstColumn: 0, lastColumn: 1 });
    const p = tableCellPlacements(table);
    expect(p[0]?.[0]).toMatchObject({ gridStart: 0, gridSpan: 2, rowSpan: 2, vMerge: "restart" });
    expect(p[1]?.[0]).toMatchObject({ gridStart: 0, gridSpan: 2, rowSpan: 0, vMerge: "continue" });
    expect(p[1]?.[1]).toMatchObject({ gridStart: 2, gridSpan: 1, rowSpan: 1 });
  });
});

describe("insertTableRow / deleteTableRows", () => {
  it("inserts above and below, copying the neighbour's cell formatting", () => {
    const { doc, table } = grid(2, 2);
    insertTableRow(table, 1);
    insertTableRow(table, 0);
    expect(texts(table)).toEqual([["", ""], ["00", "01"], ["", ""], ["10", "11"]]);
    const { xml, table: back } = roundTrip(doc);
    expect(back.rows).toHaveLength(4);
    expect(xml.match(/<w:tcW /g)).toHaveLength(8);
  });

  it("grows a vertical merge that spans the insertion point", () => {
    const { doc, table } = grid(3, 2);
    mergeTableCells(table, { firstRow: 0, lastRow: 1, firstColumn: 0, lastColumn: 0 });
    insertTableRow(table, 1);
    const p = tableCellPlacements(table);
    expect(p[0]?.[0]?.rowSpan).toBe(3);
    const { xml } = roundTrip(doc);
    expect(xml.match(/<w:vMerge\/>/g)).toHaveLength(2);
  });

  it("restarts a merge whose first row was deleted", () => {
    const { doc, table } = grid(3, 1);
    mergeTableCells(table, { firstRow: 0, lastRow: 2, firstColumn: 0, lastColumn: 0 });
    deleteTableRows(table, 0);
    expect(texts(table)[0]).toEqual(["00|10|20"]);
    expect(tableCellPlacements(table)[0]?.[0]).toMatchObject({ vMerge: "restart", rowSpan: 2 });
    const { xml } = roundTrip(doc);
    expect(xml).toContain('<w:vMerge w:val="restart"/>');
  });

  it("removeTableRow also repairs merges", () => {
    const { table } = grid(2, 1);
    mergeTableCells(table, { firstRow: 0, lastRow: 1, firstColumn: 0, lastColumn: 0 });
    removeTableRow(table, 0);
    expect(tableCellPlacements(table)[0]?.[0]?.vMerge).toBeUndefined();
  });
});

describe("insertTableColumn / deleteTableColumns", () => {
  it("inserts a column left and right and widens tblGrid", () => {
    const { doc, table } = grid(2, 2);
    insertTableColumn(table, 1);
    insertTableColumn(table, 3);
    expect(tableColumnCount(table)).toBe(4);
    expect(texts(table)[0]).toEqual(["00", "", "01", ""]);
    const { xml } = roundTrip(doc);
    expect(xml.match(/<w:gridCol /g)).toHaveLength(4);
    expect(xml).toContain('<w:tblW w:w="4000" w:type="dxa"/>');
  });

  it("widens a spanning cell instead of splitting it", () => {
    const { table } = grid(2, 2);
    mergeTableCells(table, { firstRow: 0, lastRow: 0, firstColumn: 0, lastColumn: 1 });
    insertTableColumn(table, 1);
    expect(tableCellPlacements(table)[0]?.[0]?.gridSpan).toBe(3);
    expect(table.rows[1]?.cells).toHaveLength(3);
  });

  it("deletes columns and narrows spanning cells", () => {
    const { doc, table } = grid(2, 3);
    mergeTableCells(table, { firstRow: 0, lastRow: 0, firstColumn: 0, lastColumn: 1 });
    expect(deleteTableColumns(table, 1)).toBe(true);
    expect(texts(table)).toEqual([["00|01", "02"], ["10", "12"]]);
    expect(tableColumnWidths(table)).toEqual([1000, 1000]);
    roundTrip(doc);
  });

  it("reports when no column is left", () => {
    const { table } = grid(1, 2);
    expect(deleteTableColumns(table, 0, 1)).toBe(false);
  });
});

describe("deleteTableCells", () => {
  it("shifts cells left", () => {
    const { doc, table } = grid(2, 3);
    deleteTableCells(table, { firstRow: 0, lastRow: 0, firstColumn: 0, lastColumn: 0 }, "left");
    expect(texts(table)[0]).toEqual(["01", "02"]);
    const { xml } = roundTrip(doc);
    expect(xml).toContain('<w:gridAfter w:val="1"/>');
  });

  it("shifts cells up", () => {
    const { table } = grid(3, 2);
    deleteTableCells(table, { firstRow: 0, lastRow: 0, firstColumn: 1, lastColumn: 1 }, "up");
    expect(texts(table).map((r) => r[1])).toEqual(["11", "21", ""]);
  });
});

describe("mergeTableCells / splitTableCell", () => {
  it("merges a rectangle keeping every cell's text", () => {
    const { doc, table } = grid(2, 2);
    mergeTableCells(table, { firstRow: 0, lastRow: 1, firstColumn: 0, lastColumn: 1 });
    expect(texts(table)).toEqual([["00|01|10|11"], [""]]);
    const { xml } = roundTrip(doc);
    expect(xml).toContain('<w:gridSpan w:val="2"/>');
    expect(xml).toContain('<w:vMerge w:val="restart"/>');
  });

  it("rejects a range that cuts a merged cell", () => {
    const { table } = grid(2, 3);
    mergeTableCells(table, { firstRow: 0, lastRow: 0, firstColumn: 0, lastColumn: 1 });
    expect(() =>
      mergeTableCells(table, { firstRow: 0, lastRow: 1, firstColumn: 1, lastColumn: 2 }),
    ).toThrow(RangeError);
  });

  it("splits a merged cell back into its grid columns", () => {
    const { table } = grid(1, 2);
    mergeTableCells(table, { firstRow: 0, lastRow: 0, firstColumn: 0, lastColumn: 1 });
    splitTableCell(table, 0, 0, { columns: 2, rows: 1 });
    expect(table.rows[0]?.cells).toHaveLength(2);
    expect(tableColumnCount(table)).toBe(2);
  });

  it("splits a cell into more columns than it spans by refining the grid", () => {
    const { doc, table } = grid(2, 2);
    splitTableCell(table, 0, 0, { columns: 2, rows: 1 });
    expect(tableColumnCount(table)).toBe(3);
    expect(tableColumnWidths(table)).toEqual([500, 500, 1000]);
    expect(tableCellPlacements(table)[1]?.[0]?.gridSpan).toBe(2);
    roundTrip(doc);
  });

  it("splits a cell into rows, merging the rest of the row over them", () => {
    const { doc, table } = grid(1, 2);
    splitTableCell(table, 0, 1, { columns: 1, rows: 2 });
    expect(table.rows).toHaveLength(2);
    expect(tableCellPlacements(table)[0]?.[0]?.rowSpan).toBe(2);
    expect(tableCellPlacements(table)[1]?.[1]?.rowSpan).toBe(1);
    roundTrip(doc);
  });
});

describe("splitTable", () => {
  it("splits before a row with a paragraph between", () => {
    const { doc, table } = grid(3, 1);
    const second = splitTable(doc, table, 1);
    expect(table.rows).toHaveLength(1);
    expect(second.rows).toHaveLength(2);
    expect(doc.document.body.blocks.map((b) => b.kind)).toEqual(["table", "paragraph", "table"]);
    roundTrip(doc);
  });
});

describe("table properties", () => {
  it("writes tblPr children in schema order", () => {
    const { doc, table } = grid(1, 1);
    setTableAltText(table, { title: "Totals", description: "Quarterly totals" });
    setTableLayout(table, "fixed");
    setTableDefaultCellMargins(table, { top: 20, left: 100, bottom: 20, right: 100 });
    setTableCellSpacing(table, 30);
    setTableIndent(table, 360);
    setTableAlignment(table, "center");
    setTableWidth(table, { type: "pct", value: 2500 });
    setTablePosition(table, { horizontalAnchor: "margin", xAlign: "center", y: 200 });
    const { xml } = roundTrip(doc);
    const tblPr = xml.slice(xml.indexOf("<w:tblPr>"), xml.indexOf("</w:tblPr>"));
    const order = [...tblPr.matchAll(/<w:(\w+)/g)].map((m) => m[1]);
    expect(order.filter((n) => n !== "top" && n !== "left" && n !== "bottom" && n !== "right")).toEqual([
      "tblPr",
      "tblpPr",
      "tblW",
      "jc",
      "tblCellSpacing",
      "tblInd",
      "tblLayout",
      "tblCellMar",
      "tblLook",
      "tblCaption",
      "tblDescription",
    ]);
    expect(tblPr).toContain('w:tblpXSpec="center"');
  });

  it("removes properties again", () => {
    const { doc, table } = grid(1, 1);
    setTableAlignment(table, "right");
    setTableAlignment(table, undefined);
    setTableAltText(table, { title: "x" });
    setTableAltText(table, {});
    setTablePosition(table, { allowOverlap: false });
    setTablePosition(table, undefined);
    const { xml } = roundTrip(doc);
    expect(xml).not.toMatch(/<w:jc |tblCaption|tblpPr|tblOverlap/);
  });

  it("sets cell width, margins and text direction in tcPr order", () => {
    const { doc, table } = grid(1, 1);
    const cell = table.rows[0]!.cells[0]!;
    setTableCellTextDirection(cell, "btLr");
    setTableCellMargins(cell, { left: 0, right: 0 });
    setTableCellWidth(cell, { type: "dxa", value: 1440 });
    const { xml } = roundTrip(doc);
    expect(xml).toMatch(
      /<w:tcPr><w:tcW w:w="1440" w:type="dxa"\/><w:tcMar><w:left w:w="0" w:type="dxa"\/><w:right w:w="0" w:type="dxa"\/><\/w:tcMar><w:textDirection w:val="btLr"\/><\/w:tcPr>/,
    );
  });

  it("rejects invalid widths", () => {
    const { table } = grid(1, 1);
    expect(() => setTableWidth(table, { type: "dxa", value: -1 })).toThrow(RangeError);
  });
});

describe("column widths and AutoFit", () => {
  it("sets a column width and keeps cell widths in step", () => {
    const { doc, table } = grid(1, 2);
    setTableColumnWidth(table, 0, 2000);
    expect(tableColumnWidths(table)).toEqual([2000, 1000]);
    const { xml } = roundTrip(doc);
    expect(xml).toContain('<w:tcW w:w="2000" w:type="dxa"/>');
    expect(xml).toContain('<w:tblW w:w="3000" w:type="dxa"/>');
  });

  it("distributes columns evenly", () => {
    const { table } = grid(1, 3);
    setTableColumnWidth(table, 0, 2000);
    distributeTableColumns(table, 0, 2);
    expect(tableColumnWidths(table)).toEqual([1333, 1333, 1334]);
  });

  it("AutoFit window / contents / fixed", () => {
    const { doc, table } = grid(1, 2);
    autoFitTable(table, "window");
    let xml = roundTrip(doc).xml;
    expect(xml).toContain('<w:tblW w:w="5000" w:type="pct"/>');
    expect(xml).toContain('<w:tcW w:w="2500" w:type="pct"/>');
    autoFitTable(table, "contents");
    xml = roundTrip(doc).xml;
    expect(xml).toContain('<w:tblW w:w="0" w:type="auto"/>');
    autoFitTable(table, "fixed");
    xml = roundTrip(doc).xml;
    expect(xml).toContain('<w:tblLayout w:type="fixed"/>');
    expect(xml).toContain('<w:tcW w:w="1000" w:type="dxa"/>');
  });
});

describe("setTableRangeBorders", () => {
  it("puts whole-table borders on the table", () => {
    const { doc, table } = grid(2, 2);
    setTableRangeBorders(table, { firstRow: 0, lastRow: 1, firstColumn: 0, lastColumn: 1 }, "outside", {
      style: "double",
      size: 6,
      color: "FF0000",
    });
    const { xml } = roundTrip(doc);
    expect(xml).toContain('<w:top w:val="double" w:sz="6" w:space="0" w:color="FF0000"/>');
    expect(xml).not.toContain("tcBorders");
  });

  it("puts part-table borders on the cells and their neighbours", () => {
    const { doc, table } = grid(2, 2);
    setTableRangeBorders(table, { firstRow: 0, lastRow: 0, firstColumn: 0, lastColumn: 0 }, "bottom", {
      style: "single",
      size: 12,
      color: "auto",
    });
    const { table: back } = roundTrip(doc);
    const xmlOf = (r: number, c: number): string =>
      JSON.stringify(back.rows[r]?.cells[c]?.tcPr ?? null);
    expect(xmlOf(0, 0)).toContain('"bottom"');
    expect(xmlOf(1, 0)).toContain('"top"');
    expect(xmlOf(0, 1)).not.toContain("tcBorders");
  });

  it("draws diagonal borders and clears with none", () => {
    const { doc, table } = grid(1, 1);
    const all = { firstRow: 0, lastRow: 0, firstColumn: 0, lastColumn: 0 };
    setTableRangeBorders(table, all, "tl2br", { style: "single", size: 4, color: "auto" });
    setTableRangeBorders(table, all, "none", undefined);
    const { xml } = roundTrip(doc);
    expect(xml).toContain("<w:tl2br ");
    expect(xml).toContain('<w:insideH w:val="nil"/>');
  });

  it("validates the border at the boundary", () => {
    const { table } = grid(1, 1);
    const all = { firstRow: 0, lastRow: 0, firstColumn: 0, lastColumn: 0 };
    expect(() => setTableRangeBorders(table, all, "all", { style: "zigzag", size: 4, color: "auto" })).toThrow();
    expect(() => setTableRangeBorders(table, all, "all", { style: "single", size: 1, color: "auto" })).toThrow();
  });
});

describe("sortTableRows", () => {
  it("sorts by number below a header row, descending", () => {
    const doc = createDocx({ paragraphs: [] });
    const table = addTable(doc, [["Name", "Qty"], ["b", "10"], ["a", "9"], ["c", "1,200"]]);
    sortTableRows(table, { keys: [{ column: 1, type: "number", order: "descending" }], headerRow: true });
    expect(texts(table).map((r) => r[0])).toEqual(["Name", "c", "b", "a"]);
  });

  it("sorts text with a second key", () => {
    const doc = createDocx({ paragraphs: [] });
    const table = addTable(doc, [["b", "2"], ["a", "2"], ["a", "1"]]);
    sortTableRows(table, { keys: [{ column: 0 }, { column: 1, type: "number" }] });
    expect(texts(table)).toEqual([["a", "1"], ["a", "2"], ["b", "2"]]);
  });
});

describe("text ↔ table", () => {
  it("converts tab-separated paragraphs keeping run formatting", () => {
    const doc = createDocx({ paragraphs: ["a\tb", "c\td\te"] });
    const table = convertTextToTable(doc, 0, 1, { separator: "tab" });
    expect(tableColumnCount(table)).toBe(3);
    expect(texts(table)).toEqual([["a", "b", ""], ["c", "d", "e"]]);
    roundTrip(doc);
  });

  it("converts paragraphs into cells of n columns", () => {
    const doc = createDocx({ paragraphs: ["1", "2", "3"] });
    const table = convertTextToTable(doc, 0, 2, { separator: "paragraph", columns: 2 });
    expect(texts(table)).toEqual([["1", "2"], ["3", ""]]);
  });

  it("converts a table back to comma-separated text", () => {
    const { doc } = grid(2, 2);
    const paras = unwrapTable(doc, 0, { separator: "comma" }) ?? [];
    expect(paras.map(paragraphText)).toEqual(["00,01", "10,11"]);
  });
});

describe("formulas", () => {
  it("sums above and formats the result", () => {
    const doc = createDocx({ paragraphs: [] });
    const table = addTable(doc, [["Item", "Price"], ["a", "1,000.50"], ["b", "$20"], ["Total", ""]]);
    expect(evaluateTableFormula(table, 3, 1, "=SUM(ABOVE)")).toBeCloseTo(1020.5);
    expect(formatFieldNumber(1020.5, "#,##0.00")).toBe("1,020.50");
    expect(formatFieldNumber(-5, "0;(0)")).toBe("(5)");
    expect(formatFieldNumber(0.5)).toBe("0.5");
  });

  it("supports references, ranges, functions and operators", () => {
    const doc = createDocx({ paragraphs: [] });
    const table = addTable(doc, [["2", "4"], ["6", ""]]);
    expect(evaluateTableFormula(table, 1, 1, "AVERAGE(A1:B1)+A2^2")).toBe(39);
    expect(evaluateTableFormula(table, 1, 1, "IF(A1>1, MAX(LEFT), 0)")).toBe(6);
    expect(evaluateTableFormula(table, 1, 1, "ROUND(10/3, 2)")).toBe(3.33);
    expect(() => evaluateTableFormula(table, 1, 1, "SUM(")).toThrow(SyntaxError);
  });
});

describe("tableCellBlocks", () => {
  it("returns nested tables parsed in document order", () => {
    const { doc } = grid(1, 1);
    const outer = tables(doc)[0]!;
    setTableCellText(outer, 0, 0, "before");
    const files = unzipSync(toUint8Array(doc));
    const xml = strFromU8(files["word/document.xml"]!);
    const nested =
      '<w:tbl><w:tblGrid><w:gridCol w:w="500"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>inner</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>';
    files["word/document.xml"] = strToU8(xml.replace("</w:p></w:tc>", `</w:p>${nested}</w:tc>`));
    const again = openDocx(zipSync(files));
    const cell = tables(again)[0]?.rows[0]?.cells[0];
    const blocks = tableCellBlocks(cell!);
    expect(blocks.map((b) => b.kind)).toEqual(["paragraph", "table", "paragraph"]);
    const inner = blocks[1];
    expect(inner?.kind === "table" && getTableCellText(inner, 0, 0)).toBe("inner");
  });
});
