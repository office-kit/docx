import { describe, expect, it } from "vitest";
import {
  addBuiltInTableStyle,
  addStyle,
  addTable,
  BUILT_IN_TABLE_STYLES,
  builtInTableStyle,
  childElementsOf,
  createDocx,
  getElementAttr,
  getTableLook,
  getTableStyle,
  listStyles,
  openDocx,
  setTableLook,
  setTableStyle,
  setTableStyleFormatting,
  stylesPart,
  tables,
  toUint8Array,
  validatePackage,
} from "./index.js";

function stylesXml(bytes: Uint8Array): string {
  const doc = openDocx(bytes);
  expect(validatePackage(doc.opc)).toEqual([]);
  return new TextDecoder().decode(doc.opc.parts.get("/word/styles.xml")?.data);
}

describe("built-in table styles", () => {
  it("lists Word's gallery: Table Grid, 5 plain, 49 grid and 49 list tables", () => {
    expect(BUILT_IN_TABLE_STYLES).toHaveLength(2 + 5 + 49 + 49);
    expect(BUILT_IN_TABLE_STYLES[0]).toEqual({
      styleId: "TableGrid",
      name: "Table Grid",
      category: "plain",
    });
    const ids = BUILT_IN_TABLE_STYLES.map((s) => s.styleId);
    expect(ids).toContain("GridTable4-Accent1");
    expect(ids).toContain("ListTable7Colorful-Accent6");
    expect(new Set(ids).size).toBe(ids.length);
  });

  it("colors accent styles from the Office theme with theme references", () => {
    const doc = createDocx({ paragraphs: [] });
    const style = builtInTableStyle(doc, "GridTable4-Accent1");
    const xml = JSON.stringify(style);
    // accent1 4472C4 at 60 % tint is 8EAADB, at 20 % D9E2F3 (Word's values).
    expect(xml).toContain('"8EAADB"');
    expect(xml).toContain('"D9E2F3"');
    expect(xml).toContain('"accent1"');
    const regions = childElementsOf(style)
      .filter((c) => c.name.local === "tblStylePr")
      .map((c) => getElementAttr(c, "type"));
    expect(regions).toEqual([
      "firstRow",
      "lastRow",
      "firstCol",
      "lastCol",
      "band1Vert",
      "band1Horz",
    ]);
  });

  it("adds a style once and round-trips it", () => {
    const doc = createDocx({ paragraphs: [] });
    const table = addTable(doc, [["a"]]);
    addBuiltInTableStyle(doc, "GridTable5Dark-Accent2");
    addBuiltInTableStyle(doc, "GridTable5Dark-Accent2");
    setTableStyle(table, "GridTable5Dark-Accent2");
    expect(listStyles(doc).filter((s) => s.styleId === "GridTable5Dark-Accent2")).toHaveLength(1);
    const xml = stylesXml(toUint8Array(doc));
    expect(xml).toContain('<w:style w:type="table" w:styleId="GridTable5Dark-Accent2">');
    expect(xml).toContain('<w:name w:val="Grid Table 5 Dark Accent 2"/>');
    expect(xml).toContain('<w:basedOn w:val="TableNormal"/>');
    expect(getTableStyle(tables(openDocx(toUint8Array(doc)))[0]!)).toBe("GridTable5Dark-Accent2");
  });

  it("rejects unknown ids", () => {
    const doc = createDocx({ paragraphs: [] });
    expect(() => addBuiltInTableStyle(doc, "NoSuchTable")).toThrow(RangeError);
  });
});

describe("table look", () => {
  it("writes both the bitmask and the attributes, and reads either", () => {
    const doc = createDocx({ paragraphs: [] });
    const table = addTable(doc, [["a"]]);
    // buildTextTable writes Word's default look: header row, first column, banded rows.
    expect(getTableLook(table)).toEqual({
      headerRow: true,
      totalRow: false,
      firstColumn: true,
      lastColumn: false,
      bandedRows: true,
      bandedColumns: false,
    });
    setTableLook(table, {
      headerRow: false,
      totalRow: true,
      firstColumn: false,
      lastColumn: true,
      bandedRows: false,
      bandedColumns: true,
    });
    const back = tables(openDocx(toUint8Array(doc)))[0]!;
    expect(getTableLook(back).totalRow).toBe(true);
    const xml = new TextDecoder().decode(
      openDocx(toUint8Array(doc)).opc.parts.get("/word/document.xml")?.data,
    );
    expect(xml).toContain('<w:tblLook w:val="0340" w:firstRow="0" w:lastRow="1"');
  });
});

describe("setTableStyleFormatting", () => {
  it("edits a region of a custom table style", () => {
    const doc = createDocx({ paragraphs: [] });
    addStyle(doc, { type: "table", styleId: "MyTable", name: "My Table", basedOn: "TableNormal" });
    setTableStyleFormatting(doc, "MyTable", "firstRow", {
      bold: true,
      color: "FFFFFF",
      fill: "1F4E79",
      borders: { bottom: { style: "single", size: 12, color: "000000" }, top: null },
    });
    setTableStyleFormatting(doc, "MyTable", "wholeTable", {
      borders: { insideH: { style: "dotted", size: 4, color: "auto" } },
      alignment: "center",
    });
    const xml = stylesXml(toUint8Array(doc));
    expect(xml).toContain(
      '<w:tblStylePr w:type="firstRow"><w:rPr><w:b/><w:bCs/><w:color w:val="FFFFFF"/></w:rPr><w:tcPr><w:tcBorders><w:top w:val="nil"/><w:bottom w:val="single" w:sz="12" w:space="0" w:color="000000"/></w:tcBorders><w:shd w:val="clear" w:color="auto" w:fill="1F4E79"/></w:tcPr></w:tblStylePr>',
    );
    const style = stylesPart(doc)?.styles.find((s) => getElementAttr(s, "styleId") === "MyTable");
    const order = childElementsOf(style!).map((c) => c.name.local);
    expect(order).toEqual(["name", "basedOn", "pPr", "tblPr", "tblStylePr"]);
  });

  it("rejects a missing style and bad colors", () => {
    const doc = createDocx({ paragraphs: [] });
    expect(() => setTableStyleFormatting(doc, "Nope", "firstRow", {})).toThrow(RangeError);
    addStyle(doc, { type: "table", styleId: "T" });
    expect(() => setTableStyleFormatting(doc, "T", "firstRow", { fill: "red" })).toThrow(
      RangeError,
    );
  });
});
