import {
  addBuiltInTableStyle,
  addTable,
  createDocx,
  type Docx,
  mergeTableCells,
  openDocx,
  setTableCellSpacing,
  setTableCellTextDirection,
  setTableLook,
  setTableRangeBorders,
  setTableAlignment,
  setTableRowHeight,
  setTableStyle,
  setTableWidth,
  toUint8Array,
  type WmlTable,
} from "@office-kit/docx";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { describe, expect, it } from "vitest";
import { renderDocumentHtml } from "./render.js";
import { resolveTable, shadingColor } from "./table-format.js";
import { tableStylePreviews } from "./table-preview.js";

function docWithTable(rows: number, cols: number): { doc: Docx; table: WmlTable } {
  const doc = createDocx({ paragraphs: [] });
  const table = addTable(
    doc,
    Array.from({ length: rows }, (_, r) => Array.from({ length: cols }, (_, c) => `${r}${c}`)),
  );
  return { doc, table };
}

const LOOK = {
  headerRow: true,
  totalRow: false,
  firstColumn: true,
  lastColumn: false,
  bandedRows: true,
  bandedColumns: false,
};

describe("table style conditional formatting", () => {
  it("shades the header row and bands the body rows (Grid Table 4)", () => {
    const { doc, table } = docWithTable(4, 3);
    addBuiltInTableStyle(doc, "GridTable4-Accent1");
    setTableStyle(table, "GridTable4-Accent1");
    setTableLook(table, LOOK);
    const format = resolveTable(doc, table);
    const bg = (r: number, c: number): string | undefined => format.cells[r]?.[c]?.background;
    expect(bg(0, 1)).toBe("#4472C4");
    // Rows 1 and 3 are band 1 (shaded); row 2 is band 2.
    expect(bg(1, 1)).toBe("#D9E2F3");
    expect(bg(2, 1)).toBeUndefined();
    expect(bg(3, 1)).toBe("#D9E2F3");
    expect(format.cells[0]?.[0]?.borders.top?.color).toBe("4472C4");
    const html = renderDocumentHtml(doc);
    // The header row's text is white and bold through the style's rPr.
    expect(html).toMatch(/font-weight:bold[^"]*color:#FFFFFF/);
    expect(html).toContain("background-color:#4472C4");
  });

  it("drops banding and the header when the style options are off", () => {
    const { doc, table } = docWithTable(3, 2);
    addBuiltInTableStyle(doc, "GridTable4-Accent1");
    setTableStyle(table, "GridTable4-Accent1");
    setTableLook(table, { ...LOOK, headerRow: false, bandedRows: false });
    const format = resolveTable(doc, table);
    expect(format.cells.flat().every((c) => c.background === undefined)).toBe(true);
  });

  it("applies the total row, last column and corner cells", () => {
    const { doc, table } = docWithTable(3, 3);
    addBuiltInTableStyle(doc, "GridTable5Dark-Accent2");
    setTableStyle(table, "GridTable5Dark-Accent2");
    setTableLook(table, { ...LOOK, totalRow: true, lastColumn: true });
    const format = resolveTable(doc, table);
    expect(format.cells[2]?.[1]?.background).toBe("#ED7D31");
    expect(format.cells[1]?.[2]?.background).toBe("#ED7D31");
  });

  it("uses band sizes from the style", () => {
    const { doc, table } = docWithTable(5, 1);
    addBuiltInTableStyle(doc, "PlainTable4");
    setTableStyle(table, "PlainTable4");
    setTableLook(table, LOOK);
    const files = unzipSync(toUint8Array(doc));
    files["word/styles.xml"] = strToU8(
      strFromU8(files["word/styles.xml"]!).replace(
        /(w:styleId="PlainTable4">[\s\S]*?)<w:tblStyleRowBandSize w:val="1"\/>/,
        '$1<w:tblStyleRowBandSize w:val="2"/>',
      ),
    );
    const reopened = openDocx(zipSync(files));
    const t = reopened.document.body.blocks[0];
    if (t?.kind !== "table") throw new Error("table expected");
    const shaded = resolveTable(reopened, t).cells.map((row) => row[0]?.background !== undefined);
    expect(shaded).toEqual([false, true, true, false, false]);
  });
});

describe("table layout on the canvas", () => {
  it("renders merged cells with colspan / rowspan and hides continuations", () => {
    const { doc, table } = docWithTable(3, 3);
    mergeTableCells(table, { firstRow: 0, lastRow: 1, firstColumn: 0, lastColumn: 1 });
    const html = renderDocumentHtml(doc);
    expect(html).toContain('colspan="2" rowspan="2"');
    expect(html.match(/<td class="wk-td"/g)).toHaveLength(9 - 4 + 1);
  });

  it("draws diagonal borders, vertical text, exact row heights and cell spacing", () => {
    const { doc, table } = docWithTable(1, 2);
    const one = { firstRow: 0, lastRow: 0, firstColumn: 0, lastColumn: 0 };
    setTableRangeBorders(table, one, "tl2br", { style: "single", size: 8, color: "FF0000" });
    setTableCellTextDirection(table.rows[0]!.cells[1]!, "btLr");
    setTableRowHeight(table.rows[0]!, 720, "exact");
    setTableCellSpacing(table, 40);
    const html = renderDocumentHtml(doc);
    expect(html).toContain("linear-gradient(to top right");
    expect(html).toContain("writing-mode:vertical-rl;transform:rotate(180deg)");
    expect(html).toContain("height:36pt");
    expect(html).toContain("overflow:hidden");
    expect(html).toContain("border-spacing:4pt");
  });

  it("sizes a table from tblW, or from its grid when tblW is auto", () => {
    const { doc, table } = docWithTable(1, 1);
    setTableWidth(table, { type: "pct", value: 2500 });
    setTableAlignment(table, "center");
    let html = renderDocumentHtml(doc);
    expect(html).toContain("width:50%");
    expect(html).toContain("margin-left:auto;margin-right:auto");
    setTableWidth(table, { type: "auto", value: 0 });
    html = renderDocumentHtml(doc);
    // addTable's default grid is 9000 twips.
    expect(html).toContain("width:450pt");
  });

  it("renders a nested table read-only inside its cell", () => {
    const { doc } = docWithTable(1, 1);
    const files = unzipSync(toUint8Array(doc));
    const nested =
      '<w:tbl><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>inner</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>';
    files["word/document.xml"] = strToU8(
      strFromU8(files["word/document.xml"]!).replace("</w:p></w:tc>", `</w:p>${nested}</w:tc>`),
    );
    const html = renderDocumentHtml(openDocx(zipSync(files)));
    expect(html).toContain('<div class="wk-nested" contenteditable="false"><table');
    const nestedHtml = html.slice(html.indexOf("wk-nested"));
    expect(nestedHtml.slice(0, nestedHtml.indexOf("</table>"))).not.toContain("data-wk-");
    expect(html).toContain("inner");
    // The paragraph after the nested table is the cell's second paragraph.
    expect(html).toContain('data-wk-para="1"');
  });
});

/** A `<w:shd>` element with the given attributes. */
function el(attrs: Record<string, string>) {
  return {
    kind: "element" as const,
    name: { uri: "", local: "shd", prefix: "w" },
    attrs: Object.entries(attrs).map(([local, value]) => ({
      name: { uri: "", local, prefix: "w" },
      value,
      isNamespaceDecl: false,
    })),
    children: [],
    xmlSpace: "default" as const,
    selfClosing: true,
  };
}

describe("shading", () => {
  it("blends percentage patterns and ignores automatic fills", () => {
    expect(shadingColor(el({ val: "clear", fill: "auto" }))).toBeUndefined();
    expect(shadingColor(el({ val: "clear", fill: "FF0000" }))).toBe("#FF0000");
    expect(shadingColor(el({ val: "pct50", color: "000000", fill: "FFFFFF" }))).toBe("#808080");
    expect(shadingColor(el({ val: "solid", color: "00FF00", fill: "FFFFFF" }))).toBe("#00FF00");
  });
});

describe("gallery previews", () => {
  it("previews built-in styles not yet in the document", () => {
    const doc = createDocx({ paragraphs: [] });
    const previews = tableStylePreviews(doc, ["TableGrid", "GridTable4-Accent6"], LOOK);
    expect(previews.get("TableGrid")?.[1]?.[1]?.top).toBe("1px solid #000");
    const header = previews.get("GridTable4-Accent6")?.[0]?.[1];
    expect(header?.background).toBe("#70AD47");
    expect(header?.text).toBe("#FFFFFF");
    expect(header?.bold).toBe(true);
    // Nothing was added to the document itself.
    expect(JSON.stringify(doc.stylesCache ?? {})).not.toContain("GridTable4-Accent6");
  });
});
