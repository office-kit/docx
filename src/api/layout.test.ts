import { describe, expect, it } from "vitest";
import { getPart } from "../internal/opc/index.js";
import {
  appendParagraph,
  createDocx,
  getDocumentSetting,
  getSectionProperties,
  insertSectionBreak,
  MARGINS_NORMAL,
  openDocx,
  PAGE_SIZE_A4,
  sectionCount,
  sectionIndexAt,
  setDocumentSettingOnOff,
  setDocumentSettingVal,
  setPageMargins,
  setPageOrientation,
  setPageSize,
  setSectionProperties,
  toUint8Array,
  validatePackage,
} from "./index.js";

function reopen(doc: ReturnType<typeof createDocx>): ReturnType<typeof createDocx> {
  const bytes = toUint8Array(doc);
  const next = openDocx(bytes);
  expect(validatePackage(next.opc).filter((i) => i.level === "error")).toEqual([]);
  return next;
}

function documentXml(doc: ReturnType<typeof createDocx>): string {
  toUint8Array(doc);
  return new TextDecoder().decode(getPart(doc.opc, "/word/document.xml")?.data);
}

function childOrder(xml: string, parent: string): string[] {
  const body = new RegExp(`<w:${parent}[^>]*>(.*?)</w:${parent}>`, "s").exec(xml)?.[1] ?? "";
  return [...body.matchAll(/<w:(\w+)/g)]
    .map((m) => m[1] ?? "")
    .filter((n, i, a) => a.indexOf(n) === i);
}

describe("section properties", () => {
  it("round-trips columns, line numbers, page borders, grid and flags in schema order", () => {
    const doc = createDocx({ paragraphs: ["a"] });
    setSectionProperties(doc, {
      textDirection: "tbRl",
      titlePage: true,
      verticalAlignment: "center",
      documentGrid: { type: "lines", linePitch: 360 },
      columns: {
        count: 2,
        spaceTwips: 720,
        separator: true,
        columns: [{ widthTwips: 2640, spaceTwips: 720 }, { widthTwips: 6000 }],
      },
      lineNumbering: { start: 5, countBy: 2, restart: "newSection", distanceTwips: 360 },
      pageBorders: {
        offsetFrom: "page",
        display: "notFirstPage",
        zOrder: "back",
        top: { style: "apples", size: 20, spacePt: 24, color: "auto" },
        bottom: { style: "double", size: 6, spacePt: 24, color: "FF0000", shadow: true },
      },
      start: "oddPage",
    });
    setPageMargins(doc, { ...MARGINS_NORMAL, gutter: 360 });
    const xml = documentXml(doc);
    expect(childOrder(xml, "sectPr")).toEqual([
      "type",
      "pgSz",
      "pgMar",
      "pgBorders",
      "top",
      "bottom",
      "lnNumType",
      "cols",
      "col",
      "vAlign",
      "titlePg",
      "textDirection",
      "docGrid",
    ]);
    expect(xml).toContain(
      '<w:lnNumType w:countBy="2" w:start="4" w:distance="360" w:restart="newSection"/>',
    );
    expect(xml).toContain('<w:cols w:num="2" w:space="720" w:sep="1" w:equalWidth="0">');
    const back = getSectionProperties(reopen(doc));
    expect(back.start).toBe("oddPage");
    expect(back.columns).toEqual({
      count: 2,
      spaceTwips: 720,
      separator: true,
      columns: [{ widthTwips: 2640, spaceTwips: 720 }, { widthTwips: 6000 }],
    });
    expect(back.lineNumbering).toEqual({
      start: 5,
      countBy: 2,
      restart: "newSection",
      distanceTwips: 360,
    });
    expect(back.pageBorders?.top).toEqual({
      style: "apples",
      size: 20,
      spacePt: 24,
      color: "auto",
    });
    expect(back.pageBorders?.bottom?.shadow).toBe(true);
    expect(back.pageBorders?.display).toBe("notFirstPage");
    expect(back.pageBorders?.offsetFrom).toBe("page");
    expect(back.pageBorders?.zOrder).toBe("back");
    expect(back.verticalAlignment).toBe("center");
    expect(back.titlePage).toBe(true);
    expect(back.textDirection).toBe("tbRl");
    expect(back.documentGrid).toEqual({ type: "lines", linePitch: 360 });
    expect(back.margins.gutter).toBe(360);
  });

  it("removes optional elements with null and defaults with their default value", () => {
    const doc = createDocx();
    setSectionProperties(doc, {
      lineNumbering: { start: 1, countBy: 1, restart: "continuous" },
      verticalAlignment: "bottom",
      start: "continuous",
    });
    setSectionProperties(doc, { lineNumbering: null, verticalAlignment: "top", start: "nextPage" });
    const xml = documentXml(doc);
    expect(xml).not.toContain("lnNumType");
    expect(xml).not.toContain("vAlign");
    expect(xml).not.toContain("w:type");
  });

  it("rejects invalid values without changing any section", () => {
    const doc = createDocx();
    expect(() =>
      setSectionProperties(doc, { columns: { count: 0, spaceTwips: 720, separator: false } }),
    ).toThrow(RangeError);
    expect(() =>
      setSectionProperties(doc, {
        pageBorders: {
          offsetFrom: "text",
          display: "allPages",
          zOrder: "front",
          top: { style: "nope", size: 4, spacePt: 1, color: "auto" },
        },
      }),
    ).toThrow(RangeError);
    expect(getSectionProperties(doc).columns.count).toBe(1);
  });

  it("inserts section breaks that copy the section and give the new one its start type", () => {
    const doc = createDocx({ paragraphs: ["one", "two", "three"] });
    setPageSize(doc, PAGE_SIZE_A4);
    expect(sectionCount(doc)).toBe(1);
    const created = insertSectionBreak(doc, 0, "continuous");
    expect(created).toBe(1);
    expect(sectionCount(doc)).toBe(2);
    expect(sectionIndexAt(doc, 0)).toBe(0);
    expect(sectionIndexAt(doc, 1)).toBe(1);
    setSectionProperties(doc, { columns: { count: 3, spaceTwips: 720, separator: false } }, 1);
    setPageOrientation(doc, "landscape", "all");
    const back = reopen(doc);
    expect(sectionCount(back)).toBe(2);
    expect(getSectionProperties(back, 0).pageSize.widthTwips).toBe(PAGE_SIZE_A4.heightTwips);
    expect(getSectionProperties(back, 0).columns.count).toBe(1);
    expect(getSectionProperties(back, 0).start).toBe("nextPage");
    expect(getSectionProperties(back, 1).columns.count).toBe(3);
    expect(getSectionProperties(back, 1).start).toBe("continuous");
    expect(getSectionProperties(back, 1).pageSize.orientation).toBe("landscape");
  });

  it("puts a break after a table or an existing break on a new paragraph", () => {
    const doc = createDocx({ paragraphs: ["a"] });
    insertSectionBreak(doc, 0);
    insertSectionBreak(doc, 0);
    expect(doc.document.body.blocks).toHaveLength(2);
    expect(sectionCount(doc)).toBe(3);
  });

  it("writes the paper code and keeps it when the orientation changes", () => {
    const doc = createDocx();
    setPageSize(doc, { ...PAGE_SIZE_A4, paperCode: 9 });
    setPageOrientation(doc, "landscape");
    expect(documentXml(doc)).toContain(
      '<w:pgSz w:w="16838" w:h="11906" w:orient="landscape" w:code="9"/>',
    );
    expect(getSectionProperties(reopen(doc)).pageSize.paperCode).toBe(9);
  });
});

describe("document settings order", () => {
  it("writes settings at their schema position whatever the call order", () => {
    const doc = createDocx();
    setDocumentSettingVal(doc, "hyphenationZone", "357");
    setDocumentSettingOnOff(doc, "evenAndOddHeaders", true);
    setDocumentSettingOnOff(doc, "mirrorMargins", true);
    setDocumentSettingOnOff(doc, "autoHyphenation", true);
    setDocumentSettingOnOff(doc, "displayBackgroundShape", true);
    setDocumentSettingOnOff(doc, "gutterAtTop", true);
    appendParagraph(doc, "x");
    const back = reopen(doc);
    const xml = new TextDecoder().decode(getPart(back.opc, "/word/settings.xml")?.data);
    expect([...xml.matchAll(/<w:(\w+)/g)].map((m) => m[1]).slice(1)).toEqual([
      "displayBackgroundShape",
      "mirrorMargins",
      "gutterAtTop",
      "autoHyphenation",
      "hyphenationZone",
      "evenAndOddHeaders",
    ]);
    expect(getDocumentSetting(back, "hyphenationZone")).toEqual({ present: true, val: "357" });
  });
});
