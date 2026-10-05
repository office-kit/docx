import {
  addFootnote,
  appendParagraph,
  appendSectionBreak,
  createDocx,
  ensureHeaderFooter,
  makePropsElement,
  setSectionProperties,
  setElementAttr,
  setElementOnOff,
  setElementValProp,
  type XmlElement,
} from "@office-kit/docx";
import { describe, expect, it } from "vitest";
import { collectNoteReferences, numberNotes } from "./notes.js";
import { formatNumber } from "./number-format.js";
import { documentSections, layoutSettings } from "./sections.js";

function addChild(parent: XmlElement, local: string): XmlElement {
  const el = makePropsElement(local);
  (parent.children as XmlElement[]).push(el);
  return el;
}

describe("documentSections", () => {
  it("reads defaults for a bare document", () => {
    const [only] = documentSections(createDocx());
    expect(only).toMatchObject({
      pageWidth: 12240,
      pageHeight: 15840,
      start: "nextPage",
      columns: [{ width: 9360, space: 0 }],
      vAlign: "top",
      titlePage: false,
    });
  });

  it("divides the text height into columns for vertical text", () => {
    const doc = createDocx();
    setSectionProperties(doc, {
      textDirection: "tbRl",
      columns: { count: 2, spaceTwips: 720, separator: false },
    });
    const [only] = documentSections(doc);
    // Letter, 1 in margins: 12960 twips of text height, less the gap, halved.
    expect(only).toMatchObject({ vertical: true, columns: [{ width: 6120 }, { width: 6120 }] });
  });

  it("splits blocks into sections and reads their properties", () => {
    const doc = createDocx({ paragraphs: ["one"] });
    appendSectionBreak(doc, "nextPage");
    appendParagraph(doc, "two");
    const sectPr = makePropsElement("sectPr");
    doc.document.body.sectPr = sectPr;
    setElementValProp(sectPr, "type", "continuous");
    const cols = addChild(sectPr, "cols");
    setElementAttr(cols, "num", "2");
    setElementAttr(cols, "space", "720");
    setElementAttr(cols, "sep", "1");
    setElementOnOff(sectPr, "titlePg", true);
    setElementValProp(sectPr, "vAlign", "center");
    const ln = addChild(sectPr, "lnNumType");
    setElementAttr(ln, "countBy", "5");
    setElementAttr(ln, "restart", "continuous");
    const borders = addChild(sectPr, "pgBorders");
    setElementAttr(borders, "offsetFrom", "page");
    const top = addChild(borders, "top");
    setElementAttr(top, "val", "double");
    setElementAttr(top, "sz", "12");
    const pgNum = addChild(sectPr, "pgNumType");
    setElementAttr(pgNum, "fmt", "lowerRoman");
    setElementAttr(pgNum, "start", "3");

    const sections = documentSections(doc);
    expect(sections.map((s) => [s.firstBlock, s.lastBlock])).toEqual([
      [0, 1],
      [2, 2],
    ]);
    const last = sections[1];
    expect(last?.start).toBe("continuous");
    expect(last?.columns).toEqual([
      { width: 4320, space: 720 },
      { width: 4320, space: 0 },
    ]);
    expect(last?.separator).toBe(true);
    expect(last?.titlePage).toBe(true);
    expect(last?.vAlign).toBe("center");
    expect(last?.lineNumbers).toEqual({
      countBy: 5,
      start: 1,
      restart: "continuous",
      distance: undefined,
    });
    expect(last?.borders?.offsetFrom).toBe("page");
    expect(last?.borders?.top).toEqual({ style: "double", size: 12, space: 0, color: "auto" });
    expect(last?.pageNumbers).toEqual({ format: "lowerRoman", start: 3 });
  });

  it("inherits headers from the previous section", () => {
    const doc = createDocx({ paragraphs: ["one"] });
    appendSectionBreak(doc, "nextPage");
    appendParagraph(doc, "two");
    const part = ensureHeaderFooter(doc, 0, "header", "default");
    const sections = documentSections(doc);
    expect(sections.map((s) => s.headers.default)).toEqual([part, part]);
    expect(sections[0]?.headers.first).toBeUndefined();
  });
});

describe("layoutSettings", () => {
  it("reads evenAndOddHeaders", () => {
    const doc = createDocx();
    expect(layoutSettings(doc).evenAndOddHeaders).toBe(false);
  });
});

describe("formatNumber", () => {
  it.each([
    [4, "upperRoman", "IV"],
    [14, "lowerRoman", "xiv"],
    [28, "upperLetter", "BB"],
    [3, "lowerLetter", "c"],
    [2, "ordinal", "2nd"],
    [12, "ordinal", "12th"],
    [5, "chicago", "**"],
    [7, "numberInDash", "- 7 -"],
    [7, "decimalZero", "07"],
    [7, "unknownFormat", "7"],
  ])("%i as %s is %s", (n, fmt, text) => {
    expect(formatNumber(n, fmt)).toBe(text);
  });
});

describe("numberNotes", () => {
  it("numbers footnotes in order and restarts per section or page", () => {
    const doc = createDocx({ paragraphs: ["a", "b", "c"] });
    const paras = doc.document.body.blocks.filter((b) => b.kind === "paragraph");
    for (const p of paras) addFootnote(doc, p, "note");
    const refs = collectNoteReferences(doc);
    expect(refs.map((r) => r.block)).toEqual([0, 1, 2]);
    const sections = documentSections(doc);
    const continuous = numberNotes(
      refs,
      sections,
      () => 0,
      (b) => (b < 2 ? 0 : 1),
    );
    expect([...continuous.values()]).toEqual(["1", "2", "3"]);

    const [only] = sections;
    if (!only) throw new Error("no section");
    const eachPage = numberNotes(
      refs,
      [{ ...only, footnotePr: { numRestart: "eachPage", numFmt: "lowerLetter" } }],
      () => 0,
      (b) => (b < 2 ? 0 : 1),
    );
    expect([...eachPage.values()]).toEqual(["a", "b", "a"]);
  });
});
