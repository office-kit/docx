import { strFromU8, unzipSync } from "fflate";
import { describe, expect, it } from "vitest";
import {
  addTable,
  appendTextRun,
  createDocx,
  getRunFormat,
  getTableCellText,
  HIGHLIGHT_COLORS,
  openDocx,
  paragraphs,
  paragraphText,
  setParagraphText,
  setRunFormat,
  setRunValProp,
  setTableCellText,
  tables,
  toUint8Array,
  validate,
  type WmlRun,
} from "./index.js";

// ST_HighlightColor (ECMA-376 Part 1 §17.18.40), as listed in wml.xsd.
const SCHEMA_VALUES = [
  "black",
  "blue",
  "cyan",
  "green",
  "magenta",
  "red",
  "yellow",
  "white",
  "darkBlue",
  "darkCyan",
  "darkGreen",
  "darkMagenta",
  "darkRed",
  "darkYellow",
  "darkGray",
  "lightGray",
  "none",
];

const INVALID = ["FFFF00", "#ffff00", "Yellow", "transparent", ""];

function firstRun(doc: ReturnType<typeof createDocx>): WmlRun {
  const run = paragraphs(doc)[0]?.children.find((c) => c.kind === "run");
  if (run?.kind !== "run") throw new Error("no run in paragraph");
  return run;
}

function documentXml(doc: ReturnType<typeof createDocx>): string {
  const part = unzipSync(toUint8Array(doc))["word/document.xml"];
  if (!part) throw new Error("word/document.xml missing");
  return strFromU8(part);
}

describe("highlight writers", () => {
  it("HIGHLIGHT_COLORS is exactly ST_HighlightColor", () => {
    expect([...HIGHLIGHT_COLORS].toSorted()).toEqual([...SCHEMA_VALUES].toSorted());
  });

  it.each([...HIGHLIGHT_COLORS])("setRunFormat writes w:val=%s and it reopens", (color) => {
    const doc = createDocx({ paragraphs: ["Hi"] });
    setRunFormat(firstRun(doc), { highlight: color });
    expect(documentXml(doc)).toContain(`<w:highlight w:val="${color}"/>`);
    const reopened = openDocx(toUint8Array(doc));
    expect(getRunFormat(firstRun(reopened)).highlight).toBe(color);
    expect(validate(reopened)).toHaveLength(0);
  });

  it("appendTextRun writes a named highlight", () => {
    const doc = createDocx({ paragraphs: ["Hi"] });
    appendTextRun(paragraphs(doc)[0]!, " there", { highlight: "darkGreen" });
    expect(documentXml(doc)).toContain('<w:highlight w:val="darkGreen"/>');
  });

  describe.each(INVALID)("rejects %j before changing anything", (bad) => {
    it("setRunFormat leaves the run (and its other props) untouched", () => {
      const doc = createDocx({ paragraphs: ["Hi"] });
      const run = firstRun(doc);
      setRunFormat(run, { italic: true });
      const before = structuredClone(run);
      expect(() => setRunFormat(run, { bold: true, highlight: bad })).toThrow(RangeError);
      expect(run).toEqual(before);
    });

    it("setRunFormat does not create an rPr on a plain run", () => {
      const doc = createDocx({ paragraphs: ["Hi"] });
      const run = firstRun(doc);
      expect(() => setRunFormat(run, { highlight: bad })).toThrow(RangeError);
      expect(run.rPr).toBeUndefined();
    });

    it("appendTextRun adds no run", () => {
      const doc = createDocx({ paragraphs: ["Hi"] });
      const para = paragraphs(doc)[0]!;
      const before = structuredClone(para);
      expect(() => appendTextRun(para, "x", { bold: true, highlight: bad })).toThrow(RangeError);
      expect(para).toEqual(before);
    });

    it("setParagraphText keeps the old text", () => {
      const doc = createDocx({ paragraphs: ["Hi"] });
      const para = paragraphs(doc)[0]!;
      expect(() => setParagraphText(para, "new", { highlight: bad })).toThrow(RangeError);
      expect(paragraphText(para)).toBe("Hi");
    });

    it("setTableCellText keeps the old cell", () => {
      const doc = createDocx({ paragraphs: ["Hi"] });
      addTable(doc, [["old"]]);
      const table = tables(doc)[0]!;
      expect(() => setTableCellText(table, 0, 0, "new", { highlight: bad })).toThrow(RangeError);
      expect(getTableCellText(table, 0, 0)).toBe("old");
    });
  });
});

describe("reading a highlight another tool wrote", () => {
  it("returns the raw value and keeps it through save → reopen", () => {
    const doc = createDocx({ paragraphs: ["Hi"] });
    // The generic val-prop setter writes any value, like a foreign producer.
    setRunValProp(firstRun(doc), "highlight", "FFFF00");
    const reopened = openDocx(toUint8Array(doc));
    expect(getRunFormat(firstRun(reopened)).highlight).toBe("FFFF00");
    expect(documentXml(reopened)).toContain('<w:highlight w:val="FFFF00"/>');
  });

  it("editing other props of that run leaves the raw highlight as it was", () => {
    const doc = createDocx({ paragraphs: ["Hi"] });
    setRunValProp(firstRun(doc), "highlight", "FFFF00");
    setRunFormat(firstRun(doc), { bold: true });
    expect(getRunFormat(firstRun(doc))).toMatchObject({ bold: true, highlight: "FFFF00" });
  });
});
