import { strFromU8, unzipSync } from "fflate";
import { describe, expect, it } from "vitest";
import {
  addTable,
  appendTextRun,
  createDocx,
  getRunFormat,
  getTableCellText,
  openDocx,
  paragraphs,
  paragraphText,
  type RunFormatting,
  setParagraphText,
  setRunFormat,
  setRunValProp,
  setTableCellText,
  tables,
  toUint8Array,
  validate,
  type WmlRun,
} from "./index.js";

// `w:color/@w:val` is ST_HexColor: "auto" or six hex digits (hexBinary, either
// case). `w:sz/@w:val` as a number is ST_UnsignedDecimalNumber: a non-negative
// integer, 0 included, with no small upper bound.
const VALID: RunFormatting[] = [
  { color: "FF0000" },
  { color: "00ff7f" },
  { color: "auto" },
  { fontSizeHalfPoints: 0 },
  { fontSizeHalfPoints: 1 },
  { fontSizeHalfPoints: 3276 },
  { fontSizeHalfPoints: Number.MAX_SAFE_INTEGER },
];

const INVALID: RunFormatting[] = [
  { color: "#FF0000" },
  { color: "red" },
  { color: "FFF" },
  { color: "FF00001" },
  { color: "Auto" },
  { color: "" },
  { fontSizeHalfPoints: -1 },
  { fontSizeHalfPoints: 23.5 },
  { fontSizeHalfPoints: Number.NaN },
  { fontSizeHalfPoints: Number.POSITIVE_INFINITY },
  { fontSizeHalfPoints: Number.MAX_SAFE_INTEGER + 1 },
];

// Untyped JS input (e.g. parsed JSON): the wrong runtime type must be rejected,
// not stringified — RegExp#test would let the number 123456 through.
const WRONG_RUNTIME_TYPES: RunFormatting[] = JSON.parse(
  '[{"color":123456},{"color":null},{"color":{}},{"color":true},' +
    '{"fontSizeHalfPoints":"24"},{"fontSizeHalfPoints":null},{"fontSizeHalfPoints":[24]}]',
);

function firstRun(doc: ReturnType<typeof createDocx>): WmlRun {
  const run = paragraphs(doc)[0]?.children.find((c) => c.kind === "run");
  if (run?.kind !== "run") throw new Error("no run in paragraph");
  return run;
}

function rPrTags(doc: ReturnType<typeof createDocx>): string[] {
  const part = unzipSync(toUint8Array(doc))["word/document.xml"];
  if (!part) throw new Error("word/document.xml missing");
  return strFromU8(part).match(/<w:(?:color|sz|szCs)\b[^>]*>/g) ?? [];
}

describe("color / font size writers: valid values", () => {
  it.each(VALID)("setRunFormat(%j) is written and reopens unchanged", (fmt) => {
    const doc = createDocx({ paragraphs: ["Hi"] });
    setRunFormat(firstRun(doc), fmt);
    const tags = rPrTags(doc);
    if (fmt.color !== undefined) expect(tags).toEqual([`<w:color w:val="${fmt.color}"/>`]);
    if (fmt.fontSizeHalfPoints !== undefined) {
      const v = String(fmt.fontSizeHalfPoints);
      expect(tags).toEqual([`<w:sz w:val="${v}"/>`, `<w:szCs w:val="${v}"/>`]);
    }
    const reopened = openDocx(toUint8Array(doc));
    expect(getRunFormat(firstRun(reopened))).toEqual(fmt);
    expect(validate(reopened)).toHaveLength(0);
  });

  it("appendTextRun writes them too", () => {
    const doc = createDocx({ paragraphs: ["Hi"] });
    appendTextRun(paragraphs(doc)[0]!, "!", { color: "auto", fontSizeHalfPoints: 0 });
    expect(rPrTags(doc)).toEqual([
      '<w:color w:val="auto"/>',
      '<w:sz w:val="0"/>',
      '<w:szCs w:val="0"/>',
    ]);
  });
});

describe.each([...INVALID, ...WRONG_RUNTIME_TYPES])(
  "color / font size writers reject %j before changing anything",
  (bad) => {
    it("setRunFormat leaves the run and its other props untouched", () => {
      const doc = createDocx({ paragraphs: ["Hi"] });
      const run = firstRun(doc);
      setRunFormat(run, { italic: true });
      const before = structuredClone(run);
      expect(() => setRunFormat(run, { bold: true, ...bad })).toThrow(RangeError);
      expect(run).toEqual(before);
    });

    it("setRunFormat does not create an rPr on a plain run", () => {
      const doc = createDocx({ paragraphs: ["Hi"] });
      expect(() => setRunFormat(firstRun(doc), bad)).toThrow(RangeError);
      expect(firstRun(doc).rPr).toBeUndefined();
    });

    it("appendTextRun adds no run", () => {
      const doc = createDocx({ paragraphs: ["Hi"] });
      const para = paragraphs(doc)[0]!;
      const before = structuredClone(para);
      expect(() => appendTextRun(para, "x", { bold: true, ...bad })).toThrow(RangeError);
      expect(para).toEqual(before);
    });

    it("setParagraphText keeps the old text", () => {
      const doc = createDocx({ paragraphs: ["Hi"] });
      const para = paragraphs(doc)[0]!;
      expect(() => setParagraphText(para, "new", bad)).toThrow(RangeError);
      expect(paragraphText(para)).toBe("Hi");
    });

    it("setTableCellText keeps the old cell", () => {
      const doc = createDocx({ paragraphs: ["Hi"] });
      addTable(doc, [["old"]]);
      const table = tables(doc)[0]!;
      expect(() => setTableCellText(table, 0, 0, "new", bad)).toThrow(RangeError);
      expect(getTableCellText(table, 0, 0)).toBe("old");
    });
  },
);

describe("reading color / size another tool wrote", () => {
  it("returns the raw color and keeps it through save → reopen and unrelated edits", () => {
    const doc = createDocx({ paragraphs: ["Hi"] });
    setRunValProp(firstRun(doc), "color", "red");
    setRunFormat(firstRun(doc), { bold: true });
    const reopened = openDocx(toUint8Array(doc));
    expect(getRunFormat(firstRun(reopened))).toMatchObject({ bold: true, color: "red" });
    expect(rPrTags(reopened)).toEqual(['<w:color w:val="red"/>']);
  });

  it("keeps a universal-measure size (12pt) in the file untouched", () => {
    const doc = createDocx({ paragraphs: ["Hi"] });
    setRunValProp(firstRun(doc), "sz", "12pt");
    setRunFormat(firstRun(doc), { italic: true });
    expect(rPrTags(openDocx(toUint8Array(doc)))).toEqual(['<w:sz w:val="12pt"/>']);
  });
});
