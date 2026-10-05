import { describe, expect, it } from "vitest";
import { getPart } from "../internal/opc/index.js";
import { compareDocuments } from "./compare.js";
import {
  acceptAllRevisions,
  addTable,
  createDocx,
  openDocx,
  rejectAllRevisions,
  setParagraphValProp,
  paragraphs,
  text,
  toUint8Array,
} from "./index.js";
import { revisions } from "./review.js";
import { validatePackage } from "./validator.js";

const WHO = { author: "Compare", date: "2026-10-05T00:00:00Z" };

describe("compareDocuments", () => {
  it("marks changed words, added and removed paragraphs", () => {
    const original = createDocx({ paragraphs: ["The quick brown fox", "Removed line", "Same"] });
    const revised = createDocx({ paragraphs: ["The slow brown fox jumps", "Same", "Added line"] });
    const result = compareDocuments(original, revised, WHO);
    expect(revisions(result).map((r) => [r.kind, r.text])).toEqual([
      ["delete", "quick"],
      ["insert", "slow"],
      ["insert", " jumps"],
      ["delete", "Removed line"],
      ["paragraphDelete", ""],
      // The mark before a paragraph appended at the end is the inserted one.
      ["paragraphInsert", ""],
      ["insert", "Added line"],
    ]);
    const reopened = openDocx(toUint8Array(result));
    expect(validatePackage(reopened.opc)).toEqual([]);
    const accepted = openDocx(toUint8Array(result));
    acceptAllRevisions(accepted);
    expect(text(accepted)).toBe(text(revised));
    const rejected = openDocx(toUint8Array(result));
    rejectAllRevisions(rejected);
    expect(text(rejected)).toBe(text(original));
    // The inputs are untouched.
    expect(text(original)).toBe("The quick brown fox\nRemoved line\nSame");
  });

  it("compares Japanese text character by character", () => {
    const result = compareDocuments(
      createDocx({ paragraphs: ["今日は晴れです"] }),
      createDocx({ paragraphs: ["今日は雨です"] }),
      WHO,
    );
    expect(revisions(result).map((r) => [r.kind, r.text])).toEqual([
      ["delete", "晴れ"],
      ["insert", "雨"],
    ]);
  });

  it("records paragraph property changes", () => {
    const original = createDocx({ paragraphs: ["x"] });
    const revised = createDocx({ paragraphs: ["x"] });
    const para = paragraphs(revised)[0];
    if (!para) throw new Error("paragraph");
    setParagraphValProp(para, "jc", "center");
    const result = compareDocuments(original, revised, WHO);
    const xml = new TextDecoder().decode(
      getPart(openDocx(toUint8Array(result)).opc, "/word/document.xml")?.data,
    );
    expect(xml).toContain('<w:jc w:val="center"/><w:pPrChange w:id="0" w:author="Compare"');
  });

  it("compares same-shaped tables cell by cell", () => {
    const original = createDocx({ paragraphs: [] });
    addTable(original, [["a", "b"]]);
    const revised = createDocx({ paragraphs: [] });
    addTable(revised, [["a", "c"]]);
    const result = compareDocuments(original, revised, WHO);
    expect(revisions(result).map((r) => [r.kind, r.text])).toEqual([
      ["delete", "b"],
      ["insert", "c"],
    ]);
    expect(validatePackage(openDocx(toUint8Array(result)).opc)).toEqual([]);
  });
});
