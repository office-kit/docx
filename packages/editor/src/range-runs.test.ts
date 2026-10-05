/**
 * `runsInRange` returns "every run touched by an ordered selection": the runs
 * whose characters the range covers, not every run of the end paragraphs. The
 * ribbon builds its none / mixed / value state from these runs.
 */

import {
  appendTextRun,
  createDocx,
  getRunFormat,
  paragraphs,
  runTextLength,
  type WmlRun,
} from "@office-kit/docx";
import { describe, expect, it } from "vitest";
import { type DocPosition, orderSelection, runsInRange } from "./index.js";

/** Paragraphs of runs "Aa" | "Bb" (bold) | "Cc" (Times). */
function threeRunDoc(paragraphCount = 1): ReturnType<typeof createDocx> {
  const doc = createDocx({ paragraphs: Array.from({ length: paragraphCount }, () => "Aa") });
  for (const p of paragraphs(doc)) {
    appendTextRun(p, "Bb", { bold: true });
    appendTextRun(p, "Cc", { font: "Times New Roman" });
  }
  return doc;
}

const at = (block: number, inline: number, offset: number): DocPosition => ({
  block,
  inline,
  offset,
});

function texts(runs: WmlRun[]): string[] {
  return runs.map((r) => r.pieces.map((p) => (p.kind === "text" ? p.value : "")).join(""));
}

describe("runsInRange", () => {
  it("returns only the runs the range covers inside one paragraph", () => {
    const doc = threeRunDoc();
    // From the end of "Aa" to the middle of "Bb": only "Bb" is covered.
    const runs = runsInRange(doc, orderSelection({ anchor: at(0, 0, 2), focus: at(0, 1, 1) }));
    expect(texts(runs)).toEqual(["Bb"]);
  });

  it("includes a differently formatted middle run, so the range reads as mixed", () => {
    const doc = threeRunDoc();
    const runs = runsInRange(doc, orderSelection({ anchor: at(0, 0, 1), focus: at(0, 2, 1) }));
    expect(texts(runs)).toEqual(["Aa", "Bb", "Cc"]);
    expect(runs.map((r) => getRunFormat(r).bold ?? false)).toEqual([false, true, false]);
  });

  it("across paragraphs, leaves out the runs before the start and after the end", () => {
    const doc = threeRunDoc(2);
    // From inside "Cc" of paragraph 0 to inside "Aa" of paragraph 1.
    const runs = runsInRange(doc, orderSelection({ anchor: at(1, 0, 1), focus: at(0, 2, 1) }));
    expect(texts(runs)).toEqual(["Cc", "Aa"]);
  });

  it("a collapsed caret still returns the caret paragraph's runs (documented)", () => {
    const doc = threeRunDoc();
    const runs = runsInRange(doc, orderSelection({ anchor: at(0, 1, 1), focus: at(0, 1, 1) }));
    expect(texts(runs)).toEqual(["Aa", "Bb", "Cc"]);
    expect(runs.every((r) => runTextLength(r) === 2)).toBe(true);
  });
});
