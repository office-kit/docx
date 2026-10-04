import { strFromU8, unzipSync } from "fflate";
import { describe, expect, it } from "vitest";
import {
  createDocx,
  getRunFormat,
  openDocx,
  paragraphs,
  setRunFormat,
  setRunValProp,
  toUint8Array,
  type WmlRun,
} from "./index.js";

// `w:sz/@w:val` is ST_HpsMeasure: a half-point count (ST_UnsignedDecimalNumber)
// or a positive universal measure such as "12pt" (ST_PositiveUniversalMeasure,
// units mm / cm / in / pt / pc / pi). getRunFormat reports half-points.
function runWithSize(val: string): WmlRun {
  const doc = createDocx({ paragraphs: ["Hi"] });
  const run = paragraphs(doc)[0]!.children[0]!;
  if (run.kind !== "run") throw new Error("expected a run");
  setRunValProp(run, "sz", val);
  return run;
}

describe("getRunFormat font size", () => {
  it.each([
    ["24", 24],
    ["0", 0],
    ["12pt", 24],
    ["10.5pt", 21],
    ["1in", 144],
    ["2.54cm", 144],
    ["25.4mm", 144],
    ["1pc", 24],
    ["1pi", 24],
  ])("w:sz=%s reads as %s half-points", (val, halfPoints) => {
    expect(getRunFormat(runWithSize(val)).fontSizeHalfPoints).toBe(halfPoints);
  });

  // Not a whole number of half-points, or not a valid ST_HpsMeasure: the typed
  // value is left unset rather than guessed (parseInt read "12.3pt" as 12).
  it.each(["12.3pt", "3mm", "-4", "12px", "abc", ""])("w:sz=%j has no typed size", (val) => {
    expect(getRunFormat(runWithSize(val))).not.toHaveProperty("fontSizeHalfPoints");
  });

  it("keeps a unit size in the XML when other formatting is edited", () => {
    const doc = createDocx({ paragraphs: ["Hi"] });
    const run = paragraphs(doc)[0]!.children[0]!;
    if (run.kind !== "run") throw new Error("expected a run");
    setRunValProp(run, "sz", "3mm");
    setRunFormat(run, { bold: true });
    const xml = strFromU8(unzipSync(toUint8Array(doc))["word/document.xml"]!);
    expect(xml).toContain('<w:sz w:val="3mm"/>');
    const reopened = paragraphs(openDocx(toUint8Array(doc)))[0]!.children[0]!;
    expect(reopened.kind === "run" && getRunFormat(reopened)).toEqual({ bold: true });
  });
});
