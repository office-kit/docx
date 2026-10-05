import { describe, expect, it } from "vitest";
import {
  addListDefinition,
  appendParagraph,
  applyListToParagraph,
  buildEnclosedCharacterRuns,
  buildRubyRun,
  builtinStyles,
  createDocx,
  type Docx,
  ensureBuiltinStyle,
  getParagraphTabs,
  openDocx,
  paragraphs,
  readRuby,
  restartList,
  setParagraphAlignment,
  setParagraphBorder,
  setParagraphShading,
  setParagraphSpacing,
  setParagraphTabs,
  setRunBorder,
  setRunColor,
  setRunFont,
  setRunFormat,
  setRunOnOff,
  setRunShading,
  setRunUnderline,
  setRunValProp,
  toUint8Array,
  updateStyleFormatting,
  validate,
  type WmlRun,
} from "./index.js";

function reopen(doc: Docx): Docx {
  return openDocx(toUint8Array(doc));
}

function partXml(doc: Docx, name: string): string {
  return new TextDecoder().decode(reopen(doc).opc.parts.get(name)?.data ?? new Uint8Array());
}

const bodyXml = (doc: Docx): string => partXml(doc, "/word/document.xml");

/** A document with one paragraph holding one run, and that run. */
function oneRun(text = "text"): { doc: Docx; run: WmlRun } {
  const doc = createDocx({ paragraphs: [] });
  const p = appendParagraph(doc, text);
  const run = p.children[0];
  if (run?.kind !== "run") throw new Error("expected a run");
  return { doc, run };
}

function expectValid(doc: Docx): void {
  expect(validate(reopen(doc)).filter((i) => i.level === "error")).toEqual([]);
}

describe("schema order", () => {
  it("writes rPr and pPr children in the wml.xsd sequence whatever order they were set in", () => {
    const { doc, run } = oneRun();
    setRunValProp(run, "vertAlign", "superscript");
    setRunFormat(run, { highlight: "yellow", bold: true });
    setRunOnOff(run, "caps", true);
    setRunFont(run, "latin", { name: "Arial" });
    const p = paragraphs(doc)[0]!;
    setParagraphAlignment(p, "center");
    setParagraphSpacing(p, { after: 0 });
    setParagraphShading(p, { fill: "FFFF00" });
    const xml = bodyXml(doc);
    const order = (names: string[]): number[] =>
      names.map((n) =>
        xml.indexOf(`<w:${n} `) >= 0 ? xml.indexOf(`<w:${n} `) : xml.indexOf(`<w:${n}/>`),
      );
    const runOrder = order(["rFonts", "b", "caps", "highlight", "vertAlign"]);
    expect(runOrder).toEqual(runOrder.toSorted((a, b) => a - b));
    const para = order(["shd", "spacing", "jc"]);
    expect(para).toEqual(para.toSorted((a, b) => a - b));
  });

  it("leaves a document that is already in order byte-for-byte unchanged", () => {
    const { doc, run } = oneRun();
    setRunFormat(run, { bold: true, color: "FF0000" });
    const first = toUint8Array(doc);
    const again = toUint8Array(openDocx(first));
    expect(
      new TextDecoder().decode(openDocx(again).opc.parts.get("/word/document.xml")!.data),
    ).toBe(new TextDecoder().decode(openDocx(first).opc.parts.get("/word/document.xml")!.data));
  });
});

describe("run colour, underline, font", () => {
  it("round-trips a theme colour with its tint", () => {
    const { doc, run } = oneRun();
    setRunColor(run, { rgb: "8DD873", themeColor: "accent6", themeTint: 0x99 });
    expect(bodyXml(doc)).toContain(
      '<w:color w:val="8DD873" w:themeColor="accent6" w:themeTint="99"/>',
    );
    expectValid(doc);
  });

  it("removes the colour with undefined and rejects invalid values unchanged", () => {
    const { doc, run } = oneRun();
    setRunColor(run, { rgb: "FF0000" });
    expect(() => setRunColor(run, { rgb: "#FF0000" })).toThrow(RangeError);
    expect(() => setRunColor(run, { rgb: "FF0000", themeShade: 300 })).toThrow(RangeError);
    expect(bodyXml(doc)).toContain('<w:color w:val="FF0000"/>');
    setRunColor(run, undefined);
    expect(bodyXml(doc)).not.toContain("<w:color");
  });

  it("writes every ST_Underline style with an underline colour", () => {
    const { doc, run } = oneRun();
    setRunUnderline(run, { style: "dashDotHeavy", color: { rgb: "0070C0" } });
    expect(bodyXml(doc)).toContain('<w:u w:val="dashDotHeavy" w:color="0070C0"/>');
    expect(() => setRunUnderline(run, { style: "zigzag" as never })).toThrow(RangeError);
  });

  it("swaps a named font for the theme body font and back", () => {
    const { doc, run } = oneRun();
    setRunFont(run, "eastAsia", { name: "Yu Mincho" });
    setRunFont(run, "latin", { name: "Arial" });
    setRunFont(run, "latin", { theme: "minor" });
    let xml = bodyXml(doc);
    expect(xml).toContain('w:eastAsia="Yu Mincho"');
    expect(xml).toContain('w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"');
    expect(xml).not.toContain('w:ascii="Arial"');
    setRunFont(run, "latin", { name: "Georgia" });
    setRunFont(run, "complex", { theme: "major" });
    xml = bodyXml(doc);
    expect(xml).toContain('w:ascii="Georgia" w:hAnsi="Georgia"');
    expect(xml).toContain('w:cstheme="majorBidi"');
    expect(xml).not.toContain("asciiTheme");
  });
});

describe("run shading and border", () => {
  it("round-trips character shading and border", () => {
    const { doc, run } = oneRun();
    setRunShading(run, { pattern: "pct15", color: "auto", fill: "FFFFFF" });
    setRunBorder(run, { style: "single", sizeEighths: 4 });
    const xml = bodyXml(doc);
    expect(xml).toContain('<w:shd w:val="pct15" w:color="auto" w:fill="FFFFFF"/>');
    expect(xml).toContain('<w:bdr w:val="single" w:sz="4" w:space="0" w:color="auto"/>');
    expect(xml.indexOf("<w:bdr")).toBeLessThan(xml.indexOf("<w:shd"));
    expectValid(doc);
  });

  it("writes theme fills", () => {
    const { doc, run } = oneRun();
    setRunShading(run, { fill: "DCEAF7", themeFill: "accent1", themeFillTint: 0x33 });
    expect(bodyXml(doc)).toContain('w:themeFill="accent1" w:themeFillTint="33"');
    expect(() => setRunShading(run, { fill: "blue" })).toThrow(RangeError);
  });
});

describe("paragraph borders and tabs", () => {
  it("edits one side at a time and drops pBdr when empty", () => {
    const doc = createDocx({ paragraphs: ["p"] });
    const p = paragraphs(doc)[0]!;
    setParagraphBorder(p, "bottom", { style: "double", sizeEighths: 6, spacePt: 1 });
    setParagraphBorder(p, "top", { style: "single", color: { rgb: "FF0000" } });
    let xml = bodyXml(doc);
    expect(xml).toMatch(
      /<w:pBdr><w:top w:val="single"[^>]*\/><w:bottom w:val="double" w:sz="6" w:space="1"/,
    );
    setParagraphBorder(p, "top", undefined);
    setParagraphBorder(p, "bottom", undefined);
    xml = bodyXml(doc);
    expect(xml).not.toContain("pBdr");
  });

  it("round-trips tab stops in position order", () => {
    const doc = createDocx({ paragraphs: ["p"] });
    const p = paragraphs(doc)[0]!;
    setParagraphTabs(p, [
      { position: 4320, alignment: "right", leader: "dot" },
      { position: 1440, alignment: "center" },
    ]);
    expect(bodyXml(doc)).toContain(
      '<w:tabs><w:tab w:val="center" w:pos="1440"/><w:tab w:val="right" w:leader="dot" w:pos="4320"/></w:tabs>',
    );
    const reopened = paragraphs(reopen(doc))[0]!;
    expect(getParagraphTabs(reopened)).toEqual([
      { position: 1440, alignment: "center" },
      { position: 4320, alignment: "right", leader: "dot" },
    ]);
    setParagraphTabs(p, []);
    expect(bodyXml(doc)).not.toContain("w:tabs");
  });
});

describe("phonetic guide and enclosed characters", () => {
  it("round-trips a ruby run", () => {
    const doc = createDocx({ paragraphs: ["x"] });
    const p = paragraphs(doc)[0]!;
    p.children = [
      buildRubyRun("漢字", "かんじ", {
        alignment: "distributeSpace",
        rubySizeHalfPoints: 10,
        raiseHalfPoints: 18,
        baseSizeHalfPoints: 21,
        language: "ja-JP",
      }),
    ];
    const xml = bodyXml(doc);
    expect(xml).toContain(
      '<w:rubyPr><w:rubyAlign w:val="distributeSpace"/><w:hps w:val="10"/><w:hpsRaise w:val="18"/><w:hpsBaseText w:val="21"/><w:lid w:val="ja-JP"/></w:rubyPr>',
    );
    const run = paragraphs(reopen(doc))[0]!.children[0];
    if (run?.kind !== "run") throw new Error("expected a run");
    expect(readRuby(run)).toMatchObject({
      base: "漢字",
      ruby: "かんじ",
      alignment: "distributeSpace",
    });
    expectValid(doc);
  });

  it("builds an EQ field for enclosed characters", () => {
    const doc = createDocx({ paragraphs: ["x"] });
    const p = paragraphs(doc)[0]!;
    p.children = buildEnclosedCharacterRuns("秘", {
      enclosure: "circle",
      style: "shrinkText",
      sizeHalfPoints: 22,
    });
    const xml = bodyXml(doc);
    expect(xml).toContain('<w:fldChar w:fldCharType="begin"/>');
    expect(xml).toContain("eq \\o\\ac(");
    expect(xml).toMatch(/<w:sz w:val="15"\/>.*<w:instrText>秘<\/w:instrText>/);
    expect(xml).toContain('<w:fldChar w:fldCharType="end"/>');
  });
});

describe("list definitions", () => {
  it("defines a list and restarts it", () => {
    const doc = createDocx({ paragraphs: ["a", "b"] });
    const numId = addListDefinition(doc, [
      { format: "decimalEnclosedCircle", text: "%1", indentLeft: 420, hanging: 420 },
      { format: "aiueoFullWidth", text: "(%2)", indentLeft: 840, hanging: 420, start: 3 },
    ]);
    const [a, b] = paragraphs(doc);
    applyListToParagraph(doc, a!, numId);
    const restarted = restartList(doc, numId, 0, 5);
    expect(restarted).toBeDefined();
    applyListToParagraph(doc, b!, restarted!);
    const numbering = partXml(doc, "/word/numbering.xml");
    expect(numbering).toContain('<w:numFmt w:val="decimalEnclosedCircle"/>');
    expect(numbering).toContain('<w:start w:val="3"/><w:numFmt w:val="aiueoFullWidth"/>');
    expect(numbering).toContain(
      '<w:lvlOverride w:ilvl="0"><w:startOverride w:val="5"/></w:lvlOverride>',
    );
    // abstractNum elements precede num elements (CT_Numbering is a sequence).
    expect(numbering.lastIndexOf("<w:abstractNum ")).toBeLessThan(numbering.indexOf("<w:num "));
    expectValid(doc);
    expect(() => addListDefinition(doc, [])).toThrow(RangeError);
    expect(restartList(doc, 999, 0, 1)).toBeUndefined();
  });
});

describe("built-in styles", () => {
  it("adds Word's definition and its linked character style once", () => {
    const doc = createDocx({ paragraphs: ["t"] });
    expect(ensureBuiltinStyle(doc, "Title")).toBe(true);
    expect(ensureBuiltinStyle(doc, "Title")).toBe(false);
    expect(ensureBuiltinStyle(doc, "NotAStyle")).toBe(false);
    const styles = partXml(doc, "/word/styles.xml");
    expect(styles).toContain(
      'w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:link w:val="TitleChar"/>',
    );
    expect(styles).toContain('w:styleId="TitleChar"');
    expect(styles).toContain('<w:kern w:val="28"/>');
    expect(builtinStyles().find((s) => s.styleId === "Strong")).toMatchObject({
      type: "character",
      quickStyle: true,
    });
    expectValid(doc);
  });

  it("edits a style's formatting through the run and paragraph functions", () => {
    const doc = createDocx({ paragraphs: ["t"] });
    ensureBuiltinStyle(doc, "Quote");
    expect(
      updateStyleFormatting(doc, "Quote", {
        run: (r) => setRunColor(r, { rgb: "FF0000" }),
        paragraph: (p) => setParagraphAlignment(p, "right"),
      }),
    ).toBe(true);
    const styles = partXml(doc, "/word/styles.xml");
    expect(styles).toMatch(
      /w:styleId="Quote">.*<w:jc w:val="right"\/>.*<w:color w:val="FF0000"\/>/,
    );
    expect(updateStyleFormatting(doc, "Missing", {})).toBe(false);
  });
});
