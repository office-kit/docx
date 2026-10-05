/**
 * Behavior that matches Microsoft Word: style-resolved formatting (what the
 * ribbon shows and the canvas draws), toggles that override a style, Clear
 * All Formatting, indent / line-spacing steps, page geometry, and tables drawn
 * only with the borders the document defines. The Clear All Formatting cases
 * reproduce what Word for Mac 16 wrote for the same input.
 */

import {
  addNumberedList,
  addStyle,
  childElementsOf,
  getElementAttr,
  numberingPart,
  addTable,
  appendHeading,
  appendParagraph,
  createDocx,
  type Docx,
  ensureHeadingStyles,
  getParagraphStyle,
  getRunProp,
  openDocx,
  paragraphs,
  setPageMargins,
  setPageSize,
  setParagraphAlignment,
  setParagraphIndent,
  setParagraphSpacing,
  setRunFormat,
  setRunValProp,
  setTableBorders,
  toUint8Array,
  type WmlParagraph,
  type WmlRun,
  type WmlTable,
  type XmlElement,
} from "@office-kit/docx";
import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import {
  commands,
  createStyleResolver,
  editorFor,
  type EditorModel,
  pageGeometry,
  renderDocumentHtml,
  runCommand,
} from "./index.js";
import { compatibilityMode, resolveTable } from "./table-format.js";
import { WML_NS } from "./wml-ns.js";

function runsOf(para: WmlParagraph): WmlRun[] {
  return para.children.filter((c): c is WmlRun => c.kind === "run");
}

function para(doc: Docx, index: number): WmlParagraph {
  const p = paragraphs(doc)[index];
  if (!p) throw new Error(`no paragraph ${index}`);
  return p;
}

function run(doc: Docx, paraIndex: number, runIndex = 0): WmlRun {
  const r = runsOf(para(doc, paraIndex))[runIndex];
  if (!r) throw new Error(`no run ${runIndex}`);
  return r;
}

/** A heading and a body paragraph, the shape of most edits. */
function headingAndBody(): EditorModel {
  const doc = createDocx({ paragraphs: [] });
  ensureHeadingStyles(doc);
  appendHeading(doc, "Title text", 1);
  appendParagraph(doc, "Body text");
  return editorFor(doc);
}

function caret(model: EditorModel, block: number, offset = 0): void {
  model.setSelection({
    anchor: { block, inline: 0, offset },
    focus: { block, inline: 0, offset },
  });
}

function rPrNames(r: WmlRun): string[] {
  return (r.rPr?.children ?? []).flatMap((c) => (c.kind === "element" ? [c.name.local] : []));
}

function pPrNames(p: WmlParagraph): string[] {
  return (p.pPr?.children ?? []).flatMap((c) => (c.kind === "element" ? [c.name.local] : []));
}

describe("createStyleResolver", () => {
  it("merges document defaults, the paragraph style chain and direct formatting", () => {
    const model = headingAndBody();
    const styles = createStyleResolver(model.doc);
    const heading = styles.run(para(model.doc, 0), run(model.doc, 0));
    // ensureHeadingStyles: Heading1 = bold, 20 pt, 1F497D; the font is the default's.
    expect(heading).toMatchObject({
      bold: true,
      sizeHalfPoints: 40,
      color: "1F497D",
      font: "Calibri",
    });
    const body = styles.run(para(model.doc, 1), run(model.doc, 1));
    expect(body).toMatchObject({ bold: false, sizeHalfPoints: 22, font: "Calibri" });
    expect(body.color).toBeUndefined();
    // pPrDefault: after 160, line 259 auto.
    expect(styles.paragraph(para(model.doc, 1))).toMatchObject({
      after: 160,
      line: 259,
      lineRule: "auto",
    });
  });

  it("follows basedOn, applies character styles, and reads w:val=0 as off", () => {
    const doc = createDocx({ paragraphs: ["x"] });
    addStyle(doc, {
      type: "paragraph",
      styleId: "Base",
      name: "Base",
      bold: true,
      fontSizeHalfPoints: 30,
    });
    addStyle(doc, {
      type: "paragraph",
      styleId: "Derived",
      name: "Derived",
      basedOn: "Base",
      italic: true,
    });
    addStyle(doc, { type: "character", styleId: "Red", name: "Red", color: "FF0000" });
    setRunValProp(run(doc, 0), "rStyle", "Red");
    setRunValProp(run(doc, 0), "b", "0");
    const model = editorFor(doc);
    caret(model, 0);
    runCommand(model, commands.setParagraphStyleCommand, { styleId: "Derived" });
    const resolved = createStyleResolver(model.doc).run(para(model.doc, 0), run(model.doc, 0));
    expect(resolved).toMatchObject({
      italic: true,
      sizeHalfPoints: 30,
      color: "FF0000",
      bold: false,
    });
  });

  it("toggles bold and italic across style levels but not within a basedOn chain (§17.7.3)", () => {
    const doc = createDocx({ paragraphs: ["x"] });
    addStyle(doc, { type: "paragraph", styleId: "Base", name: "Base", bold: true });
    addStyle(doc, {
      type: "paragraph",
      styleId: "Derived",
      name: "Derived",
      basedOn: "Base",
      bold: true,
      italic: true,
    });
    addStyle(doc, { type: "character", styleId: "Strong2", name: "Strong2", bold: true });
    setRunValProp(run(doc, 0), "rStyle", "Strong2");
    const model = editorFor(doc);
    caret(model, 0);
    runCommand(model, commands.setParagraphStyleCommand, { styleId: "Derived" });
    const styles = createStyleResolver(model.doc);
    // Bold from the paragraph style chain (once, despite two styles) and from
    // the character style cancel out; italic is only set at one level.
    expect(styles.run(para(model.doc, 0), run(model.doc, 0))).toMatchObject({
      bold: false,
      italic: true,
    });
    expect(styles.run(para(model.doc, 0))).toMatchObject({ bold: true, italic: true });
    // Direct formatting is absolute.
    setRunValProp(run(model.doc, 0), "b", "1");
    expect(createStyleResolver(model.doc).run(para(model.doc, 0), run(model.doc, 0)).bold).toBe(
      true,
    );
  });

  it("formats a list number with the level's run properties (§17.9.24)", () => {
    const doc = createDocx({ paragraphs: [] });
    addNumberedList(doc, ["one"]);
    const lvl = numberingPart(doc)
      ?.abstractNums.flatMap((a) => childElementsOf(a))
      .find((c) => c.name.local === "lvl" && getElementAttr(c, "ilvl") === "0");
    if (!lvl) throw new Error("no level 0");
    const w = (
      local: string,
      attrs: Record<string, string> = {},
      children: XmlElement[] = [],
    ): XmlElement => ({
      kind: "element",
      name: { uri: WML_NS, local, prefix: "w" },
      attrs: Object.entries(attrs).map(([k, value]) => ({
        name: { uri: WML_NS, local: k, prefix: "w" },
        value,
        isNamespaceDecl: false,
      })),
      children,
      xmlSpace: "default",
      selfClosing: children.length === 0,
    });
    lvl.children.push(w("rPr", {}, [w("b"), w("color", { val: "C00000" })]));
    const html = renderDocumentHtml(doc);
    const label = /<span class="wk-list-label"[^>]*style="([^"]*)"/.exec(html)?.[1] ?? "";
    expect(label).toContain("font-weight:bold");
    expect(label).toContain("#C00000");
    // The item's text is not affected.
    expect(createStyleResolver(doc).run(para(doc, 0), run(doc, 0)).bold).toBe(false);
  });

  it("formats complex-script text with its own properties (§17.3.2)", () => {
    const doc = createDocx({ paragraphs: ["Latin", "مرحبا"] });
    for (const index of [0, 1]) {
      const r = run(doc, index);
      setRunValProp(r, "szCs", "40");
      setRunValProp(r, "bCs", "1");
    }
    const html = renderDocumentHtml(doc);
    const spans = [...html.matchAll(/<span class="wk-run"[^>]*style="([^"]*)"/g)].map((m) => m[1]);
    // Latin text ignores the complex-script size and bold; Arabic uses them.
    expect(spans[0]).toContain("font-weight:normal");
    expect(spans[0]).not.toContain("font-size:20pt");
    expect(spans[1]).toContain("font-weight:bold");
    expect(spans[1]).toContain("font-size:20pt");
  });

  it("falls back to the default paragraph style for an unknown pStyle", () => {
    const model = headingAndBody();
    caret(model, 1);
    runCommand(model, commands.setParagraphStyleCommand, { styleId: "NoSuchStyle" });
    const resolved = createStyleResolver(model.doc).run(para(model.doc, 1), run(model.doc, 1));
    expect(resolved).toMatchObject({ bold: false, sizeHalfPoints: 22 });
  });

  it("names theme fonts and marks them as the body / headings font", async () => {
    const bytes = await withTheme(createDocx({ paragraphs: ["x"] }), "Aptos Display", "Aptos");
    const doc = openDocx(bytes);
    setRunValProp(run(doc, 0), "rFonts", undefined);
    const rPr = run(doc, 0);
    setRunFormat(rPr, { bold: true });
    // Point the run at the theme's minor (body) font.
    const rFonts = {
      kind: "element" as const,
      name: {
        uri: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        local: "rFonts",
        prefix: "w",
      },
      attrs: [
        {
          name: {
            uri: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            local: "asciiTheme",
            prefix: "w",
          },
          value: "minorHAnsi",
          isNamespaceDecl: false,
        },
      ],
      children: [],
      xmlSpace: "default" as const,
      selfClosing: true,
    };
    rPr.rPr?.children.push(rFonts);
    const resolved = createStyleResolver(doc).run(para(doc, 0), rPr);
    expect(resolved).toMatchObject({ font: "Aptos", fontRole: "body" });
  });
});

/** Add a minimal DrawingML theme part with the given major / minor Latin fonts. */
async function withTheme(doc: Docx, major: string, minor: string): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(toUint8Array(doc));
  const a = "http://schemas.openxmlformats.org/drawingml/2006/main";
  zip.file(
    "word/theme/theme1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="${a}" name="T"><a:themeElements><a:fontScheme name="F"><a:majorFont><a:latin typeface="${major}"/></a:majorFont><a:minorFont><a:latin typeface="${minor}"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>`,
  );
  const types = await zip.file("[Content_Types].xml")!.async("string");
  zip.file(
    "[Content_Types].xml",
    types.replace(
      "</Types>",
      '<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>',
    ),
  );
  const relsPath = "word/_rels/document.xml.rels";
  const rels = await zip.file(relsPath)!.async("string");
  zip.file(
    relsPath,
    rels.replace(
      "</Relationships>",
      '<Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/></Relationships>',
    ),
  );
  return zip.generateAsync({ type: "uint8array" });
}

describe("toggles over style formatting (Word: Bold on a Heading 1 unbolds it)", () => {
  it("reports a style's bold as active and turns it off with w:val=0", () => {
    const model = headingAndBody();
    caret(model, 0, 2);
    expect(commands.toggleBoldCommand.isActive?.(model)).toBe(true);
    runCommand(model, commands.toggleBoldCommand, undefined);
    expect(getRunProp(run(model.doc, 0), "b")).toEqual({ present: true, val: "0" });
    expect(commands.toggleBoldCommand.isActive?.(model)).toBe(false);
    // Back on: the style already bolds it, so the direct override just goes.
    runCommand(model, commands.toggleBoldCommand, undefined);
    expect(getRunProp(run(model.doc, 0), "b").present).toBe(false);
    expect(commands.toggleBoldCommand.isActive?.(model)).toBe(true);
  });

  it("writes plain <w:b/> when nothing below the run is bold", () => {
    const model = headingAndBody();
    caret(model, 1, 1);
    runCommand(model, commands.toggleBoldCommand, undefined);
    expect(getRunProp(run(model.doc, 1), "b")).toEqual({ present: true });
    runCommand(model, commands.toggleBoldCommand, undefined);
    expect(getRunProp(run(model.doc, 1), "b").present).toBe(false);
  });

  it("round-trips the explicit off value", () => {
    const model = headingAndBody();
    caret(model, 0, 2);
    runCommand(model, commands.toggleBoldCommand, undefined);
    const reopened = openDocx(toUint8Array(model.doc));
    expect(getRunProp(run(reopened, 0), "b")).toEqual({ present: true, val: "0" });
  });
});

describe("alignment buttons", () => {
  it("press Align Left for a paragraph with no w:jc, like Word", () => {
    const model = headingAndBody();
    caret(model, 1);
    expect(commands.alignLeftCommand.isActive?.(model)).toBe(true);
    expect(commands.alignCenterCommand.isActive?.(model)).toBe(false);
    runCommand(model, commands.alignCenterCommand, undefined);
    expect(commands.alignLeftCommand.isActive?.(model)).toBe(false);
    expect(commands.alignCenterCommand.isActive?.(model)).toBe(true);
  });
});

describe("clearFormattingCommand (Word's Clear All Formatting)", () => {
  function formatted(): EditorModel {
    const model = headingAndBody();
    const body = para(model.doc, 1);
    setRunFormat(run(model.doc, 1), {
      bold: true,
      italic: true,
      color: "C00000",
      highlight: "yellow",
      fontSizeHalfPoints: 28,
    });
    setParagraphAlignment(body, "center");
    setParagraphIndent(body, { left: 720 });
    return model;
  }

  it("with a caret, resets the paragraph to Normal and leaves the runs", () => {
    const model = formatted();
    caret(model, 0, 2);
    runCommand(model, commands.clearFormattingCommand, undefined);
    expect(getParagraphStyle(para(model.doc, 0))).toBeUndefined();
    caret(model, 1, 2);
    runCommand(model, commands.clearFormattingCommand, undefined);
    expect(pPrNames(para(model.doc, 1))).toEqual([]);
    expect(rPrNames(run(model.doc, 1))).toContain("b");
  });

  it("over part of a paragraph, clears character formatting but keeps highlight and the paragraph", () => {
    const model = formatted();
    model.setSelection({
      anchor: { block: 1, inline: 0, offset: 1 },
      focus: { block: 1, inline: 0, offset: 4 },
    });
    runCommand(model, commands.clearFormattingCommand, undefined);
    const runs = runsOf(para(model.doc, 1));
    expect(runs.map(rPrNames)).toEqual([
      expect.arrayContaining(["b", "highlight"]),
      ["highlight"],
      expect.arrayContaining(["b", "highlight"]),
    ]);
    expect(pPrNames(para(model.doc, 1))).toEqual(expect.arrayContaining(["jc", "ind"]));
  });

  it("across paragraphs, also resets every paragraph whose mark is selected", () => {
    const model = formatted();
    model.setSelection({
      anchor: { block: 0, inline: 0, offset: 0 },
      focus: { block: 1, inline: 0, offset: 9 },
    });
    runCommand(model, commands.clearFormattingCommand, undefined);
    expect(getParagraphStyle(para(model.doc, 0))).toBeUndefined();
    // The range ends before the body paragraph's mark, as a Word drag selection would.
    expect(pPrNames(para(model.doc, 1))).toEqual(expect.arrayContaining(["jc", "ind"]));
    expect(rPrNames(run(model.doc, 1))).toEqual(["highlight"]);
  });

  it("keeps the section properties of a paragraph", () => {
    const model = headingAndBody();
    caret(model, 0);
    runCommand(model, commands.insertSectionBreakCommand, { type: "nextPage" });
    const withSection = paragraphs(model.doc).findIndex((p) => pPrNames(p).includes("sectPr"));
    expect(withSection).toBeGreaterThanOrEqual(0);
    caret(model, withSection);
    runCommand(model, commands.clearFormattingCommand, undefined);
    expect(pPrNames(para(model.doc, withSection))).toEqual(["sectPr"]);
  });
});

describe("indentStepCommand", () => {
  it("steps the left indent by 0.5 in from the effective value", () => {
    const model = headingAndBody();
    setParagraphIndent(para(model.doc, 1), { left: 300, hanging: 200 });
    caret(model, 1);
    runCommand(model, commands.indentStepCommand, { direction: "increase" });
    expect(createStyleResolver(model.doc).paragraph(para(model.doc, 1))).toMatchObject({
      left: 720,
      hanging: 200,
    });
    runCommand(model, commands.indentStepCommand, { direction: "decrease" });
    runCommand(model, commands.indentStepCommand, { direction: "decrease" });
    expect(createStyleResolver(model.doc).paragraph(para(model.doc, 1)).left).toBe(0);
  });
});

describe("setLineSpacingCommand", () => {
  it("sets an auto line multiple and keeps the paragraph's own before / after", () => {
    const model = headingAndBody();
    setParagraphSpacing(para(model.doc, 1), { before: 120, after: 0 });
    caret(model, 1);
    runCommand(model, commands.setLineSpacingCommand, { multiple: 1.5 });
    expect(createStyleResolver(model.doc).paragraph(para(model.doc, 1))).toMatchObject({
      before: 120,
      after: 0,
      line: 360,
      lineRule: "auto",
    });
    expect(commands.lineSpacingOf(model)).toBe(1.5);
  });

  it("rejects a non-positive multiple before changing anything", () => {
    const model = headingAndBody();
    caret(model, 1);
    expect(() => runCommand(model, commands.setLineSpacingCommand, { multiple: 0 })).toThrow(
      RangeError,
    );
    expect(pPrNames(para(model.doc, 1))).toEqual([]);
  });
});

describe("pageGeometry", () => {
  it("uses Word's Letter / 1 in defaults without a sectPr", () => {
    expect(pageGeometry(createDocx({ paragraphs: ["x"] }))).toEqual({
      width: 12240,
      height: 15840,
      top: 1440,
      right: 1440,
      bottom: 1440,
      left: 1440,
    });
  });

  it("reads the body section's size and margins", () => {
    const doc = createDocx({ paragraphs: ["x"] });
    setPageSize(doc, { widthTwips: 11906, heightTwips: 16838 });
    setPageMargins(doc, {
      top: 1000,
      right: 900,
      bottom: 1100,
      left: 800,
      header: 0,
      footer: 0,
      gutter: 0,
    });
    expect(pageGeometry(doc)).toMatchObject({
      width: 11906,
      height: 16838,
      top: 1000,
      right: 900,
      bottom: 1100,
      left: 800,
    });
  });
});

describe("tables on the canvas", () => {
  it("draws no borders when the document defines none, like Word", () => {
    const doc = createDocx({ paragraphs: [] });
    addTable(doc, [["a", "b"]]);
    const table = doc.document.body.blocks[0];
    if (table?.kind !== "table") throw new Error("expected a table");
    expect(resolveTable(doc, table)).toMatchObject({
      borders: {},
      cellMargins: { left: 108, right: 108 },
    });
    const html = renderDocumentHtml(doc);
    expect(html).toContain("border-top:var(--wk-gridline,none)");
    expect(html).not.toMatch(/border-\w+:[0-9.]+pt/);
  });

  it("draws the borders a table style defines", () => {
    const doc = createDocx({ paragraphs: [] });
    const table = addTable(doc, [
      ["a", "b"],
      ["c", "d"],
    ]);
    setTableBorders(table, { style: "single", sizeEighthsOfPoint: 4, color: "000000" });
    const html = renderDocumentHtml(doc);
    expect(html).toContain("border-top:0.5pt solid #000000");
  });
});

function tableOf(doc: Docx): WmlTable {
  const table = doc.document.body.blocks.find((b) => b.kind === "table");
  if (table?.kind !== "table") throw new Error("expected a table");
  return table;
}

describe("table position and Word's compatibility mode", () => {
  it("puts the cell text on the margin in Compatibility Mode (no compatSetting)", () => {
    const doc = createDocx({ paragraphs: [] });
    addTable(doc, [["a"]]);
    expect(compatibilityMode(doc)).toBe(12);
    expect(resolveTable(doc, tableOf(doc)).leftEdge).toBe(-108);
    expect(renderDocumentHtml(doc)).toContain("margin-left:-5.4pt");
  });

  it("puts the border on the margin in Word 2013+ mode", async () => {
    const base = createDocx({ paragraphs: [] });
    addTable(base, [["a"]]);
    const doc = openDocx(await withSettings(base, 15));
    expect(compatibilityMode(doc)).toBe(15);
    expect(resolveTable(doc, tableOf(doc)).leftEdge).toBe(0);
  });
});

/** Add a settings part declaring the given Word compatibility mode. */
async function withSettings(doc: Docx, mode: number): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(toUint8Array(doc));
  const w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  zip.file(
    "word/settings.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="${w}"><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="${mode}"/></w:compat></w:settings>`,
  );
  const types = await zip.file("[Content_Types].xml")!.async("string");
  zip.file(
    "[Content_Types].xml",
    types.replace(
      "</Types>",
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>',
    ),
  );
  const relsPath = "word/_rels/document.xml.rels";
  const rels = await zip.file(relsPath)!.async("string");
  zip.file(
    relsPath,
    rels.replace(
      "</Relationships>",
      '<Relationship Id="rIdSettings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>',
    ),
  );
  return zip.generateAsync({ type: "uint8array" });
}

describe("run rendering", () => {
  it("draws superscript and falls back from Office fonts", () => {
    const model = headingAndBody();
    setRunValProp(run(model.doc, 1), "vertAlign", "superscript");
    const html = renderDocumentHtml(model.doc);
    expect(html).toContain("vertical-align:super");
    // Calibri's metric-compatible stand-in comes before the East Asian font
    // (docDefaults: Times New Roman), and the generic family last.
    expect(html).toContain("font-family:'Calibri',Carlito,'Times New Roman',sans-serif");
  });
});
