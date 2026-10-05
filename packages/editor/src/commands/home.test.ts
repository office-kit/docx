/**
 * The Home tab's commands: Font, Paragraph, Styles, and Find and Replace.
 * Each test drives the command the way the ribbon does and checks the XML
 * the saved document holds (and that the package stays valid).
 */

import {
  addStyle,
  createDocx,
  type Docx,
  getParagraphNumbering,
  getParagraphStyle,
  getRunProp,
  openDocx,
  paragraphs,
  paragraphText,
  setRunFormat,
  setRunValProp,
  toUint8Array,
  validate,
  type WmlParagraph,
  type WmlRun,
} from "@office-kit/docx";
import { describe, expect, it } from "vitest";
import { findMatches } from "../find.js";
import { BULLET_PRESETS, MULTILEVEL_PRESETS, NUMBERING_PRESETS } from "../list-presets.js";
import { editorFor, renderDocumentHtml } from "../index.js";
import type { EditorModel } from "../model.js";
import { createStyleResolver } from "../resolve.js";
import { changeCase } from "../text-case.js";
import { applyShade, applyTint, DEFAULT_THEME_COLORS, themeColorGrid } from "../theme-color.js";
import {
  applyListPresetCommand,
  continueNumberingCommand,
  removeListCommand,
  restartNumberingCommand,
  selectionListKind,
  setListLevelCommand,
} from "./list.js";
import {
  alignDistributedCommand,
  bordersPresetActive,
  bordersPresetCommand,
  indentStepCommand,
  insertHorizontalLineCommand,
  paragraphFormatCommand,
  paragraphShadingColorCommand,
  setTabsCommand,
  sortParagraphsCommand,
  toggleParagraphSpaceCommand,
} from "./paragraph.js";
import {
  applyStyleCommand,
  deleteStyleCommand,
  modifyStyleCommand,
  newStyleCommand,
} from "./style.js";
import {
  changeCaseCommand,
  clearCharacterFormattingCommand,
  copyFormat,
  eastAsianLayoutCommand,
  encloseCharactersCommand,
  findReplaceAllCommand,
  fitTextCommand,
  fontFormatCommand,
  growFontCommand,
  nextFontSize,
  pasteFormatCommand,
  phoneticGuideCommand,
  replaceNextCommand,
  setColorCommand,
  setFontCommand,
  setUnderlineColorCommand,
  setUnderlineStyleCommand,
  toggleCharacterBorderCommand,
  toggleCharacterShadingCommand,
} from "./text.js";
import { runCommand } from "./types.js";

function model(...texts: string[]): EditorModel {
  return editorFor(createDocx({ paragraphs: texts }));
}

function select(m: EditorModel, block: number, start: number, end: number, endBlock = block): void {
  m.setSelection({
    anchor: { block, inline: 0, offset: start },
    focus: { block: endBlock, inline: 0, offset: end },
  });
}

function caret(m: EditorModel, block: number): void {
  m.setSelection({
    anchor: { block, inline: 0, offset: 0 },
    focus: { block, inline: 0, offset: 0 },
  });
}

function xml(doc: Docx, part = "/word/document.xml"): string {
  return new TextDecoder().decode(openDocx(toUint8Array(doc)).opc.parts.get(part)?.data);
}

function expectValid(doc: Docx): void {
  expect(validate(openDocx(toUint8Array(doc))).filter((i) => i.level === "error")).toEqual([]);
}

function runs(p: WmlParagraph): WmlRun[] {
  return p.children.filter((c): c is WmlRun => c.kind === "run");
}

const para = (m: EditorModel, i = 0): WmlParagraph => paragraphs(m.doc)[i]!;

describe("Font group", () => {
  it("sets the theme body font and replaces it with a named font", () => {
    const m = model("Hello");
    select(m, 0, 0, 5);
    runCommand(m, setFontCommand, { theme: "minor" });
    expect(xml(m.doc)).toContain('w:asciiTheme="minorHAnsi"');
    runCommand(m, setFontCommand, { font: "Georgia" });
    expect(xml(m.doc)).toContain('<w:rFonts w:ascii="Georgia" w:hAnsi="Georgia"/>');
  });

  it("writes a theme colour with its tint", () => {
    const m = model("Hello");
    select(m, 0, 0, 5);
    const swatch = themeColorGrid(DEFAULT_THEME_COLORS)[4]![1]!;
    runCommand(m, setColorCommand, {
      color: swatch.rgb,
      themeColor: swatch.themeColor,
      themeTint: swatch.themeTint!,
    });
    expect(xml(m.doc)).toContain(
      `<w:color w:val="${swatch.rgb}" w:themeColor="accent1" w:themeTint="33"/>`,
    );
    expectValid(m.doc);
  });

  it("keeps the underline colour when the style changes, and underlines when only a colour is picked", () => {
    const m = model("Hello");
    select(m, 0, 0, 5);
    runCommand(m, setUnderlineColorCommand, { color: { rgb: "FF0000" } });
    expect(xml(m.doc)).toContain('<w:u w:val="single" w:color="FF0000"/>');
    runCommand(m, setUnderlineStyleCommand, { style: "wavyDouble" });
    expect(xml(m.doc)).toContain('<w:u w:val="wavyDouble" w:color="FF0000"/>');
  });

  it("steps sizes through Word's list", () => {
    expect(nextFontSize(11, 1)).toBe(12);
    expect(nextFontSize(72, 1)).toBe(80);
    expect(nextFontSize(8, -1)).toBe(7);
    const m = model("Hello");
    select(m, 0, 0, 5);
    runCommand(m, growFontCommand, { direction: 1 });
    expect(xml(m.doc)).toContain('<w:sz w:val="24"/>');
  });

  it("toggles character shading and border", () => {
    const m = model("Hello");
    select(m, 0, 0, 5);
    runCommand(m, toggleCharacterShadingCommand, undefined);
    runCommand(m, toggleCharacterBorderCommand, undefined);
    expect(toggleCharacterShadingCommand.isActive?.(m)).toBe(true);
    const out = xml(m.doc);
    expect(out).toContain('<w:shd w:val="pct15" w:color="auto" w:fill="FFFFFF"/>');
    expect(out).toContain('<w:bdr w:val="single" w:sz="4" w:space="0" w:color="auto"/>');
    runCommand(m, toggleCharacterShadingCommand, undefined);
    expect(xml(m.doc)).not.toContain("<w:shd");
  });

  it("applies the Font dialog: effects, exclusive pairs, and advanced spacing", () => {
    const m = model("Hello");
    select(m, 0, 0, 5);
    runCommand(m, fontFormatCommand, {
      effects: { smallCaps: true, dstrike: true },
      scale: 150,
      spacing: 20,
      position: 6,
      kern: 24,
      vertAlign: "superscript",
    });
    let out = xml(m.doc);
    for (const el of [
      "<w:smallCaps/>",
      "<w:dstrike/>",
      '<w:w w:val="150"/>',
      '<w:spacing w:val="20"/>',
      '<w:position w:val="6"/>',
      '<w:kern w:val="24"/>',
    ])
      expect(out).toContain(el);
    runCommand(m, fontFormatCommand, { effects: { caps: true, strike: true } });
    out = xml(m.doc);
    expect(out).toContain("<w:caps/>");
    expect(out).not.toContain("smallCaps");
    expect(out).not.toContain("dstrike");
    expectValid(m.doc);
  });

  it("resets character formatting but keeps the character style", () => {
    const m = model("Hello");
    const run = runs(para(m))[0]!;
    setRunFormat(run, { bold: true });
    setRunValProp(run, "rStyle", "Strong");
    select(m, 0, 0, 5);
    runCommand(m, clearCharacterFormattingCommand, undefined);
    const after = runs(para(m))[0]!;
    expect(getRunProp(after, "b").present).toBe(false);
    expect(getRunProp(after, "rStyle").val).toBe("Strong");
  });
});

describe("Change Case", () => {
  it("transforms case and Japanese widths", () => {
    expect(changeCase("hELLO wORLD. aGAIN", "sentence")).toBe("Hello world. Again");
    expect(changeCase("hello wide world", "title")).toBe("Hello Wide World");
    expect(changeCase("AbC", "toggle")).toBe("aBc");
    expect(changeCase("ABC 123", "fullWidth")).toBe("ＡＢＣ　１２３");
    expect(changeCase("ｶﾞｷﾞﾊﾟ", "fullWidth")).toBe("ガギパ");
    expect(changeCase("ガギパ", "halfWidth")).toBe("ｶﾞｷﾞﾊﾟ");
    expect(changeCase("ひらがな", "katakana")).toBe("ヒラガナ");
    expect(changeCase("カタカナ", "hiragana")).toBe("かたかな");
  });

  it("changes the selected text only", () => {
    const m = model("one two three");
    select(m, 0, 4, 7);
    runCommand(m, changeCaseCommand, { mode: "upper" });
    expect(paragraphText(para(m))).toBe("one TWO three");
    select(m, 0, 0, 3);
    runCommand(m, changeCaseCommand, { mode: "fullWidth" });
    expect(paragraphText(para(m))).toBe("ｏｎｅ TWO three");
  });
});

describe("Format Painter", () => {
  it("copies character and paragraph formatting", () => {
    const m = model("Source", "Target text");
    select(m, 0, 0, 6);
    runCommand(m, setColorCommand, { color: "00B050" });
    runCommand(m, alignDistributedCommand, undefined);
    caret(m, 0);
    const copied = copyFormat(m)!;
    expect(copied.paragraph).toBeDefined();
    select(m, 1, 0, 6);
    runCommand(m, pasteFormatCommand, copied);
    const out = xml(m.doc);
    expect(out.match(/<w:jc w:val="distribute"\/>/g)).toHaveLength(2);
    expect(out.match(/<w:color w:val="00B050"\/>/g)).toHaveLength(2);
  });
});

describe("Phonetic Guide, Enclose Characters, Asian Layout", () => {
  const options = {
    alignment: "center" as const,
    rubySizeHalfPoints: 10,
    raiseHalfPoints: 20,
    baseSizeHalfPoints: 21,
    language: "ja-JP",
  };

  it("adds, edits and removes a phonetic guide", () => {
    const m = model("日本語です");
    select(m, 0, 0, 3);
    runCommand(m, phoneticGuideCommand, { ruby: "にほんご", options });
    expect(xml(m.doc)).toContain("<w:rt>");
    expect(renderDocumentHtml(m.doc)).toContain("<ruby");
    expect(paragraphText(para(m)).endsWith("です")).toBe(true);
    runCommand(m, phoneticGuideCommand, { ruby: "", options });
    expect(xml(m.doc)).not.toContain("<w:ruby>");
    expectValid(m.doc);
  });

  it("encloses a character in an EQ field", () => {
    const m = model("秘密");
    select(m, 0, 0, 1);
    runCommand(m, encloseCharactersCommand, { enclosure: "circle", style: "enlargeSymbol" });
    const out = xml(m.doc);
    expect(out).toContain("eq \\o\\ac(");
    expect(out).toContain("<w:instrText>○</w:instrText>");
  });

  it("writes two lines in one and fit text with ids", () => {
    const m = model("ABCD");
    select(m, 0, 0, 4);
    runCommand(m, eastAsianLayoutCommand, { layout: { kind: "twoLinesInOne", brackets: "round" } });
    runCommand(m, fitTextCommand, { twips: 1440 });
    const out = xml(m.doc);
    expect(out).toContain('<w:fitText w:val="1440" w:id="2"/>');
    expect(out).toContain('<w:eastAsianLayout w:id="1" w:combine="1" w:combineBrackets="round"/>');
    expectValid(m.doc);
  });
});

describe("Paragraph group", () => {
  it("applies the Paragraph dialog", () => {
    const m = model("p");
    caret(m, 0);
    runCommand(m, paragraphFormatCommand, {
      alignment: "right",
      outlineLevel: 1,
      indent: { left: 720, hanging: 360 },
      spacing: { before: 120, line: 360, lineRule: "auto" },
      toggles: { keepNext: true, kinsoku: false, widowControl: true },
    });
    const out = xml(m.doc);
    expect(out).toMatch(
      /<w:keepNext\/><w:widowControl\/>.*<w:kinsoku w:val="0"\/>.*<w:spacing w:before="120" w:line="360" w:lineRule="auto"\/><w:ind w:left="720" w:hanging="360"\/><w:jc w:val="right"\/><w:outlineLvl w:val="1"\/>/,
    );
    expectValid(m.doc);
  });

  it("adds and removes space before", () => {
    const m = model("p");
    caret(m, 0);
    runCommand(m, toggleParagraphSpaceCommand, { side: "before" });
    expect(xml(m.doc)).toContain('w:before="240"');
    runCommand(m, toggleParagraphSpaceCommand, { side: "before" });
    expect(xml(m.doc)).toContain('w:before="0"');
  });

  it("shades with a theme colour and borders by preset", () => {
    const m = model("a", "b");
    select(m, 0, 0, 1, 1);
    runCommand(m, paragraphShadingColorCommand, {
      fill: { rgb: "DCEAF7", themeColor: "accent1", themeTint: 0x33 },
    });
    runCommand(m, bordersPresetCommand, { preset: "outside" });
    expect(bordersPresetActive(m, "outside")).toBe(true);
    runCommand(m, bordersPresetCommand, { preset: "insideHorizontal" });
    let out = xml(m.doc);
    expect(out).toContain('w:themeFill="accent1" w:themeFillTint="33"');
    expect(out).toMatch(
      /<w:pBdr><w:top [^>]*\/><w:left [^>]*\/><w:bottom [^>]*\/><w:right [^>]*\/><w:between /,
    );
    runCommand(m, bordersPresetCommand, { preset: "bottom" });
    expect(xml(m.doc)).not.toMatch(/<w:bottom /);
    runCommand(m, bordersPresetCommand, { preset: "none" });
    out = xml(m.doc);
    expect(out).not.toContain("pBdr");
    expectValid(m.doc);
  });

  it("sets tab stops", () => {
    const m = model("a\tb");
    caret(m, 0);
    runCommand(m, setTabsCommand, {
      tabs: [{ position: 2880, alignment: "decimal", leader: "dot" }],
    });
    expect(xml(m.doc)).toContain('<w:tab w:val="decimal" w:leader="dot" w:pos="2880"/>');
  });

  it("sorts paragraphs by text and by number", () => {
    const m = model("pear", "apple", "fig");
    select(m, 0, 0, 3, 2);
    runCommand(m, sortParagraphsCommand, { by: "text", order: "ascending" });
    expect(paragraphs(m.doc).map(paragraphText)).toEqual(["apple", "fig", "pear"]);
    const n = model("item 10", "item 9", "item 100");
    select(n, 0, 0, 8, 2);
    runCommand(n, sortParagraphsCommand, { by: "number", order: "descending" });
    expect(paragraphs(n.doc).map(paragraphText)).toEqual(["item 100", "item 10", "item 9"]);
  });

  it("inserts a horizontal line as a VML rule", () => {
    const m = model("a");
    caret(m, 0);
    runCommand(m, insertHorizontalLineCommand, undefined);
    expect(xml(m.doc)).toContain('o:hr="t"');
    expect(renderDocumentHtml(m.doc)).toBeTruthy();
    expectValid(m.doc);
  });
});

describe("Lists", () => {
  it("applies library entries, continues the same list, and renders labels", () => {
    const m = model("one", "two", "plain", "three");
    select(m, 0, 0, 3, 1);
    runCommand(m, applyListPresetCommand, { levels: NUMBERING_PRESETS[1]!.levels });
    select(m, 3, 0, 5);
    runCommand(m, applyListPresetCommand, { levels: NUMBERING_PRESETS[1]!.levels });
    const ids = [0, 1, 3].map((i) => getParagraphNumbering(para(m, i))?.numId);
    expect(new Set(ids).size).toBe(1);
    expect(selectionListKind(m)).toBe("numbered");
    const html = renderDocumentHtml(m.doc);
    expect(html).toContain(">1)</span>");
    expect(html).toContain(">3)</span>");
    expect(createStyleResolver(m.doc).paragraph(para(m)).left).toBe(720);
    expectValid(m.doc);
  });

  it("changes a whole list's format, levels, restarts and continues", () => {
    const m = model("a", "b", "c");
    select(m, 0, 0, 1, 2);
    runCommand(m, applyListPresetCommand, { levels: NUMBERING_PRESETS[0]!.levels });
    caret(m, 1);
    runCommand(m, setListLevelCommand, { ilvl: 1 });
    expect(getParagraphNumbering(para(m, 1))?.ilvl).toBe(1);
    caret(m, 1);
    runCommand(m, indentStepCommand, { direction: "decrease" });
    expect(getParagraphNumbering(para(m, 1))?.ilvl).toBe(0);
    caret(m, 2);
    runCommand(m, restartNumberingCommand, { start: 5 });
    expect(xml(m.doc, "/word/numbering.xml")).toContain('<w:startOverride w:val="5"/>');
    expect(renderDocumentHtml(m.doc)).toContain(">5.</span>");
    caret(m, 2);
    runCommand(m, continueNumberingCommand, undefined);
    expect(getParagraphNumbering(para(m, 2))?.numId).toBe(getParagraphNumbering(para(m, 0))?.numId);
    select(m, 0, 0, 1, 2);
    runCommand(m, applyListPresetCommand, { levels: BULLET_PRESETS[0]!.levels });
    expect(selectionListKind(m)).toBe("bullet");
    expect(renderDocumentHtml(m.doc)).toContain(">•</span>");
    runCommand(m, removeListCommand, undefined);
    expect(getParagraphNumbering(para(m, 0))).toBeUndefined();
  });

  it("links a heading multilevel list through the heading styles", () => {
    const m = model("Intro");
    caret(m, 0);
    runCommand(m, applyStyleCommand, { styleId: "Heading1" });
    runCommand(m, applyListPresetCommand, { levels: MULTILEVEL_PRESETS[3]!.levels });
    expect(getParagraphNumbering(para(m))).toBeUndefined();
    expect(xml(m.doc, "/word/styles.xml")).toMatch(
      /w:styleId="Heading1">.*<w:numPr><w:ilvl w:val="0"\/><w:numId w:val="\d+"\/><\/w:numPr>/,
    );
    expect(renderDocumentHtml(m.doc)).toContain(">1</span>");
  });
});

describe("Styles", () => {
  it("creates a built-in style the first time it is applied", () => {
    const m = model("Title text");
    caret(m, 0);
    runCommand(m, applyStyleCommand, { styleId: "Title" });
    expect(getParagraphStyle(para(m))).toBe("Title");
    expect(createStyleResolver(m.doc).paragraph(para(m)).toggles.contextualSpacing).toBe(true);
    runCommand(m, applyStyleCommand, { styleId: "Normal" });
    expect(getParagraphStyle(para(m))).toBeUndefined();
  });

  it("applies a character style, and a linked style's character half to part of a paragraph", () => {
    const m = model("Hello world");
    select(m, 0, 0, 5);
    runCommand(m, applyStyleCommand, { styleId: "Strong" });
    expect(getRunProp(runs(para(m))[0]!, "rStyle").val).toBe("Strong");
    select(m, 0, 0, 5);
    runCommand(m, applyStyleCommand, { styleId: "Quote" });
    expect(getParagraphStyle(para(m))).toBeUndefined();
    expect(getRunProp(runs(para(m))[0]!, "rStyle").val).toBe("QuoteChar");
  });

  it("creates, modifies and deletes a style", () => {
    const m = model("Body");
    caret(m, 0);
    const id = runCommand(m, newStyleCommand, {
      type: "paragraph",
      name: "My Note",
      basedOn: "Normal",
      next: "Normal",
      quickStyle: true,
      font: { effects: { italic: true }, color: { rgb: "7030A0" } },
      paragraph: { alignment: "center" },
    })!;
    expect(id).toBe("MyNote");
    expect(getParagraphStyle(para(m))).toBe("MyNote");
    expect(() => runCommand(m, newStyleCommand, { type: "paragraph", name: "My Note" })).toThrow();
    runCommand(m, modifyStyleCommand, {
      styleId: id,
      name: "My Note",
      basedOn: "Normal",
      font: { sizeHalfPoints: 28 },
    });
    const styles = xml(m.doc, "/word/styles.xml");
    expect(styles).toMatch(
      /w:styleId="MyNote" w:customStyle="1"><w:name w:val="My Note"\/><w:basedOn w:val="Normal"\/><w:next w:val="Normal"\/><w:qFormat\/><w:pPr><w:jc w:val="center"\/><\/w:pPr><w:rPr><w:i\/><w:color w:val="7030A0"\/><w:sz w:val="28"\/>/,
    );
    runCommand(m, deleteStyleCommand, { styleId: id });
    expect(getParagraphStyle(para(m))).toBeUndefined();
    expect(() => runCommand(m, deleteStyleCommand, { styleId: "Normal" })).toThrow();
    expectValid(m.doc);
  });

  it("previews a style's formatting", () => {
    const m = model("x");
    addStyle(m.doc, {
      type: "paragraph",
      styleId: "Big",
      name: "Big",
      bold: true,
      fontSizeHalfPoints: 48,
    });
    expect(createStyleResolver(m.doc).style("Big").run).toMatchObject({
      bold: true,
      sizeHalfPoints: 48,
    });
  });
});

describe("Find and Replace", () => {
  it("finds with match case, whole words, wildcards and specials", () => {
    const m = model("The cat sat on the mat.", "Cat\tdog");
    expect(findMatches(m.doc, { query: "the" })).toHaveLength(2);
    expect(findMatches(m.doc, { query: "the", matchCase: true })).toHaveLength(1);
    expect(findMatches(m.doc, { query: "at", wholeWords: true })).toHaveLength(0);
    expect(findMatches(m.doc, { query: "<[cms]at>", wildcards: true }).map((x) => x.text)).toEqual([
      "cat",
      "sat",
      "mat",
    ]);
    expect(findMatches(m.doc, { query: "Cat^tdog" })).toHaveLength(1);
    expect(findMatches(m.doc, { query: "?at{1,1}", wildcards: true })).toHaveLength(4);
  });

  it("filters by format", () => {
    const m = model("plain bold plain");
    select(m, 0, 6, 10);
    runCommand(m, fontFormatCommand, { effects: { bold: true } });
    expect(findMatches(m.doc, { query: "bold", format: { bold: true } })).toHaveLength(1);
    expect(findMatches(m.doc, { query: "plain", format: { bold: true } })).toHaveLength(0);
    expect(findMatches(m.doc, { query: "", format: { bold: true } }).map((x) => x.text)).toEqual([
      "bold",
    ]);
  });

  it("replaces all with wildcard groups and formatting, keeping the found text's formatting", () => {
    const m = model("Smith, John and Doe, Jane");
    const n = runCommand(m, findReplaceAllCommand, {
      find: { query: "(<[A-Z][a-z]@), (<[A-Z][a-z]@)", wildcards: true },
      replacement: "\\2 \\1",
      format: { italic: true },
    });
    expect(n).toBe(2);
    expect(paragraphText(para(m))).toBe("John Smith and Jane Doe");
    expect(xml(m.doc).match(/<w:i\/>/g)).toHaveLength(2);
  });

  it("replaces one match at a time and moves to the next", () => {
    const m = model("a-a-a");
    caret(m, 0);
    const params = { find: { query: "a" }, replacement: "bb" };
    expect(runCommand(m, replaceNextCommand, params)).toBe(false);
    expect(runCommand(m, replaceNextCommand, params)).toBe(true);
    expect(paragraphText(para(m))).toBe("bb-a-a");
    expect(runCommand(m, replaceNextCommand, params)).toBe(true);
    expect(paragraphText(para(m))).toBe("bb-bb-a");
  });
});

describe("Theme colours", () => {
  it("builds Word's 10 × 6 grid with Word's tints and shades", () => {
    const grid = themeColorGrid(DEFAULT_THEME_COLORS);
    expect(grid).toHaveLength(10);
    expect(grid.every((col) => col.length === 6)).toBe(true);
    // White: 5 % darker first; black: 50 % lighter first.
    expect(grid[0]![1]).toMatchObject({
      themeColor: "background1",
      themeShade: 0xf2,
      rgb: "F2F2F2",
    });
    expect(grid[1]![1]).toMatchObject({ themeColor: "text1", themeTint: 0x80, rgb: "7F7F7F" });
    expect(applyTint("4472C4", 0x33)).toBe("DAE3F3");
    // Word writes 2F5496 for Heading 1 (accent1 4472C4, themeShade BF).
    expect(applyShade("4472C4", 0xbf)).toBe("2F5496");
  });
});

describe("Rendering", () => {
  it("draws the text effects the Font dialog sets", () => {
    const m = model("Hello");
    select(m, 0, 0, 5);
    runCommand(m, fontFormatCommand, {
      effects: { smallCaps: true, outline: true, vanish: true },
      spacing: 40,
    });
    runCommand(m, toggleCharacterShadingCommand, undefined);
    const html = renderDocumentHtml(m.doc);
    expect(html).toContain("font-variant-caps:small-caps");
    expect(html).toContain("-webkit-text-stroke");
    expect(html).toContain("--wk-hidden-display");
    expect(html).toContain("letter-spacing:2pt");
    expect(html).toContain("background-color:#d9d9d9");
  });
});
