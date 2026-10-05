import {
  createDocx,
  getDefaultParagraphSpacing,
  getDocumentSetting,
  getParagraphProp,
  getSectionProperties,
  getTheme,
  getWatermark,
  MARGINS_NORMAL,
  openDocx,
  PAGE_SIZE_A4,
  paragraphs,
  paragraphText,
  sectionCount,
  STYLE_SETS,
  THEME_COLOR_SCHEMES,
  THEMES,
  toUint8Array,
  validatePackage,
  type XmlNode,
} from "@office-kit/docx";
import { describe, expect, it } from "vitest";
import { editorFor } from "../index.js";
import type { EditorModel } from "../model.js";
import { resolvePageBackground, resolveSectionLayout } from "../section-layout.js";
import { caretAt } from "../selection.js";
import { resolveTheme, themePalette } from "../theme.js";
import * as c from "./index.js";
import { runCommand } from "./types.js";

function editor(texts: string[], caret = { block: 0, inline: 0, offset: 0 }): EditorModel {
  const model = editorFor(createDocx({ paragraphs: texts }));
  model.setSelection(caretAt(caret));
  return model;
}

function saved(model: EditorModel): ReturnType<typeof openDocx> {
  const doc = openDocx(toUint8Array(model.doc));
  expect(validatePackage(doc.opc).filter((i) => i.level === "error")).toEqual([]);
  return doc;
}

function attrs(n: XmlNode | undefined): Record<string, string> {
  return Object.fromEntries(
    (n && n.kind === "element" ? n.attrs : []).map((a) => [a.name.local, a.value]),
  );
}

describe("section breaks and Apply to", () => {
  it("splits the caret paragraph and ends the section with the first half", () => {
    const model = editor(["Hello world"], { block: 0, inline: 0, offset: 5 });
    runCommand(model, c.insertSectionBreakCommand, { type: "continuous" });
    const doc = saved(model);
    expect(paragraphs(doc).map(paragraphText)).toEqual(["Hello", " world"]);
    expect(sectionCount(doc)).toBe(2);
    expect(getSectionProperties(doc, 1).start).toBe("continuous");
    expect(model.selection?.focus.block).toBe(1);
  });

  it("starts a continuous section at the caret for columns applied to this point forward", () => {
    const model = editor(["one", "two"], { block: 1, inline: 0, offset: 0 });
    runCommand(model, c.columnsCommand, {
      columns: { count: 2, spaceTwips: 720, separator: false },
      target: "forward",
    });
    const doc = saved(model);
    expect(sectionCount(doc)).toBe(2);
    expect(getSectionProperties(doc, 0).columns.count).toBe(1);
    expect(getSectionProperties(doc, 1).columns.count).toBe(2);
    expect(getSectionProperties(doc, 1).start).toBe("continuous");
  });

  it("applies page setup to the whole document and to the caret section", () => {
    const model = editor(["a", "b"], { block: 1, inline: 0, offset: 0 });
    runCommand(model, c.insertSectionBreakCommand, { type: "nextPage" });
    runCommand(model, c.pageSetupCommand, {
      target: "document",
      pageSize: PAGE_SIZE_A4,
      margins: { ...MARGINS_NORMAL, gutter: 720 },
      section: { verticalAlignment: "center", titlePage: true },
      settings: {
        mirrorMargins: true,
        gutterAtTop: true,
        evenAndOddHeaders: true,
        bookFoldPrintingSheets: 4,
      },
    });
    runCommand(model, c.setOrientationCommand, { orientation: "landscape" });
    const doc = saved(model);
    for (const i of [0, 1]) {
      expect(getSectionProperties(doc, i).pageSize.widthTwips > 0).toBe(true);
      expect(getSectionProperties(doc, i).verticalAlignment).toBe("center");
      expect(getSectionProperties(doc, i).margins.gutter).toBe(720);
    }
    expect(getSectionProperties(doc, 0).pageSize.orientation).toBe("portrait");
    expect(getSectionProperties(doc, 1).pageSize.orientation).toBe("landscape");
    expect(getDocumentSetting(doc, "mirrorMargins").present).toBe(true);
    expect(getDocumentSetting(doc, "bookFoldPrintingSheets").val).toBe("4");
  });
});

describe("layout commands", () => {
  it("inserts page, column and text-wrapping breaks at the caret", () => {
    const model = editor(["abc"], { block: 0, inline: 0, offset: 1 });
    runCommand(model, c.insertBreakCommand, { kind: "textWrapping" });
    runCommand(model, c.insertBreakCommand, { kind: "column" });
    const doc = saved(model);
    const pieces =
      paragraphs(doc)[0]?.children.flatMap((r) => (r.kind === "run" ? r.pieces : [])) ?? [];
    expect(pieces.filter((p) => p.kind === "break")).toEqual([
      { kind: "break", breakType: "textWrapping", clear: "all" },
      { kind: "break", breakType: "column" },
    ]);
  });

  it("sets and removes line numbering, and suppresses it per paragraph", () => {
    const model = editor(["a"]);
    runCommand(model, c.lineNumbersCommand, {
      lineNumbering: { start: 1, countBy: 1, restart: "newPage" },
    });
    expect(getSectionProperties(model.doc).lineNumbering?.restart).toBe("newPage");
    runCommand(model, c.suppressLineNumbersCommand, undefined);
    expect(c.suppressLineNumbersCommand.isActive?.(model)).toBe(true);
    runCommand(model, c.lineNumbersCommand, { lineNumbering: null });
    const doc = saved(model);
    expect(getSectionProperties(doc).lineNumbering).toBeUndefined();
    expect(getParagraphProp(paragraphs(doc)[0]!, "suppressLineNumbers").present).toBe(true);
  });

  it("writes hyphenation settings", () => {
    const model = editor(["a"]);
    runCommand(model, c.hyphenationCommand, {
      automatic: true,
      hyphenateCaps: false,
      zoneTwips: 357,
      consecutiveLimit: 2,
    });
    expect(c.hyphenationCommand.isActive?.(model)).toBe(true);
    const doc = saved(model);
    expect(getDocumentSetting(doc, "doNotHyphenateCaps").present).toBe(true);
    expect(getDocumentSetting(doc, "hyphenationZone").val).toBe("357");
    expect(getDocumentSetting(doc, "consecutiveHyphenLimit").val).toBe("2");
    expect(() => runCommand(model, c.hyphenationCommand, { consecutiveLimit: -1 })).toThrow(
      RangeError,
    );
  });

  it("changes one indent / spacing box without dropping the others", () => {
    const model = editor(["a"]);
    runCommand(model, c.layoutIndentCommand, { right: 360 });
    runCommand(model, c.layoutIndentCommand, { left: 720 });
    runCommand(model, c.layoutSpacingCommand, { after: 120 });
    runCommand(model, c.layoutSpacingCommand, { before: 240 });
    const p = paragraphs(saved(model))[0];
    const ind = p?.pPr?.children.find((n) => n.kind === "element" && n.name.local === "ind");
    const spacing = p?.pPr?.children.find(
      (n) => n.kind === "element" && n.name.local === "spacing",
    );
    expect(attrs(ind)).toEqual({ left: "720", right: "360" });
    expect(attrs(spacing)).toEqual({ before: "240", after: "120" });
  });

  it("sets the section text direction", () => {
    const model = editor(["a"]);
    runCommand(model, c.textDirectionCommand, { direction: "tbRl" });
    expect(getSectionProperties(saved(model)).textDirection).toBe("tbRl");
  });
});

describe("design commands", () => {
  it("applies a theme, colours and a style set, and restores the style set's spacing", () => {
    const model = editor(["a"]);
    runCommand(model, c.applyThemeCommand, { theme: THEMES[3]! });
    runCommand(model, c.themeColorsCommand, { colors: THEME_COLOR_SCHEMES[5]! });
    runCommand(model, c.applyStyleSetCommand, { set: STYLE_SETS[2]! });
    runCommand(model, c.paragraphSpacingCommand, { spacing: { before: 0, after: 0, line: 240 } });
    runCommand(model, c.paragraphSpacingCommand, { spacing: "default" });
    const doc = saved(model);
    expect(getTheme(doc)?.name).toBe(THEMES[3]!.name);
    expect(getTheme(doc)?.colors.name).toBe(THEME_COLOR_SCHEMES[5]!.name);
    expect(getDefaultParagraphSpacing(doc)).toEqual({
      before: 0,
      after: STYLE_SETS[2]!.after,
      line: STYLE_SETS[2]!.line,
    });
    runCommand(model, c.resetThemeCommand, undefined);
    expect(getTheme(model.doc)?.name).toBe("Office Theme");
  });

  it("adds and removes a watermark, page colour and page borders", () => {
    const model = editor(["a"]);
    runCommand(model, c.watermarkCommand, {
      watermark: {
        kind: "text",
        text: "DRAFT",
        font: "Calibri",
        size: "auto",
        color: "C0C0C0",
        semitransparent: true,
        layout: "diagonal",
      },
    });
    runCommand(model, c.pageColorCommand, { color: { color: "DDEEFF" } });
    runCommand(model, c.pageBordersCommand, {
      borders: {
        offsetFrom: "text",
        display: "allPages",
        zOrder: "front",
        top: { style: "single", size: 4, spacePt: 24, color: "auto" },
      },
      target: "document",
    });
    const background = resolvePageBackground(saved(model));
    expect(background.color?.color).toBe("DDEEFF");
    expect(background.watermark).toMatchObject({ kind: "text", text: "DRAFT" });
    expect(getSectionProperties(model.doc).pageBorders?.top?.style).toBe("single");
    runCommand(model, c.watermarkCommand, { watermark: undefined });
    runCommand(model, c.pageBordersCommand, { borders: null, target: "section" });
    expect(getWatermark(saved(model))).toBeUndefined();
    expect(getSectionProperties(model.doc).pageBorders).toBeUndefined();
  });
});

describe("editor helpers", () => {
  it("resolves each section's block range", () => {
    const model = editor(["a", "b", "c"], { block: 1, inline: 0, offset: 0 });
    runCommand(model, c.insertSectionBreakCommand, { type: "nextPage" });
    const sections = resolveSectionLayout(model.doc);
    expect(sections.map((s) => [s.startBlock, s.endBlock])).toEqual([
      [0, 1],
      [2, 3],
    ]);
  });

  it("falls back to the Office theme and builds Word's 6×10 palette", () => {
    const theme = resolveTheme(createDocx());
    expect(theme.name).toBe("Office Theme");
    const palette = themePalette(theme.colors);
    expect(palette).toHaveLength(6);
    expect(palette.every((row) => row.length === 10)).toBe(true);
    expect(palette[0]?.[0]).toEqual({ themeColor: "background1", rgb: "FFFFFF" });
    expect(palette[1]?.[0]).toMatchObject({
      themeColor: "background1",
      themeShade: "F2",
      rgb: "F2F2F2",
    });
  });
});
