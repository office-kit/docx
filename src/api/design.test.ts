import { describe, expect, it } from "vitest";
import { getPart, hasPart } from "../internal/opc/index.js";
import {
  addHeader,
  applyStyleSet,
  createDocx,
  currentStyleSet,
  getDefaultParagraphSpacing,
  getDocumentSetting,
  getPageColor,
  getTheme,
  getWatermark,
  insertSectionBreak,
  openDocx,
  PARAGRAPH_SPACING_PRESETS,
  setDefaultParagraphSpacing,
  setDocumentSettingOnOff,
  setPageColor,
  setSectionProperties,
  setTheme,
  setThemeColors,
  setThemeEffects,
  setThemeFonts,
  setWatermark,
  STYLE_SETS,
  stylesPart,
  THEME_COLOR_SCHEMES,
  THEME_EFFECT_SCHEMES,
  THEME_FONT_SCHEMES,
  themeColorValue,
  THEMES,
  toUint8Array,
  validatePackage,
  type Docx,
} from "./index.js";

function reopen(doc: Docx): Docx {
  const next = openDocx(toUint8Array(doc));
  expect(validatePackage(next.opc).filter((i) => i.level === "error")).toEqual([]);
  return next;
}

function partXml(doc: Docx, name: string): string {
  return new TextDecoder().decode(getPart(doc.opc, name)?.data);
}

function pick<T extends { name: string }>(list: readonly T[], name: string): T {
  const found = list.find((x) => x.name === name);
  if (!found) throw new Error(name);
  return found;
}

// A 1×1 PNG.
const PNG = Uint8Array.from(
  atob(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==",
  ),
  (c) => c.charCodeAt(0),
);

describe("themes", () => {
  it("creates the theme part and relationship and round-trips every built-in theme", () => {
    for (const theme of THEMES) {
      const doc = createDocx();
      setTheme(doc, theme);
      const back = reopen(doc);
      expect(hasPart(back.opc, "/word/theme/theme1.xml")).toBe(true);
      const info = getTheme(back);
      expect(info?.name).toBe(theme.name);
      expect(info?.colors.accent1).toBe(theme.colors.accent1);
      expect(info?.fonts.major).toBe(theme.fonts.major);
      expect(info?.effects).toBe(theme.effects.name);
    }
  });

  it("writes Office's text and background colours as system colours", () => {
    const doc = createDocx();
    setTheme(doc, pick(THEMES, "Office Theme"));
    const xml = partXml(reopen(doc), "/word/theme/theme1.xml");
    expect(xml).toContain('<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>');
    expect(xml).toContain("<a:objectDefaults/><a:extraClrSchemeLst/>");
  });

  it("replaces one scheme at a time, creating the Office theme first", () => {
    const doc = createDocx();
    setThemeColors(doc, pick(THEME_COLOR_SCHEMES, "Green"));
    setThemeFonts(doc, pick(THEME_FONT_SCHEMES, "Garamond"));
    setThemeEffects(doc, pick(THEME_EFFECT_SCHEMES, "Office 2007 - 2010"));
    const info = getTheme(reopen(doc));
    expect(info?.name).toBe("Office Theme");
    expect(info?.colors.name).toBe("Green");
    expect(info?.colors.accent1).toBe("549E39");
    expect(info?.fonts).toEqual({ name: "Garamond", major: "Garamond", minor: "Garamond" });
    expect(info?.effects).toBe("Office 2007 - 2010");
  });

  it("recomputes stored theme colour values when the colours change", () => {
    const doc = createDocx();
    applyStyleSet(doc, pick(STYLE_SETS, "Default"));
    setThemeColors(doc, pick(THEME_COLOR_SCHEMES, "Red"));
    const styles = partXml(reopen(doc), "/word/styles.xml");
    const expected = themeColorValue(pick(THEME_COLOR_SCHEMES, "Red"), "accent1", { shade: "BF" });
    expect(styles).toContain(`w:val="${expected}" w:themeColor="accent1" w:themeShade="BF"`);
  });

  it("derives shades and tints from HSL luminance", () => {
    const office = pick(THEME_COLOR_SCHEMES, "Office 2013 - 2022");
    expect(themeColorValue(office, "accent1")).toBe("4472C4");
    expect(themeColorValue(office, "text1", { tint: "80" })).toBe("7F7F7F");
    expect(themeColorValue(office, "background1", { shade: "BF" })).toBe("BFBFBF");
    expect(themeColorValue(office, "nope")).toBeUndefined();
  });
});

describe("style sets and paragraph spacing", () => {
  it("rewrites the defaults and styles, and recognizes the set afterwards", () => {
    for (const set of STYLE_SETS) {
      const doc = createDocx();
      applyStyleSet(doc, set);
      const back = reopen(doc);
      expect(currentStyleSet(back)?.name).toBe(set.name);
      expect(getDefaultParagraphSpacing(back)).toEqual({
        before: 0,
        after: set.after,
        line: set.line,
      });
    }
  });

  it("keeps style element children in schema order", () => {
    const doc = createDocx();
    applyStyleSet(doc, pick(STYLE_SETS, "Lines (Distinctive)"));
    const styles = partXml(reopen(doc), "/word/styles.xml");
    const heading = /<w:style [^>]*w:styleId="Heading1">(.*?)<\/w:style>/s.exec(styles)?.[1] ?? "";
    const order = [...heading.matchAll(/<w:(\w+)/g)].map((m) => m[1]);
    expect(order).toEqual([
      "name",
      "basedOn",
      "next",
      "uiPriority",
      "qFormat",
      "pPr",
      "keepNext",
      "keepLines",
      "pBdr",
      "bottom",
      "spacing",
      "outlineLvl",
      "rPr",
      "rFonts",
      "b",
      "bCs",
      "color",
      "sz",
      "szCs",
    ]);
  });

  it("sets the default paragraph spacing", () => {
    const doc = createDocx();
    const preset = pick(PARAGRAPH_SPACING_PRESETS, "Relaxed");
    setDefaultParagraphSpacing(doc, preset);
    expect(getDefaultParagraphSpacing(reopen(doc))).toEqual({ before: 0, after: 120, line: 360 });
    expect(stylesPart(doc)?.styles.length).toBeGreaterThan(0);
    expect(() => setDefaultParagraphSpacing(doc, { before: -1, after: 0, line: 240 })).toThrow(
      RangeError,
    );
  });
});

describe("page colour", () => {
  it("writes w:background before the body and turns on displayBackgroundShape", () => {
    const doc = createDocx({ paragraphs: ["x"] });
    setPageColor(doc, { color: "fff2cc", themeColor: "accent4", themeTint: "33" });
    const back = reopen(doc);
    const xml = partXml(back, "/word/document.xml");
    expect(xml).toMatch(
      /<w:background w:color="FFF2CC" w:themeColor="accent4" w:themeTint="33"\/><w:body>/,
    );
    expect(getDocumentSetting(back, "displayBackgroundShape").present).toBe(true);
    expect(getPageColor(back)).toEqual({ color: "FFF2CC", themeColor: "accent4", themeTint: "33" });
  });

  it("writes a VML gradient fill and replaces the previous background", () => {
    const doc = createDocx();
    setPageColor(doc, { color: "FF0000" });
    setPageColor(doc, {
      color: "FFFF00",
      gradient: { color2: "00B0F0", style: "vertical", variant: 3 },
    });
    const back = reopen(doc);
    const xml = partXml(back, "/word/document.xml");
    expect(xml.match(/<w:background/g)).toHaveLength(1);
    expect(xml).toContain('xmlns:v="urn:schemas-microsoft-com:vml"');
    expect(getPageColor(back)?.gradient).toEqual({
      color2: "00B0F0",
      style: "vertical",
      variant: 3,
    });
  });

  it("removes the background", () => {
    const doc = createDocx();
    setPageColor(doc, { color: "FF0000" });
    setPageColor(doc, undefined);
    const back = reopen(doc);
    expect(partXml(back, "/word/document.xml")).not.toContain("w:background");
    expect(getPageColor(back)).toBeUndefined();
    expect(getDocumentSetting(back, "displayBackgroundShape").present).toBe(false);
  });
});

describe("watermarks", () => {
  it("creates headers for every header type in use and puts a text watermark in each", () => {
    const doc = createDocx({ paragraphs: ["a", "b"] });
    setSectionProperties(doc, { titlePage: true });
    setDocumentSettingOnOff(doc, "evenAndOddHeaders", true);
    insertSectionBreak(doc, 0);
    setWatermark(doc, {
      kind: "text",
      text: "CONFIDENTIAL",
      font: "Calibri",
      size: "auto",
      color: "C0C0C0",
      semitransparent: true,
      layout: "diagonal",
    });
    const back = reopen(doc);
    for (const n of [1, 2, 3]) {
      const xml = partXml(back, `/word/header${n}.xml`);
      expect(xml).toContain('<w:docPartGallery w:val="Watermarks"/>');
      expect(xml).toContain('type="#_x0000_t136"');
      expect(xml).toContain('string="CONFIDENTIAL"');
      expect(xml).toContain('xmlns:w10="urn:schemas-microsoft-com:office:word"');
    }
    expect(getWatermark(back)).toEqual({
      kind: "text",
      text: "CONFIDENTIAL",
      font: "Calibri",
      size: "auto",
      color: "C0C0C0",
      semitransparent: true,
      layout: "diagonal",
    });
  });

  it("adds the watermark to an existing header and replaces an earlier one", () => {
    const doc = createDocx();
    addHeader(doc, "Header text");
    const text = {
      kind: "text",
      text: "DRAFT",
      font: "Arial",
      size: 72,
      color: "FF0000",
      semitransparent: false,
      layout: "horizontal",
    } as const;
    setWatermark(doc, text);
    setWatermark(doc, { ...text, text: "SAMPLE" });
    const back = reopen(doc);
    const xml = partXml(back, "/word/header1.xml");
    expect(xml).toContain("Header text");
    expect(xml).not.toContain("DRAFT");
    expect(xml.match(/PowerPlusWaterMarkObject/g)).toHaveLength(1);
    expect(getWatermark(back)).toMatchObject({
      text: "SAMPLE",
      size: 72,
      color: "FF0000",
      layout: "horizontal",
    });
  });

  it("writes a washed-out picture watermark with its image relationship", () => {
    const doc = createDocx();
    setWatermark(doc, { kind: "picture", bytes: PNG, widthPt: 200, heightPt: 100, washout: true });
    const back = reopen(doc);
    const xml = partXml(back, "/word/header1.xml");
    expect(xml).toContain('gain="19661f" blacklevel="22938f"');
    expect(partXml(back, "/word/_rels/header1.xml.rels")).toContain("media/image1.png");
    expect(getWatermark(back)).toEqual({
      kind: "picture",
      widthPt: 200,
      heightPt: 100,
      washout: true,
      partName: "/word/media/image1.png",
    });
  });

  it("removes the watermark and keeps the header", () => {
    const doc = createDocx();
    addHeader(doc, "Keep me");
    setWatermark(doc, {
      kind: "text",
      text: "ASAP",
      font: "Calibri",
      size: "auto",
      color: "C0C0C0",
      semitransparent: true,
      layout: "diagonal",
    });
    setWatermark(doc, undefined);
    const back = reopen(doc);
    expect(getWatermark(back)).toBeUndefined();
    expect(partXml(back, "/word/header1.xml")).toContain("Keep me");
  });
});
