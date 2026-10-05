/**
 * Document design: the theme part (colours, fonts, effects), style sets,
 * default paragraph spacing, and the page background (page colour and
 * watermarks) — Word's Design tab.
 *
 * Page borders are a section property; see `setSectionProperties`.
 */

import {
  addPart,
  addRelationship,
  hasPart,
  normalizePartName,
  partRelationships,
  relationshipsByType,
  resolveInternalTarget,
  setContentTypeDefault,
} from "../internal/opc/index.js";
import {
  extensionForImageContentType,
  findStyle,
  MINIMAL_STYLES_XML,
  parseStylesPart,
  sniffImageContentType,
  WML_CONTENT_TYPES,
  WML_NS,
  WML_RELATIONSHIPS,
  type WmlStylesPart,
} from "../internal/wordprocessingml/index.js";
import { insertOrderedChild, SECT_PR_ORDER } from "../internal/wordprocessingml/schema-order.js";
import { sectionPropertiesList } from "../internal/wordprocessingml/sections.js";
import {
  attrOf,
  intAttrOf,
  onOffChild,
  wmlChild,
  wmlChildren,
  wmlElement,
} from "../internal/wordprocessingml/wml-element.js";
import {
  parseXml,
  serializeXml,
  type XmlAttr,
  type XmlElement,
  type XmlNode,
} from "../internal/xml/index.js";
import {
  getDocumentSetting,
  getRawPartRoot,
  markRawPartDirty,
  setDocumentSettingOnOff,
  stylesPart,
  type Docx,
} from "./docx.js";
import {
  STYLE_SETS,
  THEME_COLOR_SLOTS,
  THEMES,
  type StyleColor,
  type StyleSetBorder,
  type StyleSetDefinition,
  type StyleSetStyle,
  type StyleSetStyleId,
  type ThemeColorScheme,
  type ThemeDefinition,
  type ThemeEffectScheme,
  type ThemeFontScheme,
} from "./design-presets.js";

export {
  PARAGRAPH_SPACING_PRESETS,
  type ParagraphSpacingPreset,
  STYLE_SETS,
  type StyleColor,
  type StyleSetBorder,
  type StyleSetDefinition,
  type StyleSetStyle,
  type StyleSetStyleId,
  type StyleThemeColor,
  THEME_COLOR_SCHEMES,
  THEME_COLOR_SLOTS,
  THEME_EFFECT_SCHEMES,
  THEME_FONT_SCHEMES,
  type ThemeColorScheme,
  type ThemeDefinition,
  type ThemeEffectScheme,
  type ThemeFontScheme,
  THEMES,
} from "./design-presets.js";

const DML_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const VML_NS = "urn:schemas-microsoft-com:vml";
const OFFICE_NS = "urn:schemas-microsoft-com:office:office";
const WORD_VML_NS = "urn:schemas-microsoft-com:office:word";
const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const DEFAULT_THEME_PART = "/word/theme/theme1.xml";
const STYLES_PART_NAME = "/word/styles.xml";
const HEX6 = /^[0-9A-Fa-f]{6}$/;

function esc(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function requireHex(value: string, what: string): string {
  if (!HEX6.test(value)) throw new RangeError(`${what} must be hex RRGGBB, got ${value}.`);
  return value.toUpperCase();
}

/** Parse a markup fragment whose root declares its own namespaces, dropping those declarations. */
function fragment(xml: string): XmlElement {
  const root = parseXml(xml).root;
  return { ...root, attrs: root.attrs.filter((a) => !a.isNamespaceDecl) };
}

/** Declare `prefix` on `root` unless it is already declared. */
function ensureNamespace(root: XmlElement, prefix: string, uri: string): void {
  if (root.attrs.some((a) => a.isNamespaceDecl && a.name.local === prefix)) return;
  (root.attrs as XmlAttr[]).push({
    name: { uri: "http://www.w3.org/2000/xmlns/", local: prefix, prefix: "xmlns" },
    value: uri,
    isNamespaceDecl: true,
  });
}

// --- Theme part ---------------------------------------------------------------

// Word's per-script fonts (§20.1.4.1.16) for the Office font schemes; the
// built-in schemes share them, so every scheme written here carries the same set.
const SCRIPT_FONTS: Readonly<Record<"major" | "minor", ReadonlyArray<readonly [string, string]>>> =
  {
    major: [
      ["Jpan", "游ゴシック Light"],
      ["Hang", "맑은 고딕"],
      ["Hans", "等线 Light"],
      ["Hant", "新細明體"],
      ["Arab", "Times New Roman"],
      ["Hebr", "Times New Roman"],
      ["Thai", "Angsana New"],
    ],
    minor: [
      ["Jpan", "游明朝"],
      ["Hang", "맑은 고딕"],
      ["Hans", "等线"],
      ["Hant", "新細明體"],
      ["Arab", "Arial"],
      ["Hebr", "Arial"],
      ["Thai", "Cordia New"],
    ],
  };

function clrSchemeXml(c: ThemeColorScheme): string {
  const slot = (key: (typeof THEME_COLOR_SLOTS)[number]): string => {
    const value = requireHex(c[key], `Theme colour ${key}`);
    // Word stores Office's text and background colours as system colours.
    if (key === "dk1" && value === "000000") return '<a:sysClr val="windowText" lastClr="000000"/>';
    if (key === "lt1" && value === "FFFFFF") return '<a:sysClr val="window" lastClr="FFFFFF"/>';
    return `<a:srgbClr val="${value}"/>`;
  };
  const slots = THEME_COLOR_SLOTS.map((k) => `<a:${k}>${slot(k)}</a:${k}>`).join("");
  return `<a:clrScheme name="${esc(c.name)}">${slots}</a:clrScheme>`;
}

function fontSchemeXml(f: ThemeFontScheme): string {
  const half = (tag: "majorFont" | "minorFont", latin: string, role: "major" | "minor"): string =>
    `<a:${tag}><a:latin typeface="${esc(latin)}" panose=""/><a:ea typeface=""/><a:cs typeface=""/>` +
    SCRIPT_FONTS[role]
      .map(([script, face]) => `<a:font script="${script}" typeface="${face}"/>`)
      .join("") +
    `</a:${tag}>`;
  return `<a:fontScheme name="${esc(f.name)}">${half("majorFont", f.major, "major")}${half("minorFont", f.minor, "minor")}</a:fontScheme>`;
}

function fmtSchemeXml(e: ThemeEffectScheme): string {
  return `<a:fmtScheme name="${esc(e.name)}">${e.styles}</a:fmtScheme>`;
}

function themeXml(theme: ThemeDefinition): string {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
    `<a:theme xmlns:a="${DML_NS}" name="${esc(theme.name)}"><a:themeElements>` +
    clrSchemeXml(theme.colors) +
    fontSchemeXml(theme.fonts) +
    fmtSchemeXml(theme.effects) +
    "</a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>"
  );
}

/** The theme part's name, if the document has one (the main part's theme relationship). */
function themePartName(doc: Docx): string | undefined {
  const rels = partRelationships(doc.opc, doc.partName);
  const rel = relationshipsByType(rels, WML_RELATIONSHIPS.theme)[0];
  if (!rel) return undefined;
  const name = resolveInternalTarget(normalizePartName(doc.partName), rel.target);
  return hasPart(doc.opc, name) ? name : undefined;
}

function themeRoot(doc: Docx): XmlElement | undefined {
  const name = themePartName(doc);
  return name ? getRawPartRoot(doc, name) : undefined;
}

/**
 * Write `theme` as the document's theme part, replacing the current theme or
 * creating `/word/theme/theme1.xml` and its relationship when there is none.
 */
export function setTheme(doc: Docx, theme: ThemeDefinition): void {
  const xml = themeXml(theme);
  let name = themePartName(doc);
  if (!name) {
    name = DEFAULT_THEME_PART;
    addPart(doc.opc, {
      name,
      contentType: WML_CONTENT_TYPES.theme,
      data: new TextEncoder().encode(xml),
    });
    addRelationship(partRelationships(doc.opc, doc.partName), {
      type: WML_RELATIONSHIPS.theme,
      target: "theme/theme1.xml",
    });
  }
  // The raw-part cache is the theme's single in-memory copy (the canvas and
  // the raw-XML inspector read it there), so the new tree replaces it.
  doc.rawParts.set(name, parseXml(xml));
  markRawPartDirty(doc, name);
  doc.dirty = true;
}

/** The theme's `<a:themeElements>`, creating the Office theme first when the document has none. */
function themeElements(doc: Docx): XmlElement {
  if (!themeRoot(doc)) setTheme(doc, officeTheme());
  const root = themeRoot(doc);
  const elements = root?.children.find(
    (c): c is XmlElement => c.kind === "element" && c.name.local === "themeElements",
  );
  if (!elements) throw new Error("The theme part has no <a:themeElements>.");
  return elements;
}

function officeTheme(): ThemeDefinition {
  const office = THEMES[0];
  if (!office) throw new Error("The built-in theme list is empty.");
  return office;
}

function replaceSchemeElement(doc: Docx, local: string, xml: string): void {
  const elements = themeElements(doc);
  const replacement = fragment(xml.replace(`<a:${local} `, `<a:${local} xmlns:a="${DML_NS}" `));
  const children = elements.children as XmlNode[];
  const index = children.findIndex((c) => c.kind === "element" && c.name.local === local);
  if (index >= 0) children[index] = replacement;
  else children.push(replacement);
  const name = themePartName(doc);
  if (name) markRawPartDirty(doc, name);
  doc.dirty = true;
}

/**
 * Replace the theme's colour scheme (Design ▸ Colors). Run, paragraph and
 * table colours that reference a theme colour have their stored RGB fallback
 * recomputed from the new scheme, as Word does.
 */
export function setThemeColors(doc: Docx, colors: ThemeColorScheme): void {
  replaceSchemeElement(doc, "clrScheme", clrSchemeXml(colors));
  refreshThemeColorValues(doc, colors);
}

/** Replace the theme's font scheme (Design ▸ Fonts). */
export function setThemeFonts(doc: Docx, fonts: ThemeFontScheme): void {
  replaceSchemeElement(doc, "fontScheme", fontSchemeXml(fonts));
}

/** Replace the theme's effect scheme (Design ▸ Effects). */
export function setThemeEffects(doc: Docx, effects: ThemeEffectScheme): void {
  replaceSchemeElement(doc, "fmtScheme", fmtSchemeXml(effects));
}

/** The document's theme, as names and values; `undefined` when it has no theme part. */
export interface ThemeInfo {
  readonly name: string;
  readonly colors: ThemeColorScheme;
  readonly fonts: ThemeFontScheme;
  /** The `<a:fmtScheme>` name. */
  readonly effects: string;
}

function readColor(slot: XmlElement | undefined): string {
  const value = slot?.children.find((c): c is XmlElement => c.kind === "element");
  if (!value) return "000000";
  return (
    (value.name.local === "sysClr"
      ? attrOf(value, "lastClr")
      : attrOf(value, "val")
    )?.toUpperCase() ?? "000000"
  );
}

function dmlChild(el: XmlElement | undefined, local: string): XmlElement | undefined {
  return el?.children.find((c): c is XmlElement => c.kind === "element" && c.name.local === local);
}

/** Read the document's theme (Design ▸ Themes / Colors / Fonts / Effects state). */
export function getTheme(doc: Docx): ThemeInfo | undefined {
  const root = themeRoot(doc);
  if (!root) return undefined;
  const elements = dmlChild(root, "themeElements");
  const clr = dmlChild(elements, "clrScheme");
  const font = dmlChild(elements, "fontScheme");
  const fmt = dmlChild(elements, "fmtScheme");
  const slot = (k: (typeof THEME_COLOR_SLOTS)[number]): string => readColor(dmlChild(clr, k));
  const latin = (half: string): string =>
    attrOf(dmlChild(dmlChild(font, half), "latin"), "typeface") ?? "";
  return {
    name: attrOf(root, "name") ?? "",
    colors: {
      name: attrOf(clr, "name") ?? "",
      dk1: slot("dk1"),
      lt1: slot("lt1"),
      dk2: slot("dk2"),
      lt2: slot("lt2"),
      accent1: slot("accent1"),
      accent2: slot("accent2"),
      accent3: slot("accent3"),
      accent4: slot("accent4"),
      accent5: slot("accent5"),
      accent6: slot("accent6"),
      hlink: slot("hlink"),
      folHlink: slot("folHlink"),
    },
    fonts: {
      name: attrOf(font, "name") ?? "",
      major: latin("majorFont"),
      minor: latin("minorFont"),
    },
    effects: attrOf(fmt, "name") ?? "",
  };
}

// --- Theme colour resolution -----------------------------------------------------

// `w:themeColor` (ST_ThemeColor, §17.18.97) → clrScheme slot, through Word's
// default colour mapping (bg1 = lt1, tx1 = dk1, bg2 = lt2, tx2 = dk2).
const THEME_COLOR_SLOT: Readonly<Record<string, (typeof THEME_COLOR_SLOTS)[number]>> = {
  dark1: "dk1",
  light1: "lt1",
  dark2: "dk2",
  light2: "lt2",
  accent1: "accent1",
  accent2: "accent2",
  accent3: "accent3",
  accent4: "accent4",
  accent5: "accent5",
  accent6: "accent6",
  hyperlink: "hlink",
  followedHyperlink: "folHlink",
  background1: "lt1",
  text1: "dk1",
  background2: "lt2",
  text2: "dk2",
};

function rgbToHsl(hex: string): [number, number, number] {
  const n = Number.parseInt(hex, 16);
  const r = ((n >> 16) & 0xff) / 255;
  const g = ((n >> 8) & 0xff) / 255;
  const b = (n & 0xff) / 255;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) return [0, 0, l];
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  const h =
    max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
  return [h / 6, s, l];
}

function hue(p: number, q: number, t: number): number {
  const u = t < 0 ? t + 1 : t > 1 ? t - 1 : t;
  if (u < 1 / 6) return p + (q - p) * 6 * u;
  if (u < 1 / 2) return q;
  if (u < 2 / 3) return p + (q - p) * (2 / 3 - u) * 6;
  return p;
}

function hslToRgb(h: number, s: number, l: number): string {
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  const channels = s === 0 ? [l, l, l] : [hue(p, q, h + 1 / 3), hue(p, q, h), hue(p, q, h - 1 / 3)];
  return channels
    .map((v) =>
      Math.round(v * 255)
        .toString(16)
        .padStart(2, "0"),
    )
    .join("")
    .toUpperCase();
}

/**
 * The RGB value of a WordprocessingML theme colour reference
 * (`w:themeColor` + optional `w:themeShade` / `w:themeTint`, hex 00–FF) under
 * `colors`. Shades darken and tints lighten the colour's HSL luminance, which
 * is how Word derives the "Lighter 40%" / "Darker 25%" palette entries.
 * Returns `undefined` for an unknown theme colour name.
 */
export function themeColorValue(
  colors: ThemeColorScheme,
  themeColor: string,
  modifiers: { readonly shade?: string | undefined; readonly tint?: string | undefined } = {},
): string | undefined {
  const slot = THEME_COLOR_SLOT[themeColor];
  if (!slot) return undefined;
  const base = colors[slot].toUpperCase();
  const shade =
    modifiers.shade === undefined ? undefined : Number.parseInt(modifiers.shade, 16) / 255;
  const tint = modifiers.tint === undefined ? undefined : Number.parseInt(modifiers.tint, 16) / 255;
  if (shade === undefined && tint === undefined) return base;
  const [h, s, l] = rgbToHsl(base);
  let lum = l;
  if (shade !== undefined && Number.isFinite(shade)) lum *= shade;
  if (tint !== undefined && Number.isFinite(tint)) lum = lum * tint + (1 - tint);
  return hslToRgb(h, s, Math.min(1, Math.max(0, lum)));
}

// Elements whose colour attribute pairs with theme attributes (§17.3.2.6,
// §17.3.5 borders, §17.3.5 shd fill).
const THEMED_COLOR_ATTRS: ReadonlyArray<readonly [string, string, string, string]> = [
  ["color", "themeColor", "themeShade", "themeTint"],
  ["fill", "themeFill", "themeFillShade", "themeFillTint"],
];

function refreshElement(el: XmlElement, colors: ThemeColorScheme): void {
  for (const [valueAttr, themeAttr, shadeAttr, tintAttr] of THEMED_COLOR_ATTRS) {
    const theme = attrOf(el, themeAttr);
    if (theme === undefined) continue;
    // `w:color` stores its colour in `w:val`; borders and shading in `w:color` / `w:fill`.
    const target = el.name.local === "color" && valueAttr === "color" ? "val" : valueAttr;
    const value = themeColorValue(colors, theme, {
      shade: attrOf(el, shadeAttr),
      tint: attrOf(el, tintAttr),
    });
    const attr = el.attrs.find((a) => a.name.local === target);
    if (value && attr) (attr as { value: string }).value = value;
  }
  for (const child of el.children) if (child.kind === "element") refreshElement(child, colors);
}

function refreshThemeColorValues(doc: Docx, colors: ThemeColorScheme): void {
  const styles = stylesPart(doc);
  if (styles) {
    if (styles.docDefaults) refreshElement(styles.docDefaults, colors);
    for (const style of styles.styles) refreshElement(style, colors);
    doc.stylesDirty = true;
  }
  const visitParagraphs = (blocks: Docx["document"]["body"]["blocks"]): void => {
    for (const block of blocks) {
      if (block.kind === "raw") refreshElement(block.node, colors);
      else if (block.kind === "table") {
        if (block.tblPr) refreshElement(block.tblPr, colors);
        for (const row of block.rows) {
          for (const cell of row.cells) {
            if (cell.tcPr) refreshElement(cell.tcPr, colors);
            visitParagraphs(cell.paragraphs);
          }
        }
      } else {
        if (block.pPr) refreshElement(block.pPr, colors);
        for (const inline of block.children) {
          if (inline.kind === "run") {
            if (inline.rPr) refreshElement(inline.rPr, colors);
          } else if (inline.kind === "raw") {
            refreshElement(inline.node, colors);
          }
        }
      }
    }
  };
  visitParagraphs(doc.document.body.blocks);
  const background = pageBackgroundElement(doc);
  if (background) refreshElement(background, colors);
  doc.dirty = true;
}

// --- Style sets and paragraph spacing ----------------------------------------------

function ensureStyles(doc: Docx): WmlStylesPart {
  const existing = stylesPart(doc);
  if (existing) return existing;
  addPart(doc.opc, {
    name: STYLES_PART_NAME,
    contentType: WML_CONTENT_TYPES.styles,
    data: new TextEncoder().encode(MINIMAL_STYLES_XML),
  });
  const rels = partRelationships(doc.opc, doc.partName);
  if (relationshipsByType(rels, WML_RELATIONSHIPS.styles).length === 0) {
    addRelationship(rels, { type: WML_RELATIONSHIPS.styles, target: "styles.xml" });
  }
  doc.stylesCache = parseStylesPart(parseXml(MINIMAL_STYLES_XML));
  doc.stylesDirty = true;
  return doc.stylesCache;
}

function colorAttrs(c: StyleColor, valueAttr: string): Record<string, string | undefined> {
  return {
    [valueAttr]: requireHex(c.rgb, "Style colour"),
    [valueAttr === "fill" ? "themeFill" : "themeColor"]: c.theme,
    [valueAttr === "fill" ? "themeFillShade" : "themeShade"]: c.shade,
    [valueAttr === "fill" ? "themeFillTint" : "themeTint"]: c.tint,
  };
}

function borderEl(side: "top" | "bottom", b: StyleSetBorder): XmlElement {
  return wmlElement(side, {
    val: b.style,
    sz: b.size,
    space: b.space,
    ...colorAttrs(b.color, "color"),
  });
}

const THEME_FONT_ATTRS = {
  major: {
    asciiTheme: "majorHAnsi",
    eastAsiaTheme: "majorEastAsia",
    hAnsiTheme: "majorHAnsi",
    cstheme: "majorBidi",
  },
  minor: {
    asciiTheme: "minorHAnsi",
    eastAsiaTheme: "minorEastAsia",
    hAnsiTheme: "minorHAnsi",
    cstheme: "minorBidi",
  },
} as const;

/** `<w:pPr>` / `<w:rPr>` children of a style-set style, in CT_PPrBase / CT_RPr order. */
function styleSetProps(
  s: StyleSetStyle,
  outlineLevel: number | undefined,
): { pPr: XmlElement[]; rPr: XmlElement[] } {
  const pPr: XmlElement[] = [];
  if (outlineLevel !== undefined) pPr.push(wmlElement("keepNext"), wmlElement("keepLines"));
  if (s.borderTop || s.borderBottom) {
    pPr.push(
      wmlElement("pBdr", {}, [
        ...(s.borderTop ? [borderEl("top", s.borderTop)] : []),
        ...(s.borderBottom ? [borderEl("bottom", s.borderBottom)] : []),
      ]),
    );
  }
  if (s.shading)
    pPr.push(wmlElement("shd", { val: "clear", color: "auto", ...colorAttrs(s.shading, "fill") }));
  if (s.before !== undefined || s.after !== undefined || s.line !== undefined) {
    pPr.push(
      wmlElement("spacing", {
        before: s.before,
        after: s.after,
        line: s.line,
        lineRule: s.line === undefined ? undefined : "auto",
      }),
    );
  }
  if (s.contextualSpacing) pPr.push(wmlElement("contextualSpacing"));
  if (s.alignment) pPr.push(wmlElement("jc", { val: s.alignment }));
  if (outlineLevel !== undefined) pPr.push(wmlElement("outlineLvl", { val: outlineLevel }));
  const rPr: XmlElement[] = [];
  if (s.font) rPr.push(wmlElement("rFonts", THEME_FONT_ATTRS[s.font]));
  if (s.bold) rPr.push(wmlElement("b"), wmlElement("bCs"));
  if (s.italic) rPr.push(wmlElement("i"), wmlElement("iCs"));
  if (s.caps) rPr.push(wmlElement("caps"));
  if (s.smallCaps) rPr.push(wmlElement("smallCaps"));
  if (s.color) rPr.push(wmlElement("color", colorAttrs(s.color, "val")));
  if (s.characterSpacing !== undefined)
    rPr.push(wmlElement("spacing", { val: s.characterSpacing }));
  if (s.kern !== undefined) rPr.push(wmlElement("kern", { val: s.kern }));
  if (s.sizeHalfPoints !== undefined) {
    rPr.push(
      wmlElement("sz", { val: s.sizeHalfPoints }),
      wmlElement("szCs", { val: s.sizeHalfPoints }),
    );
  }
  return { pPr, rPr };
}

const STYLE_SET_META: Readonly<
  Record<StyleSetStyleId, { name: string; uiPriority: number; outline?: number }>
> = {
  Normal: { name: "Normal", uiPriority: 0 },
  Title: { name: "Title", uiPriority: 10 },
  Subtitle: { name: "Subtitle", uiPriority: 11 },
  Heading1: { name: "heading 1", uiPriority: 9, outline: 0 },
  Heading2: { name: "heading 2", uiPriority: 9, outline: 1 },
  Heading3: { name: "heading 3", uiPriority: 9, outline: 2 },
};

const STYLE_SET_IDS = /* @__PURE__ */ Object.keys(STYLE_SET_META) as StyleSetStyleId[];

function buildStyleSetStyle(id: StyleSetStyleId, s: StyleSetStyle): XmlElement {
  const meta = STYLE_SET_META[id];
  const { pPr, rPr } = styleSetProps(s, meta.outline);
  const normal = id === "Normal";
  return wmlElement("style", { type: "paragraph", default: normal ? 1 : undefined, styleId: id }, [
    wmlElement("name", { val: meta.name }),
    ...(normal
      ? []
      : [wmlElement("basedOn", { val: "Normal" }), wmlElement("next", { val: "Normal" })]),
    ...(normal ? [] : [wmlElement("uiPriority", { val: meta.uiPriority })]),
    wmlElement("qFormat"),
    ...(pPr.length > 0 ? [wmlElement("pPr", {}, pPr)] : []),
    ...(rPr.length > 0 ? [wmlElement("rPr", {}, rPr)] : []),
  ]);
}

function buildDocDefaults(set: StyleSetDefinition, existing: XmlElement | undefined): XmlElement {
  // Keep the document's language defaults; the style set does not change them.
  const lang = wmlChild(wmlChild(wmlChild(existing, "rPrDefault"), "rPr"), "lang");
  return wmlElement("docDefaults", {}, [
    wmlElement("rPrDefault", {}, [
      wmlElement("rPr", {}, [
        wmlElement("rFonts", THEME_FONT_ATTRS.minor),
        wmlElement("kern", { val: 2 }),
        wmlElement("sz", { val: set.bodySizeHalfPoints }),
        wmlElement("szCs", { val: set.bodySizeHalfPoints }),
        ...(lang ? [lang] : []),
      ]),
    ]),
    wmlElement("pPrDefault", {}, [
      wmlElement("pPr", {}, [
        wmlElement("spacing", { after: set.after, line: set.line, lineRule: "auto" }),
      ]),
    ]),
  ]);
}

/**
 * Apply a style set (Design ▸ Document Formatting): rewrite the document
 * defaults and the Normal, Title, Subtitle and Heading 1–3 styles. Fonts
 * reference the theme (`asciiTheme`), so the set follows Design ▸ Fonts.
 */
export function applyStyleSet(doc: Docx, set: StyleSetDefinition): void {
  const part = ensureStyles(doc);
  part.docDefaults = buildDocDefaults(set, part.docDefaults);
  for (const id of STYLE_SET_IDS) {
    const el = buildStyleSetStyle(id, set.styles[id] ?? {});
    const existing = findStyle(part, id);
    const index = existing ? part.styles.indexOf(existing) : -1;
    if (index >= 0) part.styles[index] = el;
    else part.styles.push(el);
  }
  doc.stylesDirty = true;
  doc.dirty = true;
}

function serializeElement(el: XmlElement | undefined): string {
  return el ? serializeXml({ prologue: [], root: el, epilogue: [] }) : "";
}

/**
 * The built-in style set the document's styles match exactly, if any (Word
 * highlights it in the gallery). Only sets applied by {@link applyStyleSet}
 * match, since the comparison is on the written markup.
 */
export function currentStyleSet(doc: Docx): StyleSetDefinition | undefined {
  const part = stylesPart(doc);
  if (!part) return undefined;
  const actual = STYLE_SET_IDS.map((id) => serializeElement(findStyle(part, id)));
  return STYLE_SETS.find((set) =>
    STYLE_SET_IDS.every(
      (id, i) => actual[i] === serializeElement(buildStyleSetStyle(id, set.styles[id] ?? {})),
    ),
  );
}

/** Default paragraph spacing (`w:pPrDefault`): twips before/after, line in 240ths of a line. */
export interface DefaultParagraphSpacing {
  readonly before: number;
  readonly after: number;
  readonly line: number;
}

/**
 * Set the document's default paragraph spacing (Design ▸ Paragraph
 * Spacing), stored in `<w:docDefaults>` so every style that does not set its
 * own spacing follows it.
 */
export function setDefaultParagraphSpacing(doc: Docx, spacing: DefaultParagraphSpacing): void {
  const fields = { before: spacing.before, after: spacing.after, line: spacing.line };
  for (const [what, v] of Object.entries(fields)) {
    if (!Number.isInteger(v) || v < 0 || v > 31680)
      throw new RangeError(`${what} must be an integer in 0..31680.`);
  }
  const part = ensureStyles(doc);
  const docDefaults = part.docDefaults ?? wmlElement("docDefaults");
  part.docDefaults = docDefaults;
  let pPrDefault = wmlChild(docDefaults, "pPrDefault");
  if (!pPrDefault) {
    pPrDefault = wmlElement("pPrDefault");
    (docDefaults.children as XmlNode[]).push(pPrDefault);
  }
  let pPr = wmlChild(pPrDefault, "pPr");
  if (!pPr) {
    pPr = wmlElement("pPr");
    (pPrDefault.children as XmlNode[]).push(pPr);
  }
  const el = wmlElement("spacing", {
    before: spacing.before === 0 ? undefined : spacing.before,
    after: spacing.after,
    line: spacing.line,
    lineRule: "auto",
  });
  const children = pPr.children as XmlNode[];
  const index = children.findIndex((c) => c.kind === "element" && c.name.local === "spacing");
  if (index >= 0) children[index] = el;
  else children.push(el);
  doc.stylesDirty = true;
  doc.dirty = true;
}

/** The document's default paragraph spacing, or `undefined` when none is set. */
export function getDefaultParagraphSpacing(doc: Docx): DefaultParagraphSpacing | undefined {
  const spacing = wmlChild(
    wmlChild(wmlChild(stylesPart(doc)?.docDefaults, "pPrDefault"), "pPr"),
    "spacing",
  );
  if (!spacing) return undefined;
  return {
    before: intAttrOf(spacing, "before") ?? 0,
    after: intAttrOf(spacing, "after") ?? 0,
    line: intAttrOf(spacing, "line") ?? 240,
  };
}

// --- Page colour -------------------------------------------------------------------

/** A two-colour gradient page fill (Design ▸ Page Color ▸ Fill Effects ▸ Gradient). */
export interface PageGradient {
  /** Hex RGB of the second colour. */
  readonly color2: string;
  readonly style: "horizontal" | "vertical" | "diagonalUp" | "diagonalDown";
  /** Word's four variants: 1 = colour → colour 2, 2 = reversed, 3 / 4 = mirrored from the middle. */
  readonly variant: 1 | 2 | 3 | 4;
}

/** The page background (`w:background`, §17.2.1). */
export interface PageColor {
  /** Hex RGB. */
  readonly color: string;
  readonly themeColor?: string;
  readonly themeShade?: string;
  readonly themeTint?: string;
  readonly gradient?: PageGradient;
}

// VML gradient angles and focus values for Word's shading styles and variants.
const GRADIENT_ANGLE: Readonly<Record<PageGradient["style"], number | undefined>> = {
  horizontal: undefined,
  vertical: -90,
  diagonalUp: -45,
  diagonalDown: -135,
};
const GRADIENT_FOCUS: Readonly<Record<PageGradient["variant"], string | undefined>> = {
  1: undefined,
  2: "100%",
  3: "50%",
  4: "-50%",
};

function pageBackgroundElement(doc: Docx): XmlElement | undefined {
  for (const extra of doc.document.extras) {
    const node = extra.node;
    if (node.kind === "element" && node.name.uri === WML_NS && node.name.local === "background")
      return node;
  }
  return undefined;
}

/**
 * Set (or with `undefined`, remove) the page colour. Word only paints the
 * background when `w:displayBackgroundShape` is on in the settings, so this
 * turns that on and off with it.
 */
export function setPageColor(doc: Docx, color: PageColor | undefined): void {
  let background: XmlElement | undefined;
  if (color) {
    const hex = requireHex(color.color, "Page colour");
    const children: XmlElement[] = [];
    if (color.gradient) {
      const g = color.gradient;
      const focus = GRADIENT_FOCUS[g.variant];
      const angle = GRADIENT_ANGLE[g.style];
      children.push(
        fragment(
          `<v:background xmlns:v="${VML_NS}" xmlns:o="${OFFICE_NS}" id="_x0000_s1025" o:bwmode="white" fillcolor="#${hex}" o:targetscreensize="1024,768">` +
            `<v:fill color2="#${requireHex(g.color2, "Gradient colour")}"${focus ? ` focus="${focus}"` : ""} type="gradient"${angle === undefined ? "" : ` angle="${angle}"`}/>` +
            "</v:background>",
        ),
      );
      ensureDocumentNamespace(doc, "v", VML_NS);
      ensureDocumentNamespace(doc, "o", OFFICE_NS);
    }
    background = wmlElement(
      "background",
      {
        color: hex,
        themeColor: color.themeColor,
        themeShade: color.themeShade,
        themeTint: color.themeTint,
      },
      children,
    );
  }
  // Pass-through slots index <w:document>'s child list, where <w:body> is the
  // one modelled child. Rebuild that list, swap the background, and number the
  // pass-through nodes again; <w:background> comes first (§17.2.3).
  const BODY = Symbol("body");
  const list: Array<XmlNode | typeof BODY> = [BODY];
  for (const extra of doc.document.extras.toSorted((a, b) => a.slot - b.slot)) {
    list.splice(Math.min(extra.slot, list.length), 0, extra.node);
  }
  const kept = list.filter(
    (n) =>
      n === BODY ||
      !(n.kind === "element" && n.name.uri === WML_NS && n.name.local === "background"),
  );
  if (background) kept.unshift(background);
  const extras = kept.flatMap((node, slot) => (node === BODY ? [] : [{ slot, node }]));
  doc.document = { ...doc.document, extras };
  setDocumentSettingOnOff(doc, "displayBackgroundShape", color !== undefined);
  doc.dirty = true;
}

function ensureDocumentNamespace(doc: Docx, prefix: string, uri: string): void {
  if (doc.document.rootAttrs.some((a) => a.isNamespaceDecl && a.name.local === prefix)) return;
  doc.document = {
    ...doc.document,
    rootAttrs: [
      ...doc.document.rootAttrs,
      {
        name: { uri: "http://www.w3.org/2000/xmlns/", local: prefix, prefix: "xmlns" },
        value: uri,
        isNamespaceDecl: true,
      },
    ],
  };
}

/** The page colour, or `undefined` when the page has none (or Word would not show it). */
export function getPageColor(doc: Docx): PageColor | undefined {
  const bg = pageBackgroundElement(doc);
  const color = attrOf(bg, "color");
  if (!bg || color === undefined || color === "auto") return undefined;
  const themeColor = attrOf(bg, "themeColor");
  const themeShade = attrOf(bg, "themeShade");
  const themeTint = attrOf(bg, "themeTint");
  const fill = bg.children
    .filter((c): c is XmlElement => c.kind === "element" && c.name.local === "background")
    .flatMap((v) =>
      v.children.filter((c): c is XmlElement => c.kind === "element" && c.name.local === "fill"),
    )[0];
  let gradient: PageGradient | undefined;
  const color2 = attrOf(fill, "color2")?.replace(/^#/, "");
  if (fill && attrOf(fill, "type") === "gradient" && color2 && HEX6.test(color2)) {
    const angle = attrOf(fill, "angle");
    const focus = attrOf(fill, "focus");
    const style = (Object.keys(GRADIENT_ANGLE) as Array<PageGradient["style"]>).find(
      (k) => String(GRADIENT_ANGLE[k] ?? "") === (angle ?? ""),
    );
    const variant = ([1, 2, 3, 4] as const).find(
      (v) => (GRADIENT_FOCUS[v] ?? "") === (focus ?? ""),
    );
    gradient = {
      color2: color2.toUpperCase(),
      style: style ?? "horizontal",
      variant: variant ?? 1,
    };
  }
  return {
    color: color.toUpperCase(),
    ...(themeColor === undefined ? {} : { themeColor }),
    ...(themeShade === undefined ? {} : { themeShade }),
    ...(themeTint === undefined ? {} : { themeTint }),
    ...(gradient ? { gradient } : {}),
  };
}

// --- Watermarks ----------------------------------------------------------------------

/** A text watermark (Design ▸ Watermark ▸ Custom Watermark ▸ Text watermark). */
export interface TextWatermark {
  readonly kind: "text";
  readonly text: string;
  readonly font: string;
  /** Font size in points; `"auto"` sizes the text to the page width. */
  readonly size: number | "auto";
  /** Hex RGB. Word's default is silver (C0C0C0). */
  readonly color: string;
  /** Draw at 50 % opacity. */
  readonly semitransparent: boolean;
  readonly layout: "diagonal" | "horizontal";
}

/** A picture watermark (Custom Watermark ▸ Picture watermark). */
export interface PictureWatermark {
  readonly kind: "picture";
  readonly bytes: Uint8Array;
  /** Override the sniffed image content type. */
  readonly contentType?: string;
  /** Displayed size in points. */
  readonly widthPt: number;
  readonly heightPt: number;
  /** Word's "Washout": lighter, lower-contrast rendering. */
  readonly washout: boolean;
}

export type Watermark = TextWatermark | PictureWatermark;

/** What {@link getWatermark} reports (a picture watermark's bytes stay in the package). */
export type WatermarkInfo =
  | TextWatermark
  | (Omit<PictureWatermark, "bytes" | "contentType"> & { readonly partName: string });

const TEXT_WATERMARK_ID = "PowerPlusWaterMarkObject";
const PICTURE_WATERMARK_ID = "WordPictureWatermark";
const WATERMARK_GALLERY = "Watermarks";
// VML picture adjustments Word writes for "Washout" (gain / black level, 16.16 fixed point).
const WASHOUT_GAIN = "19661f";
const WASHOUT_BLACK_LEVEL = "22938f";
// Average glyph advance as a fraction of the font size, to size a fixed-size
// text watermark's box; the text path stretches the text to fit the box.
const GLYPH_WIDTH_EM = 0.6;
const POINTS_PER_TWIP = 1 / 20;
const DEFAULT_TEXT_WIDTH_TWIPS = 9360;

// The VML shape types Word writes with its watermarks: `#_x0000_t136`
// (text path, ECMA-376 Part 4 §19.1.2.20) and `#_x0000_t75` (picture frame).
const TEXT_SHAPETYPE =
  '<v:shapetype id="_x0000_t136" coordsize="21600,21600" o:spt="136" adj="10800" path="m@7,l@8,m@5,21600l@6,21600e">' +
  '<v:formulas><v:f eqn="sum #0 0 10800"/><v:f eqn="prod #0 2 1"/><v:f eqn="sum 21600 0 @1"/><v:f eqn="sum 0 0 @2"/>' +
  '<v:f eqn="sum 21600 0 @3"/><v:f eqn="if @0 @3 0"/><v:f eqn="if @0 21600 @1"/><v:f eqn="if @0 0 @2"/>' +
  '<v:f eqn="if @0 @4 21600"/><v:f eqn="mid @5 @6"/><v:f eqn="mid @8 @5"/><v:f eqn="mid @7 @8"/>' +
  '<v:f eqn="mid @6 @7"/><v:f eqn="sum @6 0 @5"/></v:formulas>' +
  '<v:path textpathok="t" o:connecttype="custom" o:connectlocs="@9,0;@10,10800;@11,21600;@12,10800" o:connectangles="270,180,90,0"/>' +
  '<v:textpath on="t" fitshape="t"/><v:handles><v:h position="#0,bottomRight" xrange="6629,14971"/></v:handles>' +
  '<o:lock v:ext="edit" text="t" shapetype="t"/></v:shapetype>';

const PICTURE_SHAPETYPE =
  '<v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75" o:preferrelative="t" path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f">' +
  '<v:stroke joinstyle="miter"/><v:formulas><v:f eqn="if lineDrawn pixelLineWidth 0"/><v:f eqn="sum @0 1 0"/>' +
  '<v:f eqn="sum 0 0 @1"/><v:f eqn="prod @2 1 2"/><v:f eqn="prod @3 21600 pixelWidth"/><v:f eqn="prod @3 21600 pixelHeight"/>' +
  '<v:f eqn="sum @0 0 1"/><v:f eqn="prod @6 1 2"/><v:f eqn="prod @7 21600 pixelWidth"/><v:f eqn="sum @8 21600 0"/>' +
  '<v:f eqn="prod @7 21600 pixelHeight"/><v:f eqn="sum @10 21600 0"/></v:formulas>' +
  '<v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/><o:lock v:ext="edit" aspectratio="t"/></v:shapetype>';

const POSITION =
  "position:absolute;margin-left:0;margin-top:0;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin";
// Behind the text, as Word places watermarks.
const BEHIND_TEXT_Z = -251657216;

function pt(n: number): string {
  return `${Math.round(n * 100) / 100}pt`;
}

function textWatermarkShape(w: TextWatermark, n: number, textWidthTwips: number): string {
  if (w.text.trim() === "") throw new RangeError("The watermark text is empty.");
  const color = requireHex(w.color, "Watermark colour");
  const chars = Math.max(1, [...w.text].length);
  let width: number;
  let height: number;
  if (w.size === "auto") {
    width = textWidthTwips * POINTS_PER_TWIP;
    height = width / (chars * GLYPH_WIDTH_EM);
  } else {
    if (!(w.size > 0 && w.size <= 1638))
      throw new RangeError(`Watermark size must be 1–1638 pt, got ${w.size}.`);
    height = w.size;
    width = w.size * chars * GLYPH_WIDTH_EM;
  }
  const rotation = w.layout === "diagonal" ? ";rotation:315" : "";
  return (
    `<v:shape id="${TEXT_WATERMARK_ID}${n}" o:spid="_x0000_s${2049 + n}" type="#_x0000_t136" ` +
    `style="${POSITION};width:${pt(width)};height:${pt(height)}${rotation};z-index:${BEHIND_TEXT_Z}" ` +
    `o:allowincell="f" fillcolor="#${color}" stroked="f">` +
    (w.semitransparent ? '<v:fill opacity=".5"/>' : "") +
    `<v:textpath style="font-family:&quot;${esc(w.font)}&quot;;font-size:${w.size === "auto" ? 1 : w.size}pt" string="${esc(w.text)}"/>` +
    '<w10:wrap anchorx="margin" anchory="margin"/></v:shape>'
  );
}

function pictureWatermarkShape(w: PictureWatermark, n: number, relId: string): string {
  if (!(w.widthPt > 0 && w.heightPt > 0))
    throw new RangeError("A picture watermark needs a positive size.");
  const washout = w.washout ? ` gain="${WASHOUT_GAIN}" blacklevel="${WASHOUT_BLACK_LEVEL}"` : "";
  return (
    `<v:shape id="${PICTURE_WATERMARK_ID}${n}" o:spid="_x0000_s${2049 + n}" type="#_x0000_t75" ` +
    `style="${POSITION};width:${pt(w.widthPt)};height:${pt(w.heightPt)};z-index:${BEHIND_TEXT_Z}" o:allowincell="f">` +
    `<v:imagedata r:id="${relId}" o:title=""${washout}/>` +
    '<w10:wrap anchorx="margin" anchory="margin"/></v:shape>'
  );
}

/** Word wraps the watermark in a docPart gallery content control so it can find and replace it. */
function watermarkBlock(shapeType: string, shape: string, n: number): XmlElement {
  return fragment(
    `<w:sdt xmlns:w="${WML_NS}" xmlns:v="${VML_NS}" xmlns:o="${OFFICE_NS}" xmlns:w10="${WORD_VML_NS}" xmlns:r="${REL_NS}">` +
      `<w:sdtPr><w:id w:val="${-1000 - n}"/><w:docPartObj><w:docPartGallery w:val="${WATERMARK_GALLERY}"/><w:docPartUnique/></w:docPartObj></w:sdtPr>` +
      `<w:sdtContent><w:p><w:pPr><w:pStyle w:val="Header"/></w:pPr><w:r><w:rPr><w:noProof/></w:rPr><w:pict>${shapeType}${shape}</w:pict></w:r></w:p></w:sdtContent></w:sdt>`,
  );
}

function isWatermarkShape(el: XmlElement): boolean {
  if (el.name.uri !== VML_NS || el.name.local !== "shape") return false;
  const id = attrOf(el, "id") ?? "";
  return id.startsWith(TEXT_WATERMARK_ID) || id.startsWith(PICTURE_WATERMARK_ID);
}

function containsWatermark(el: XmlElement): boolean {
  if (isWatermarkShape(el)) return true;
  return el.children.some((c) => c.kind === "element" && containsWatermark(c));
}

function isWatermarkSdt(el: XmlElement): boolean {
  if (el.name.uri !== WML_NS || el.name.local !== "sdt") return false;
  const gallery = wmlChild(wmlChild(wmlChild(el, "sdtPr"), "docPartObj"), "docPartGallery");
  return attrOf(gallery, "val") === WATERMARK_GALLERY || containsWatermark(el);
}

/** Remove watermark content controls and loose watermark runs from a header root; returns whether any were found. */
function stripWatermarks(el: XmlElement): boolean {
  let found = false;
  const children = el.children as XmlNode[];
  for (let i = children.length - 1; i >= 0; i--) {
    const c = children[i];
    if (!c || c.kind !== "element") continue;
    const isRunWithShape = c.name.uri === WML_NS && c.name.local === "r" && containsWatermark(c);
    if (isWatermarkSdt(c) || isRunWithShape) {
      children.splice(i, 1);
      found = true;
    } else if (stripWatermarks(c)) {
      found = true;
    }
  }
  return found;
}

type HeaderType = "default" | "first" | "even";

/** Header parts (by part name) the document's sections reference, with each header type in use. */
function headerTargets(doc: Docx): string[] {
  const sections = sectionPropertiesList(doc.document);
  const types = new Set<HeaderType>(["default"]);
  if (getDocumentSetting(doc, "evenAndOddHeaders").present) types.add("even");
  if (sections.some((s) => onOffChild(s, "titlePg"))) types.add("first");
  const first = sections[0];
  if (!first) return [];
  // A section without a header reference of some type inherits the previous
  // section's (§17.10.5), so only the first section must define every type.
  const defined = new Set(
    wmlChildren(first, "headerReference").map((r) => attrOf(r, "type") ?? "default"),
  );
  for (const type of types) if (!defined.has(type)) createHeader(doc, first, type);
  const rels = partRelationships(doc.opc, doc.partName);
  const source = normalizePartName(doc.partName);
  const names = new Set<string>();
  for (const sectPr of sections) {
    for (const ref of wmlChildren(sectPr, "headerReference")) {
      const id = ref.attrs.find((a) => a.name.local === "id" && a.name.uri === REL_NS)?.value;
      const rel = rels.relationships.find((r) => r.id === id);
      if (rel) names.add(resolveInternalTarget(source, rel.target));
    }
  }
  return [...names];
}

function createHeader(doc: Docx, sectPr: XmlElement, type: HeaderType): void {
  let n = 1;
  while (hasPart(doc.opc, `/word/header${n}.xml`)) n++;
  const name = `/word/header${n}.xml`;
  addPart(doc.opc, {
    name,
    contentType: WML_CONTENT_TYPES.header,
    data: new TextEncoder().encode(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<w:hdr xmlns:w="${WML_NS}"><w:p><w:pPr><w:pStyle w:val="Header"/></w:pPr></w:p></w:hdr>`,
    ),
  });
  const rel = addRelationship(partRelationships(doc.opc, doc.partName), {
    type: WML_RELATIONSHIPS.header,
    target: `header${n}.xml`,
  });
  const ref = wmlElement("headerReference", { type });
  (ref.attrs as XmlAttr[]).push({
    name: { uri: REL_NS, local: "id", prefix: "r" },
    value: rel.id,
    isNamespaceDecl: false,
  });
  insertOrderedChild(sectPr, ref, SECT_PR_ORDER);
}

function addImagePart(doc: Docx, headerPart: string, w: PictureWatermark): string {
  const contentType = w.contentType ?? sniffImageContentType(w.bytes);
  if (!contentType) throw new RangeError("The watermark picture is not a recognized image format.");
  const ext = extensionForImageContentType(contentType);
  let n = 1;
  while (hasPart(doc.opc, `/word/media/image${n}.${ext}`)) n++;
  addPart(doc.opc, { name: `/word/media/image${n}.${ext}`, contentType, data: w.bytes });
  setContentTypeDefault(doc.opc.contentTypes, ext, contentType);
  return addRelationship(partRelationships(doc.opc, headerPart), {
    type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
    target: `media/image${n}.${ext}`,
  }).id;
}

/**
 * Put a watermark behind every page, or remove it with `undefined` — as Word
 * does it: a VML shape in each header, created where a header type in use
 * has none. Replaces any existing watermark.
 */
export function setWatermark(doc: Docx, watermark: Watermark | undefined): void {
  const targets = watermark ? headerTargets(doc) : existingHeaderParts(doc);
  const textWidth = (() => {
    const sectPr = sectionPropertiesList(doc.document, false)[0];
    const pgSz = wmlChild(sectPr, "pgSz");
    const pgMar = wmlChild(sectPr, "pgMar");
    const w = intAttrOf(pgSz, "w");
    if (w === undefined) return DEFAULT_TEXT_WIDTH_TWIPS;
    return w - (intAttrOf(pgMar, "left") ?? 1440) - (intAttrOf(pgMar, "right") ?? 1440);
  })();
  targets.forEach((name, i) => {
    const root = getRawPartRoot(doc, name);
    if (!root) return;
    stripWatermarks(root);
    if (watermark) {
      ensureNamespace(root, "v", VML_NS);
      ensureNamespace(root, "o", OFFICE_NS);
      ensureNamespace(root, "w10", WORD_VML_NS);
      ensureNamespace(root, "r", REL_NS);
      const block =
        watermark.kind === "text"
          ? watermarkBlock(TEXT_SHAPETYPE, textWatermarkShape(watermark, i, textWidth), i)
          : watermarkBlock(
              PICTURE_SHAPETYPE,
              pictureWatermarkShape(watermark, i, addImagePart(doc, name, watermark)),
              i,
            );
      (root.children as XmlNode[]).unshift(block);
    }
    markRawPartDirty(doc, name);
  });
  doc.dirty = true;
}

function existingHeaderParts(doc: Docx): string[] {
  const rels = partRelationships(doc.opc, doc.partName);
  const source = normalizePartName(doc.partName);
  return relationshipsByType(rels, WML_RELATIONSHIPS.header)
    .map((r) => resolveInternalTarget(source, r.target))
    .filter((name) => hasPart(doc.opc, name));
}

function findWatermarkShape(el: XmlElement): XmlElement | undefined {
  if (isWatermarkShape(el)) return el;
  for (const c of el.children) {
    if (c.kind !== "element") continue;
    const found = findWatermarkShape(c);
    if (found) return found;
  }
  return undefined;
}

function styleValue(style: string, key: string): string | undefined {
  const match = new RegExp(`(?:^|;)\\s*${key}\\s*:\\s*([^;]+)`).exec(style);
  return match?.[1]?.trim();
}

/** The document's watermark, read from the first header that has one. */
export function getWatermark(doc: Docx): WatermarkInfo | undefined {
  for (const name of existingHeaderParts(doc)) {
    const root = getRawPartRoot(doc, name);
    const shape = root && findWatermarkShape(root);
    if (!shape) continue;
    const style = attrOf(shape, "style") ?? "";
    const width = Number.parseFloat(styleValue(style, "width") ?? "0");
    const height = Number.parseFloat(styleValue(style, "height") ?? "0");
    const textpath = shape.children.find(
      (c): c is XmlElement => c.kind === "element" && c.name.local === "textpath",
    );
    if (textpath) {
      const tpStyle = attrOf(textpath, "style") ?? "";
      const font = (styleValue(tpStyle, "font-family") ?? "Calibri").replace(/^"|"$/g, "");
      const size = Number.parseFloat(styleValue(tpStyle, "font-size") ?? "1");
      const fill = shape.children.find(
        (c): c is XmlElement => c.kind === "element" && c.name.local === "fill",
      );
      const color = (attrOf(shape, "fillcolor") ?? "#C0C0C0").replace(/^#/, "");
      return {
        kind: "text",
        text: attrOf(textpath, "string") ?? "",
        font,
        // Word writes 1pt for "Auto" and lets the shape size the text.
        size: size <= 1 ? "auto" : size,
        // VML also allows colour names; Word's own default is "silver".
        color: HEX6.test(color) ? color.toUpperCase() : "C0C0C0",
        semitransparent: fill !== undefined && attrOf(fill, "opacity") !== undefined,
        layout: styleValue(style, "rotation") ? "diagonal" : "horizontal",
      };
    }
    const imagedata = shape.children.find(
      (c): c is XmlElement => c.kind === "element" && c.name.local === "imagedata",
    );
    const relId = imagedata?.attrs.find(
      (a) => a.name.local === "id" && a.name.uri === REL_NS,
    )?.value;
    const rel = partRelationships(doc.opc, name).relationships.find((r) => r.id === relId);
    return {
      kind: "picture",
      widthPt: width,
      heightPt: height,
      washout: attrOf(imagedata, "gain") !== undefined,
      partName: rel ? resolveInternalTarget(normalizePartName(name), rel.target) : "",
    };
  }
  return undefined;
}
