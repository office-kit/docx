/**
 * Character and paragraph formatting beyond {@link setRunFormat}: theme
 * colours, every underline style, theme fonts, run shading and borders,
 * per-side paragraph borders, tab stops, phonetic guides (ruby), enclosed
 * characters, list definitions, Word's built-in styles, and editing a
 * style's own formatting. Everything here writes ECMA-376 Part 1
 * WordprocessingML only.
 */

import {
  addPart,
  addRelationship,
  getPart,
  partRelationships,
  relationshipsByType,
} from "../internal/opc/index.js";
import {
  assertThemeColor,
  buildShading,
  childElementsOf,
  EMPTY_NUMBERING_XML,
  getElementAttr,
  makePropsElement,
  numId as readNumId,
  parseNumberingPart,
  type ShadingOptions,
  type ThemeColor,
  ucharHex,
  WML_CONTENT_TYPES,
  WML_NS,
  WML_RELATIONSHIPS,
  type WmlNumberingPart,
  type WmlParagraph,
  type WmlRun,
  type WmlRunPiece,
} from "../internal/wordprocessingml/index.js";
import { parseXml, type XmlAttr, type XmlElement } from "../internal/xml/index.js";
import { addStyle, type Docx, numberingPart, stylesPart } from "./docx.js";

export {
  SHADING_PATTERNS,
  type ShadingOptions,
  type ShadingPattern,
  THEME_COLORS,
  type ThemeColor,
} from "../internal/wordprocessingml/index.js";

// ---------------------------------------------------------------------------
// Small XML helpers

function wAttr(local: string, value: string): XmlAttr {
  return { name: { uri: WML_NS, local, prefix: "w" }, value, isNamespaceDecl: false };
}

function wEl(local: string, attrs: XmlAttr[] = [], children: XmlElement[] = []): XmlElement {
  return {
    kind: "element",
    name: { uri: WML_NS, local, prefix: "w" },
    attrs,
    children,
    xmlSpace: "default",
    selfClosing: children.length === 0,
  };
}

function isW(el: XmlElement, local: string): boolean {
  return el.name.uri === WML_NS && el.name.local === local;
}

function findChild(container: XmlElement | undefined, local: string): XmlElement | undefined {
  return container && childElementsOf(container).find((c) => isW(c, local));
}

function removeChildren(container: XmlElement, local: string): void {
  const children = container.children as XmlElement[];
  for (let i = children.length - 1; i >= 0; i--) {
    const c = children[i];
    if (c && c.kind === "element" && isW(c, local)) children.splice(i, 1);
  }
}

/** Replace (or, with `undefined`, remove) a container's `local` child. */
function replaceChild(container: XmlElement, local: string, el: XmlElement | undefined): void {
  removeChildren(container, local);
  if (el) (container.children as XmlElement[]).push(el);
}

function ensureRPr(run: WmlRun): XmlElement {
  run.rPr ??= makePropsElement("rPr");
  return run.rPr;
}

function ensurePPr(p: WmlParagraph): XmlElement {
  p.pPr ??= makePropsElement("pPr");
  return p.pPr;
}

// ST_HexColor: "auto" or six hex digits.
const HEX_COLOR = /^(?:auto|[0-9A-Fa-f]{6})$/;

function assertHex(value: string, what: string): void {
  if (typeof value !== "string" || !HEX_COLOR.test(value)) {
    throw new RangeError(
      `${what} must be six hex digits (no "#") or "auto", got ${JSON.stringify(value)}.`,
    );
  }
}

function assertInt(value: number, min: number, max: number, what: string): void {
  if (!Number.isInteger(value) || value < min || value > max) {
    throw new RangeError(`${what} must be an integer in ${min}–${max}, got ${String(value)}.`);
  }
}

// ---------------------------------------------------------------------------
// Colours

/**
 * A colour as WordprocessingML stores it: the RGB value every consumer can
 * use, plus, for a theme colour, which theme slot it came from and the tint
 * or shade applied (0–255; `themeTint` 0x33 is Word's "Lighter 80%").
 * Consumers that know the theme recompute the colour from the slot, so a
 * theme change recolours the text (§17.3.2.6).
 */
export interface ColorValue {
  /** Six hex digits without `#`, or `"auto"`. */
  readonly rgb: string;
  readonly themeColor?: ThemeColor;
  readonly themeTint?: number;
  readonly themeShade?: number;
}

/** `w:val` / `w:color` plus the theme attributes, validated. */
function colorAttrs(color: ColorValue, valName: "val" | "color"): XmlAttr[] {
  assertHex(color.rgb, "Color");
  const attrs = [wAttr(valName, color.rgb)];
  if (color.themeColor !== undefined) {
    assertThemeColor(color.themeColor, "themeColor");
    attrs.push(wAttr("themeColor", color.themeColor));
  }
  if (color.themeTint !== undefined)
    attrs.push(wAttr("themeTint", ucharHex(color.themeTint, "themeTint")));
  if (color.themeShade !== undefined)
    attrs.push(wAttr("themeShade", ucharHex(color.themeShade, "themeShade")));
  return attrs;
}

/**
 * Set a run's text colour (`<w:color>`), theme colours included; `undefined`
 * removes it so the style's colour shows. Throws a `RangeError` for a value
 * outside the schema, before changing anything.
 */
export function setRunColor(run: WmlRun, color: ColorValue | undefined): void {
  const el = color && wEl("color", colorAttrs(color, "val"));
  replaceChild(ensureRPr(run), "color", el);
}

// ---------------------------------------------------------------------------
// Underline

/** Every `w:u/@w:val`: ST_Underline (§17.18.99), in Word's menu order. */
export const UNDERLINE_STYLES = [
  "single",
  "words",
  "double",
  "thick",
  "dotted",
  "dottedHeavy",
  "dash",
  "dashedHeavy",
  "dashLong",
  "dashLongHeavy",
  "dotDash",
  "dashDotHeavy",
  "dotDotDash",
  "dashDotDotHeavy",
  "wave",
  "wavyHeavy",
  "wavyDouble",
  "none",
] as const;

export type UnderlineStyle = (typeof UNDERLINE_STYLES)[number];

const UNDERLINE_SET: ReadonlySet<string> = new Set(UNDERLINE_STYLES);

/** A run's underline: the line style and, optionally, its own colour. */
export interface UnderlineValue {
  readonly style: UnderlineStyle;
  /** Without a colour the line takes the text colour. */
  readonly color?: ColorValue;
}

/**
 * Set a run's underline (`<w:u>`); `undefined` removes it so the style's
 * underline shows, and `style: "none"` writes an explicit no-underline that
 * overrides the style. Throws a `RangeError` for a value outside the schema.
 */
export function setRunUnderline(run: WmlRun, underline: UnderlineValue | undefined): void {
  let el: XmlElement | undefined;
  if (underline) {
    if (!UNDERLINE_SET.has(underline.style)) {
      throw new RangeError(
        `Underline style must be an ST_Underline value, got ${JSON.stringify(underline.style)}.`,
      );
    }
    el = wEl("u", [
      wAttr("val", underline.style),
      ...(underline.color ? colorAttrs(underline.color, "color") : []),
    ]);
  }
  replaceChild(ensureRPr(run), "u", el);
}

// ---------------------------------------------------------------------------
// Fonts

/** Which characters a font applies to (`w:rFonts`, §17.3.2.26). */
export type FontScript = "latin" | "eastAsia" | "complex";

/** A font by name, or the theme's headings (`major`) / body (`minor`) font. */
export type FontChoice = { readonly name: string } | { readonly theme: "major" | "minor" };

// Latin text uses two slots: ASCII and high ANSI.
const SCRIPT_SLOTS: Readonly<
  Record<FontScript, { names: string[]; themes: string[]; themeValue: string }>
> = {
  latin: { names: ["ascii", "hAnsi"], themes: ["asciiTheme", "hAnsiTheme"], themeValue: "HAnsi" },
  eastAsia: { names: ["eastAsia"], themes: ["eastAsiaTheme"], themeValue: "EastAsia" },
  // The schema spells this attribute in lower case.
  complex: { names: ["cs"], themes: ["cstheme"], themeValue: "Bidi" },
};

/**
 * Set the font a run uses for one script; `undefined` removes the run's own
 * choice so the style's font shows. A theme attribute overrides the named
 * font on the same element (§17.3.2.26), so choosing one clears the other.
 * Throws a `RangeError` for an empty font name.
 */
export function setRunFont(run: WmlRun, script: FontScript, font: FontChoice | undefined): void {
  if (font && "name" in font && (typeof font.name !== "string" || font.name.trim() === "")) {
    throw new RangeError("Font name must be a non-empty string.");
  }
  const rPr = ensureRPr(run);
  const existing = findChild(rPr, "rFonts");
  const slots = SCRIPT_SLOTS[script];
  const drop = new Set([...slots.names, ...slots.themes]);
  const attrs = (existing?.attrs ?? []).filter((a) => !drop.has(a.name.local));
  if (font && "name" in font) {
    for (const slot of slots.names) attrs.push(wAttr(slot, font.name));
  } else if (font) {
    for (const slot of slots.themes) attrs.push(wAttr(slot, `${font.theme}${slots.themeValue}`));
  }
  replaceChild(rPr, "rFonts", attrs.length > 0 ? wEl("rFonts", attrs) : undefined);
}

// ---------------------------------------------------------------------------
// Run shading and borders

/**
 * Shade a run (`<w:shd>` in `<w:rPr>`, Word's Character Shading and the
 * Shading tab of Borders and Shading applied to text); `undefined` removes
 * it. Throws a `RangeError` for a value outside the schema.
 */
export function setRunShading(run: WmlRun, shading: ShadingOptions | undefined): void {
  if (shading) {
    assertHex(shading.fill ?? "auto", "Shading fill");
    assertHex(shading.color ?? "auto", "Shading color");
  }
  const el = shading && buildShading(shading);
  replaceChild(ensureRPr(run), "shd", el);
}

/**
 * The line styles of ST_Border (§17.18.2) that apply to text and paragraphs;
 * the art borders (apples, stars, …) are for page borders only.
 */
export const BORDER_LINE_STYLES = [
  "single",
  "dotted",
  "dashSmallGap",
  "dashed",
  "dotDash",
  "dotDotDash",
  "double",
  "triple",
  "thinThickSmallGap",
  "thickThinSmallGap",
  "thinThickThinSmallGap",
  "thinThickMediumGap",
  "thickThinMediumGap",
  "thinThickThinMediumGap",
  "thinThickLargeGap",
  "thickThinLargeGap",
  "thinThickThinLargeGap",
  "wave",
  "doubleWave",
  "dashDotStroked",
  "threeDEmboss",
  "threeDEngrave",
  "outset",
  "inset",
  "thick",
  "none",
  "nil",
] as const;

export type BorderLineStyle = (typeof BORDER_LINE_STYLES)[number];

const BORDER_STYLE_SET: ReadonlySet<string> = new Set(BORDER_LINE_STYLES);
// ST_EighthPointMeasure as Word accepts it for a line border: ¼ pt to 12 pt.
const MAX_BORDER_EIGHTHS = 96;
// ST_PointMeasure for w:space; Word keeps 0–31 pt.
const MAX_BORDER_SPACE = 31;

/** One border line (`CT_Border`, §17.18.2 / §17.3.4). */
export interface BorderLine {
  readonly style: BorderLineStyle;
  /** Width in eighths of a point (`w:sz`); 4 (½ pt) when omitted. */
  readonly sizeEighths?: number;
  /** Gap between the border and the text in points (`w:space`); 0 when omitted. */
  readonly spacePt?: number;
  /** `auto` when omitted. */
  readonly color?: ColorValue;
  /** `w:shadow`: draw the border with a shadow. */
  readonly shadow?: boolean;
}

const DEFAULT_BORDER_EIGHTHS = 4;

function borderElement(local: string, line: BorderLine): XmlElement {
  if (!BORDER_STYLE_SET.has(line.style)) {
    throw new RangeError(
      `Border style must be an ST_Border line style, got ${JSON.stringify(line.style)}.`,
    );
  }
  const size = line.sizeEighths ?? DEFAULT_BORDER_EIGHTHS;
  assertInt(size, 0, MAX_BORDER_EIGHTHS, "Border width (eighths of a point)");
  const space = line.spacePt ?? 0;
  assertInt(space, 0, MAX_BORDER_SPACE, "Border spacing (points)");
  const attrs = [
    wAttr("val", line.style),
    wAttr("sz", String(size)),
    wAttr("space", String(space)),
    ...colorAttrs(line.color ?? { rgb: "auto" }, "color"),
  ];
  if (line.shadow) attrs.push(wAttr("shadow", "1"));
  return wEl(local, attrs);
}

/**
 * Put a border around a run (`<w:bdr>`, Word's Character Border); runs next
 * to each other with the same border share one box. `undefined` removes it.
 * Throws a `RangeError` for a value outside the schema.
 */
export function setRunBorder(run: WmlRun, line: BorderLine | undefined): void {
  const el = line && borderElement("bdr", line);
  replaceChild(ensureRPr(run), "bdr", el);
}

/** A side of a paragraph's border box; `between` draws between paragraphs that share it. */
export type ParagraphBorderSide = "top" | "left" | "bottom" | "right" | "between" | "bar";

/**
 * Set one side of a paragraph's borders (`<w:pBdr>`), keeping the others;
 * `undefined` removes that side, and the `<w:pBdr>` goes when no side is
 * left. Throws a `RangeError` for a value outside the schema.
 */
export function setParagraphBorder(
  p: WmlParagraph,
  side: ParagraphBorderSide,
  line: BorderLine | undefined,
): void {
  const el = line && borderElement(side, line);
  const pPr = ensurePPr(p);
  const pBdr = findChild(pPr, "pBdr") ?? wEl("pBdr");
  replaceChild(pBdr, side, el);
  replaceChild(pPr, "pBdr", childElementsOf(pBdr).length > 0 ? pBdr : undefined);
}

// ---------------------------------------------------------------------------
// Tab stops

/** ST_TabJc values Word offers, plus `clear` (removes an inherited stop). */
export type TabAlignment = "left" | "center" | "right" | "decimal" | "bar" | "clear";
/** ST_TabTlc (§17.18.85). */
export type TabLeader = "none" | "dot" | "hyphen" | "underscore" | "heavy" | "middleDot";

const TAB_ALIGNMENTS: ReadonlySet<string> = new Set([
  "left",
  "center",
  "right",
  "decimal",
  "bar",
  "clear",
]);
const TAB_LEADERS: ReadonlySet<string> = new Set([
  "none",
  "dot",
  "hyphen",
  "underscore",
  "heavy",
  "middleDot",
]);
// ST_SignedTwipsMeasure; Word's tab ruler stops at 22 in either way.
const MAX_TAB_TWIPS = 31680;

/** A custom tab stop (`<w:tab>` in `<w:tabs>`, §17.3.1.37). */
export interface TabStop {
  /** Twips from the paragraph's leading indent. */
  readonly position: number;
  readonly alignment: TabAlignment;
  readonly leader?: TabLeader;
}

/**
 * Replace a paragraph's custom tab stops; an empty list removes `<w:tabs>`.
 * Stops are written in position order. Throws a `RangeError` for a value
 * outside the schema.
 */
export function setParagraphTabs(p: WmlParagraph, tabs: readonly TabStop[]): void {
  for (const tab of tabs) {
    assertInt(tab.position, -MAX_TAB_TWIPS, MAX_TAB_TWIPS, "Tab position (twips)");
    if (!TAB_ALIGNMENTS.has(tab.alignment))
      throw new RangeError(`Tab alignment must be one of ${[...TAB_ALIGNMENTS].join(", ")}.`);
    if (tab.leader !== undefined && !TAB_LEADERS.has(tab.leader))
      throw new RangeError(`Tab leader must be one of ${[...TAB_LEADERS].join(", ")}.`);
  }
  const stops = tabs
    .toSorted((a, b) => a.position - b.position)
    .map((tab) =>
      wEl("tab", [
        wAttr("val", tab.alignment),
        ...(tab.leader && tab.leader !== "none" ? [wAttr("leader", tab.leader)] : []),
        wAttr("pos", String(tab.position)),
      ]),
    );
  replaceChild(ensurePPr(p), "tabs", stops.length > 0 ? wEl("tabs", [], stops) : undefined);
}

/** A paragraph's own custom tab stops (not the style's), in document order. */
export function getParagraphTabs(p: WmlParagraph): TabStop[] {
  const tabs = findChild(p.pPr, "tabs");
  if (!tabs) return [];
  return childElementsOf(tabs).flatMap((tab) => {
    const position = Number(getElementAttr(tab, "pos"));
    const val = getElementAttr(tab, "val") ?? "left";
    // `start` / `end` are the bidi-aware spellings of left / right.
    const alignment = val === "start" ? "left" : val === "end" ? "right" : val;
    const leader = getElementAttr(tab, "leader");
    if (!Number.isInteger(position) || !TAB_ALIGNMENTS.has(alignment)) return [];
    return [
      {
        position,
        alignment: alignment as TabAlignment,
        ...(leader !== undefined && TAB_LEADERS.has(leader) ? { leader: leader as TabLeader } : {}),
      },
    ];
  });
}

// ---------------------------------------------------------------------------
// Phonetic guide (ruby)

/** ST_RubyAlign (§17.18.75). */
export const RUBY_ALIGNMENTS = [
  "center",
  "distributeLetter",
  "distributeSpace",
  "left",
  "right",
  "rightVertical",
] as const;

export type RubyAlignment = (typeof RUBY_ALIGNMENTS)[number];

const RUBY_ALIGN_SET: ReadonlySet<string> = new Set(RUBY_ALIGNMENTS);

/** A phonetic guide's layout (`<w:rubyPr>`, §17.3.3.24). */
export interface RubyOptions {
  readonly alignment: RubyAlignment;
  /** Size of the guide text in half-points (`w:hps`). */
  readonly rubySizeHalfPoints: number;
  /** Distance from the base text's baseline to the guide's, in half-points (`w:hpsRaise`). */
  readonly raiseHalfPoints: number;
  /** Size of the base text in half-points (`w:hpsBaseText`). */
  readonly baseSizeHalfPoints: number;
  /** Language of the guide (`w:lid`), e.g. `ja-JP`. */
  readonly language: string;
  /** Font of the guide text; the base text's font when omitted. */
  readonly rubyFont?: string;
}

/** A phonetic guide read back from a run. */
export interface RubyInfo extends RubyOptions {
  readonly base: string;
  readonly ruby: string;
}

const MAX_HALF_POINTS = 3276;

const XML_NS = "http://www.w3.org/XML/1998/namespace";
// Leading / trailing whitespace is dropped by XML consumers unless preserved.
const EDGE_SPACE = /^\s|\s$/;

function textRun(text: string, rPr: XmlElement | undefined): XmlElement {
  const preserve = EDGE_SPACE.test(text);
  const t: XmlElement = {
    kind: "element",
    name: { uri: WML_NS, local: "t", prefix: "w" },
    attrs: preserve
      ? [
          {
            name: { uri: XML_NS, local: "space", prefix: "xml" },
            value: "preserve",
            isNamespaceDecl: false,
          },
        ]
      : [],
    children: [{ kind: "text", value: text }],
    xmlSpace: preserve ? "preserve" : "default",
    selfClosing: false,
  };
  return wEl("r", [], rPr ? [rPr, t] : [t]);
}

function cloneElement(el: XmlElement): XmlElement {
  return {
    ...el,
    attrs: el.attrs.map((a) => ({ ...a })),
    children: el.children.map((c) => (c.kind === "element" ? cloneElement(c) : { ...c })),
  };
}

/** `rPr` with its `<w:sz>` / `<w:szCs>` (and, optionally, the font) replaced. */
function withSize(rPr: XmlElement | undefined, halfPoints: number, font?: string): XmlElement {
  const out = rPr ? cloneElement(rPr) : makePropsElement("rPr");
  if (font !== undefined) {
    replaceChild(
      out,
      "rFonts",
      wEl("rFonts", [wAttr("ascii", font), wAttr("hAnsi", font), wAttr("eastAsia", font)]),
    );
  }
  replaceChild(out, "sz", wEl("sz", [wAttr("val", String(halfPoints))]));
  replaceChild(out, "szCs", wEl("szCs", [wAttr("val", String(halfPoints))]));
  return out;
}

/**
 * Build a run that shows `ruby` as a phonetic guide over `base` (`<w:ruby>`,
 * Word's Phonetic Guide). `rPr` is the base text's formatting; the guide gets
 * the same formatting at its own size. Throws a `RangeError` for empty text or
 * a value outside the schema.
 */
export function buildRubyRun(
  base: string,
  ruby: string,
  options: RubyOptions,
  rPr?: XmlElement,
): WmlRun {
  if (base === "" || ruby === "")
    throw new RangeError("Ruby base and guide text must not be empty.");
  if (!RUBY_ALIGN_SET.has(options.alignment))
    throw new RangeError(`Ruby alignment must be one of ${RUBY_ALIGNMENTS.join(", ")}.`);
  assertInt(options.rubySizeHalfPoints, 1, MAX_HALF_POINTS, "Ruby size (half-points)");
  assertInt(options.baseSizeHalfPoints, 1, MAX_HALF_POINTS, "Ruby base size (half-points)");
  assertInt(
    options.raiseHalfPoints,
    -MAX_HALF_POINTS,
    MAX_HALF_POINTS,
    "Ruby offset (half-points)",
  );
  if (options.language.trim() === "") throw new RangeError("Ruby language must not be empty.");
  const rubyPr = wEl(
    "rubyPr",
    [],
    [
      wEl("rubyAlign", [wAttr("val", options.alignment)]),
      wEl("hps", [wAttr("val", String(options.rubySizeHalfPoints))]),
      wEl("hpsRaise", [wAttr("val", String(options.raiseHalfPoints))]),
      wEl("hpsBaseText", [wAttr("val", String(options.baseSizeHalfPoints))]),
      wEl("lid", [wAttr("val", options.language)]),
    ],
  );
  const rubyEl = wEl(
    "ruby",
    [],
    [
      rubyPr,
      wEl("rt", [], [textRun(ruby, withSize(rPr, options.rubySizeHalfPoints, options.rubyFont))]),
      wEl("rubyBase", [], [textRun(base, rPr && cloneElement(rPr))]),
    ],
  );
  const piece: WmlRunPiece = { kind: "raw", node: rubyEl };
  return { kind: "run", ...(rPr ? { rPr: cloneElement(rPr) } : {}), pieces: [piece], extras: [] };
}

function elementText(el: XmlElement | undefined): string {
  if (!el) return "";
  if (isW(el, "t")) return el.children.map((c) => (c.kind === "text" ? c.value : "")).join("");
  return childElementsOf(el).map(elementText).join("");
}

/** The phonetic guide a run holds, or `undefined` when it holds none. */
export function readRuby(run: WmlRun): RubyInfo | undefined {
  const piece = run.pieces.find((p) => p.kind === "raw" && isW(p.node, "ruby"));
  if (!piece || piece.kind !== "raw") return undefined;
  const rubyPr = findChild(piece.node, "rubyPr");
  const val = (local: string): string | undefined => {
    const el = findChild(rubyPr, local);
    return el && getElementAttr(el, "val");
  };
  const num = (local: string, fallback: number): number => {
    const n = Number(val(local));
    return Number.isInteger(n) ? n : fallback;
  };
  const alignment = val("rubyAlign") ?? "center";
  const rtRun = findChild(findChild(piece.node, "rt"), "r");
  const rtFonts = findChild(findChild(rtRun, "rPr"), "rFonts");
  const rubyFont = rtFonts && getElementAttr(rtFonts, "eastAsia");
  return {
    base: elementText(findChild(piece.node, "rubyBase")),
    ruby: elementText(findChild(piece.node, "rt")),
    alignment: RUBY_ALIGN_SET.has(alignment) ? (alignment as RubyAlignment) : "center",
    rubySizeHalfPoints: num("hps", 10),
    raiseHalfPoints: num("hpsRaise", 20),
    baseSizeHalfPoints: num("hpsBaseText", 21),
    language: val("lid") ?? "ja-JP",
    ...(rubyFont !== undefined ? { rubyFont } : {}),
  };
}

// ---------------------------------------------------------------------------
// Enclosed characters

/** The symbols Word's Enclose Characters offers. */
export const ENCLOSURES = { circle: "○", square: "□", triangle: "△", diamond: "◇" } as const;
export type Enclosure = keyof typeof ENCLOSURES;

/** Enclose Characters options: which shape, and which part is resized to fit. */
export interface EncloseOptions {
  readonly enclosure: Enclosure;
  /** `shrinkText` keeps the symbol at the text size; `enlargeSymbol` keeps the text. */
  readonly style: "shrinkText" | "enlargeSymbol";
  /** The surrounding text size in half-points. */
  readonly sizeHalfPoints: number;
}

// The proportions Word uses for Enclose Characters (observed in documents
// Word for Mac saves): the shrunk text is ~70 % of the size and raised by
// about a tenth of it; the enlarged symbol is ~140 % and lowered by ~a fifth.
const SHRINK_RATIO = 0.7;
const SHRINK_RAISE_RATIO = 0.1;
const ENLARGE_RATIO = 1.4;
const ENLARGE_DROP_RATIO = 0.2;

function fieldChar(type: "begin" | "end"): WmlRunPiece {
  return { kind: "fieldChar", charType: type, raw: wEl("fldChar", [wAttr("fldCharType", type)]) };
}

function instrRun(text: string, rPr?: XmlElement): WmlRun {
  return {
    kind: "run",
    ...(rPr ? { rPr } : {}),
    pieces: [{ kind: "instrText", value: text, preserveSpace: EDGE_SPACE.test(text) }],
    extras: [],
  };
}

function sizedRPr(base: XmlElement | undefined, halfPoints: number, position: number): XmlElement {
  const rPr = withSize(base, halfPoints);
  replaceChild(rPr, "position", wEl("position", [wAttr("val", String(position))]));
  return rPr;
}

/**
 * Build the runs of Word's Enclose Characters: an `EQ \o\ac(symbol,text)`
 * field (§17.16.5.20), which overstrikes the symbol and the text, centred.
 * `rPr` is the text's formatting. Throws a `RangeError` for empty text or a
 * size outside the schema.
 */
export function buildEnclosedCharacterRuns(
  text: string,
  options: EncloseOptions,
  rPr?: XmlElement,
): WmlRun[] {
  if (text === "") throw new RangeError("The text to enclose must not be empty.");
  if (!(options.enclosure in ENCLOSURES)) throw new RangeError("Unknown enclosure.");
  assertInt(options.sizeHalfPoints, 1, MAX_HALF_POINTS, "Enclosure size (half-points)");
  const size = options.sizeHalfPoints;
  const symbol = ENCLOSURES[options.enclosure];
  const symbolRPr =
    options.style === "enlargeSymbol"
      ? sizedRPr(rPr, Math.round(size * ENLARGE_RATIO), -Math.round(size * ENLARGE_DROP_RATIO))
      : sizedRPr(rPr, size, 0);
  const textRPr =
    options.style === "shrinkText"
      ? sizedRPr(rPr, Math.round(size * SHRINK_RATIO), Math.round(size * SHRINK_RAISE_RATIO))
      : sizedRPr(rPr, size, 0);
  const begin: WmlRun = { kind: "run", pieces: [fieldChar("begin")], extras: [] };
  const end: WmlRun = { kind: "run", pieces: [fieldChar("end")], extras: [] };
  return [
    begin,
    instrRun(" eq \\o\\ac("),
    instrRun(symbol, symbolRPr),
    instrRun(","),
    instrRun(text, textRPr),
    instrRun(")"),
    end,
  ];
}

// ---------------------------------------------------------------------------
// List definitions

/** Every ST_NumberFormat value (§17.18.59). */
export const NUMBER_FORMATS = [
  "decimal",
  "upperRoman",
  "lowerRoman",
  "upperLetter",
  "lowerLetter",
  "ordinal",
  "cardinalText",
  "ordinalText",
  "hex",
  "chicago",
  "ideographDigital",
  "japaneseCounting",
  "aiueo",
  "iroha",
  "decimalFullWidth",
  "decimalHalfWidth",
  "japaneseLegal",
  "japaneseDigitalTenThousand",
  "decimalEnclosedCircle",
  "decimalFullWidth2",
  "aiueoFullWidth",
  "irohaFullWidth",
  "decimalZero",
  "bullet",
  "ganada",
  "chosung",
  "decimalEnclosedFullstop",
  "decimalEnclosedParen",
  "decimalEnclosedCircleChinese",
  "ideographEnclosedCircle",
  "ideographTraditional",
  "ideographZodiac",
  "ideographZodiacTraditional",
  "taiwaneseCounting",
  "ideographLegalTraditional",
  "taiwaneseCountingThousand",
  "taiwaneseDigital",
  "chineseCounting",
  "chineseLegalSimplified",
  "chineseCountingThousand",
  "koreanDigital",
  "koreanCounting",
  "koreanLegal",
  "koreanDigital2",
  "vietnameseCounting",
  "russianLower",
  "russianUpper",
  "none",
  "numberInDash",
  "hebrew1",
  "hebrew2",
  "arabicAlpha",
  "arabicAbjad",
  "hindiVowels",
  "hindiConsonants",
  "hindiNumbers",
  "hindiCounting",
  "thaiLetters",
  "thaiNumbers",
  "thaiCounting",
] as const;

export type ListNumberFormat = (typeof NUMBER_FORMATS)[number];

const NUMBER_FORMAT_SET: ReadonlySet<string> = new Set(NUMBER_FORMATS);
const MAX_LIST_LEVELS = 9;

/** One level of a list definition (`<w:lvl>`, §17.9.6). */
export interface ListLevel {
  readonly format: ListNumberFormat;
  /** The label; `%1` … `%9` stand for the levels' numbers. A bullet's label is its symbol. */
  readonly text: string;
  /** First number; 1 when omitted. */
  readonly start?: number;
  readonly alignment?: "left" | "center" | "right";
  /** Left indent of the text in twips. */
  readonly indentLeft: number;
  /** How far the label hangs left of the text, in twips. */
  readonly hanging: number;
  /** Font of the label (the bullet font, e.g. `Symbol`). */
  readonly font?: string;
  /** Label colour; the paragraph's text colour when omitted. */
  readonly color?: string;
  /** What follows the label; a tab when omitted. */
  readonly suffix?: "tab" | "space" | "nothing";
  /** Show earlier levels' numbers as Arabic numerals (`w:isLgl`). */
  readonly legal?: boolean;
  /** The paragraph style this level is linked to (`w:pStyle`), e.g. `Heading1`. */
  readonly style?: string;
}

/** Ensure `word/numbering.xml` exists; mirrors what the list builders in docx.ts do. */
function ensureNumbering(doc: Docx): WmlNumberingPart {
  const existing = numberingPart(doc);
  if (existing) return existing;
  const name = "/word/numbering.xml";
  addPart(doc.opc, {
    name,
    contentType: WML_CONTENT_TYPES.numbering,
    data: new TextEncoder().encode(EMPTY_NUMBERING_XML),
  });
  const rels = partRelationships(doc.opc, doc.partName);
  if (relationshipsByType(rels, WML_RELATIONSHIPS.numbering).length === 0) {
    addRelationship(rels, { type: WML_RELATIONSHIPS.numbering, target: "numbering.xml" });
  }
  const xml = new TextDecoder("utf-8").decode(getPart(doc.opc, name)?.data ?? new Uint8Array());
  doc.numberingCache = parseNumberingPart(parseXml(xml));
  return doc.numberingCache;
}

function nextFreeId(ids: Iterable<number | undefined>, from: number): number {
  const used = new Set(ids);
  let id = from;
  while (used.has(id)) id++;
  return id;
}

function levelElement(level: ListLevel, ilvl: number): XmlElement {
  if (!NUMBER_FORMAT_SET.has(level.format))
    throw new RangeError(
      `List format must be an ST_NumberFormat value, got ${JSON.stringify(level.format)}.`,
    );
  assertInt(level.start ?? 1, 0, Number.MAX_SAFE_INTEGER, "List start");
  assertInt(level.indentLeft, -MAX_TAB_TWIPS, MAX_TAB_TWIPS, "List indent (twips)");
  assertInt(level.hanging, 0, MAX_TAB_TWIPS, "List hanging indent (twips)");
  if (level.color !== undefined) assertHex(level.color, "List label color");
  const children = [
    wEl("start", [wAttr("val", String(level.start ?? 1))]),
    wEl("numFmt", [wAttr("val", level.format)]),
  ];
  if (level.style) children.push(wEl("pStyle", [wAttr("val", level.style)]));
  if (level.legal) children.push(wEl("isLgl"));
  if (level.suffix && level.suffix !== "tab")
    children.push(wEl("suff", [wAttr("val", level.suffix)]));
  children.push(wEl("lvlText", [wAttr("val", level.text)]));
  children.push(wEl("lvlJc", [wAttr("val", level.alignment ?? "left")]));
  children.push(
    wEl(
      "pPr",
      [],
      [
        wEl("ind", [
          wAttr("left", String(level.indentLeft)),
          wAttr("hanging", String(level.hanging)),
        ]),
      ],
    ),
  );
  const rPr: XmlElement[] = [];
  if (level.font) {
    rPr.push(
      wEl("rFonts", [
        wAttr("ascii", level.font),
        wAttr("hAnsi", level.font),
        wAttr("hint", "default"),
      ]),
    );
  }
  if (level.color) rPr.push(wEl("color", [wAttr("val", level.color)]));
  if (rPr.length > 0) children.push(wEl("rPr", [], rPr));
  return wEl("lvl", [wAttr("ilvl", String(ilvl))], children);
}

/**
 * Define a new list (an `<w:abstractNum>` with 1–9 levels, and a `<w:num>`
 * instance of it) and return its `numId` for {@link applyListToParagraph}.
 * Throws a `RangeError` for a value outside the schema.
 */
export function addListDefinition(doc: Docx, levels: readonly ListLevel[]): number {
  if (levels.length === 0 || levels.length > MAX_LIST_LEVELS) {
    throw new RangeError(`A list has 1–${MAX_LIST_LEVELS} levels, got ${levels.length}.`);
  }
  const lvlEls = levels.map(levelElement);
  const part = ensureNumbering(doc);
  const abstractId = nextFreeId(
    part.abstractNums.map((a) => Number(getElementAttr(a, "abstractNumId"))),
    0,
  );
  const numIdValue = nextFreeId(part.nums.map(readNumId), 1);
  part.abstractNums.push(
    wEl(
      "abstractNum",
      [wAttr("abstractNumId", String(abstractId))],
      [
        wEl("multiLevelType", [wAttr("val", levels.length > 1 ? "multilevel" : "singleLevel")]),
        ...lvlEls,
      ],
    ),
  );
  part.nums.push(
    wEl(
      "num",
      [wAttr("numId", String(numIdValue))],
      [wEl("abstractNumId", [wAttr("val", String(abstractId))])],
    ),
  );
  doc.numberingDirty = true;
  doc.dirty = true;
  return numIdValue;
}

/**
 * A new list instance (`<w:num>`) of the same definition as `numId` whose
 * level `ilvl` starts at `start` (`<w:lvlOverride><w:startOverride>`): Word's
 * Restart Numbering and Set Numbering Value. Returns the new `numId`, or
 * `undefined` when `numId` names no list.
 */
export function restartList(
  doc: Docx,
  numId: number,
  ilvl: number,
  start: number,
): number | undefined {
  assertInt(ilvl, 0, MAX_LIST_LEVELS - 1, "List level");
  assertInt(start, 0, Number.MAX_SAFE_INTEGER, "List start");
  const part = numberingPart(doc);
  const source = part?.nums.find((n) => readNumId(n) === numId);
  const abstractRef = findChild(source, "abstractNumId");
  const abstractId = abstractRef && getElementAttr(abstractRef, "val");
  if (!part || abstractId === undefined) return undefined;
  const numIdValue = nextFreeId(part.nums.map(readNumId), 1);
  part.nums.push(
    wEl(
      "num",
      [wAttr("numId", String(numIdValue))],
      [
        wEl("abstractNumId", [wAttr("val", abstractId)]),
        wEl(
          "lvlOverride",
          [wAttr("ilvl", String(ilvl))],
          [wEl("startOverride", [wAttr("val", String(start))])],
        ),
      ],
    ),
  );
  doc.numberingDirty = true;
  doc.dirty = true;
  return numIdValue;
}

// ---------------------------------------------------------------------------
// Built-in styles

const THEME_HEADINGS =
  '<w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:cstheme="majorBidi"/>';
const THEME_HEADINGS_EAST_ASIA =
  '<w:rFonts w:eastAsiaTheme="majorEastAsia" w:cstheme="majorBidi"/>';
const ACCENT1_DARK = '<w:color w:val="0F4761" w:themeColor="accent1" w:themeShade="BF"/>';
const TEXT1_LIGHT_35 = '<w:color w:val="595959" w:themeColor="text1" w:themeTint="A6"/>';
const TEXT1_LIGHT_25 = '<w:color w:val="404040" w:themeColor="text1" w:themeTint="BF"/>';
const TEXT1_LIGHT_15 = '<w:color w:val="272727" w:themeColor="text1" w:themeTint="D8"/>';
const SIZE = (hp: number): string => `<w:sz w:val="${hp}"/><w:szCs w:val="${hp}"/>`;
const ITALIC = "<w:i/><w:iCs/>";
const BOLD = "<w:b/><w:bCs/>";

interface BuiltinStyle {
  readonly type: "paragraph" | "character";
  readonly name: string;
  readonly uiPriority: number;
  /** Shown in the gallery and the Styles pane without being used first. */
  readonly visible: boolean;
  readonly pPr?: string;
  readonly rPr?: string;
  /** Paragraph styles Word pairs with a linked character style. */
  readonly linked?: boolean;
  readonly basedOn?: string;
}

function heading(level: number, pPr: string, rPr: string): BuiltinStyle {
  return {
    type: "paragraph",
    name: `heading ${level}`,
    uiPriority: 9,
    visible: level <= 2,
    linked: true,
    pPr: `<w:keepNext/><w:keepLines/>${pPr}<w:outlineLvl w:val="${level - 1}"/>`,
    rPr,
  };
}

/**
 * Word's built-in styles with the definitions Word (Microsoft 365, the Office
 * 2023 theme) writes when one is first used. Colours carry theme references
 * so they follow the document's theme.
 */
const BUILTIN_STYLES: Readonly<Record<string, BuiltinStyle>> = {
  Heading1: heading(
    1,
    '<w:spacing w:before="360" w:after="80"/>',
    `${THEME_HEADINGS}${ACCENT1_DARK}${SIZE(40)}`,
  ),
  Heading2: heading(
    2,
    '<w:spacing w:before="160" w:after="80"/>',
    `${THEME_HEADINGS}${ACCENT1_DARK}${SIZE(32)}`,
  ),
  Heading3: heading(
    3,
    '<w:spacing w:before="160" w:after="80"/>',
    `${THEME_HEADINGS_EAST_ASIA}${ACCENT1_DARK}${SIZE(28)}`,
  ),
  Heading4: heading(
    4,
    '<w:spacing w:before="80" w:after="40"/>',
    `${THEME_HEADINGS_EAST_ASIA}${ITALIC}${ACCENT1_DARK}`,
  ),
  Heading5: heading(
    5,
    '<w:spacing w:before="80" w:after="40"/>',
    `${THEME_HEADINGS_EAST_ASIA}${ACCENT1_DARK}`,
  ),
  Heading6: heading(
    6,
    '<w:spacing w:before="40" w:after="0"/>',
    `${THEME_HEADINGS_EAST_ASIA}${ITALIC}${TEXT1_LIGHT_35}`,
  ),
  Heading7: heading(
    7,
    '<w:spacing w:before="40" w:after="0"/>',
    `${THEME_HEADINGS_EAST_ASIA}${TEXT1_LIGHT_35}`,
  ),
  Heading8: heading(
    8,
    '<w:spacing w:after="0"/>',
    `${THEME_HEADINGS_EAST_ASIA}${ITALIC}${TEXT1_LIGHT_15}`,
  ),
  Heading9: heading(9, '<w:spacing w:after="0"/>', `${THEME_HEADINGS_EAST_ASIA}${TEXT1_LIGHT_15}`),
  Title: {
    type: "paragraph",
    name: "Title",
    uiPriority: 10,
    visible: true,
    linked: true,
    pPr: '<w:spacing w:after="80" w:line="240" w:lineRule="auto"/><w:contextualSpacing/>',
    rPr: `${THEME_HEADINGS}<w:spacing w:val="-10"/><w:kern w:val="28"/>${SIZE(56)}`,
  },
  Subtitle: {
    type: "paragraph",
    name: "Subtitle",
    uiPriority: 11,
    visible: true,
    linked: true,
    pPr: '<w:spacing w:after="160"/>',
    rPr: `${THEME_HEADINGS_EAST_ASIA}${TEXT1_LIGHT_35}<w:spacing w:val="15"/>${SIZE(28)}`,
  },
  NoSpacing: {
    type: "paragraph",
    name: "No Spacing",
    uiPriority: 1,
    visible: true,
    pPr: '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>',
  },
  Quote: {
    type: "paragraph",
    name: "Quote",
    uiPriority: 29,
    visible: true,
    linked: true,
    pPr: '<w:spacing w:before="160"/><w:jc w:val="center"/>',
    rPr: `${ITALIC}${TEXT1_LIGHT_25}`,
  },
  IntenseQuote: {
    type: "paragraph",
    name: "Intense Quote",
    uiPriority: 30,
    visible: true,
    linked: true,
    pPr: '<w:pBdr><w:top w:val="single" w:sz="4" w:space="10" w:color="0F4761" w:themeColor="accent1" w:themeShade="BF"/><w:bottom w:val="single" w:sz="4" w:space="10" w:color="0F4761" w:themeColor="accent1" w:themeShade="BF"/></w:pBdr><w:spacing w:before="360" w:after="360"/><w:ind w:left="864" w:right="864"/><w:jc w:val="center"/>',
    rPr: `${ITALIC}${ACCENT1_DARK}`,
  },
  ListParagraph: {
    type: "paragraph",
    name: "List Paragraph",
    uiPriority: 34,
    visible: true,
    pPr: '<w:ind w:left="720"/><w:contextualSpacing/>',
  },
  Caption: {
    type: "paragraph",
    name: "caption",
    uiPriority: 35,
    visible: false,
    pPr: '<w:spacing w:after="200" w:line="240" w:lineRule="auto"/>',
    rPr: `${ITALIC}<w:color w:val="0E2841" w:themeColor="text2"/>${SIZE(18)}`,
  },
  TOCHeading: {
    type: "paragraph",
    name: "TOC Heading",
    uiPriority: 39,
    visible: false,
    basedOn: "Heading1",
    pPr: '<w:outlineLvl w:val="9"/>',
  },
  SubtleEmphasis: {
    type: "character",
    name: "Subtle Emphasis",
    uiPriority: 19,
    visible: true,
    rPr: `${ITALIC}${TEXT1_LIGHT_25}`,
  },
  Emphasis: { type: "character", name: "Emphasis", uiPriority: 20, visible: true, rPr: ITALIC },
  IntenseEmphasis: {
    type: "character",
    name: "Intense Emphasis",
    uiPriority: 21,
    visible: true,
    rPr: `${ITALIC}${ACCENT1_DARK}`,
  },
  Strong: { type: "character", name: "Strong", uiPriority: 22, visible: true, rPr: BOLD },
  SubtleReference: {
    type: "character",
    name: "Subtle Reference",
    uiPriority: 31,
    visible: true,
    rPr: '<w:smallCaps/><w:color w:val="5A5A5A" w:themeColor="text1" w:themeTint="A5"/>',
  },
  IntenseReference: {
    type: "character",
    name: "Intense Reference",
    uiPriority: 32,
    visible: true,
    rPr: `${BOLD}<w:smallCaps/>${ACCENT1_DARK}<w:spacing w:val="5"/>`,
  },
  BookTitle: {
    type: "character",
    name: "Book Title",
    uiPriority: 33,
    visible: true,
    rPr: `${BOLD}${ITALIC}<w:spacing w:val="5"/>`,
  },
};

/** A built-in style Word knows by id, whether or not the document defines it yet. */
export interface BuiltinStyleInfo {
  readonly styleId: string;
  readonly type: "paragraph" | "character";
  readonly name: string;
  readonly uiPriority: number;
  /** In Word's gallery from the start (the rest appear once used). */
  readonly quickStyle: boolean;
}

/** Word's built-in quick styles this library can create, by style id. */
export function builtinStyles(): BuiltinStyleInfo[] {
  return Object.entries(BUILTIN_STYLES).map(([styleId, s]) => ({
    styleId,
    type: s.type,
    name: s.name,
    uiPriority: s.uiPriority,
    quickStyle: s.visible,
  }));
}

function styleXml(id: string, s: BuiltinStyle, linkId: string | undefined): string {
  const flags = s.visible ? "<w:qFormat/>" : "<w:semiHidden/><w:unhideWhenUsed/><w:qFormat/>";
  const base = s.basedOn ?? (s.type === "paragraph" ? "Normal" : "DefaultParagraphFont");
  return [
    `<w:style xmlns:w="${WML_NS}" w:type="${s.type}" w:styleId="${id}">`,
    `<w:name w:val="${s.name}"/>`,
    `<w:basedOn w:val="${base}"/>`,
    s.type === "paragraph" ? '<w:next w:val="Normal"/>' : "",
    linkId ? `<w:link w:val="${linkId}"/>` : "",
    `<w:uiPriority w:val="${s.uiPriority}"/>`,
    flags,
    s.pPr ? `<w:pPr>${s.pPr}</w:pPr>` : "",
    s.rPr ? `<w:rPr>${s.rPr}</w:rPr>` : "",
    "</w:style>",
  ].join("");
}

function linkedCharXml(id: string, paragraphId: string, s: BuiltinStyle): string {
  return [
    `<w:style xmlns:w="${WML_NS}" w:type="character" w:customStyle="1" w:styleId="${id}">`,
    `<w:name w:val="${s.name.replace(/^heading/, "Heading")} Char"/>`,
    '<w:basedOn w:val="DefaultParagraphFont"/>',
    `<w:link w:val="${paragraphId}"/>`,
    `<w:uiPriority w:val="${s.uiPriority}"/>`,
    s.rPr ? `<w:rPr>${s.rPr}</w:rPr>` : "",
    "</w:style>",
  ].join("");
}

function parseStyle(xml: string): XmlElement {
  const root = parseXml(xml).root;
  // The namespace declaration was only needed to parse the snippet;
  // styles.xml declares `w` on its root.
  return { ...root, attrs: root.attrs.filter((a) => !a.isNamespaceDecl) };
}

/**
 * Add one of Word's built-in styles (`Heading1` … `Heading9`, `Title`,
 * `Subtitle`, `Quote`, `IntenseQuote`, `NoSpacing`, `ListParagraph`,
 * `Strong`, `Emphasis`, …; see {@link builtinStyles}) with Word's definition,
 * as Word does the first time a latent style is applied. Paragraph styles
 * Word links to a character style get that `…Char` style too. A style the
 * document already defines is left alone. Returns whether it added anything.
 */
export function ensureBuiltinStyle(doc: Docx, styleId: string): boolean {
  const s = BUILTIN_STYLES[styleId];
  if (!s) return false;
  const existingIds = new Set(
    (stylesPart(doc)?.styles ?? []).map((el) => getElementAttr(el, "styleId")),
  );
  if (existingIds.has(styleId)) return false;
  const linkId = s.linked && !existingIds.has(`${styleId}Char`) ? `${styleId}Char` : undefined;
  // addStyle creates styles.xml when the package has none.
  addStyle(doc, { type: s.type, styleId });
  const part = stylesPart(doc);
  if (!part) return false;
  const index = part.styles.findIndex((el) => getElementAttr(el, "styleId") === styleId);
  part.styles[index] = parseStyle(styleXml(styleId, s, linkId));
  if (linkId) part.styles.push(parseStyle(linkedCharXml(linkId, styleId, s)));
  doc.stylesDirty = true;
  doc.dirty = true;
  return true;
}

/**
 * Edit a style's own formatting with the run and paragraph functions: `run`
 * receives a run whose `<w:rPr>` is the style's, `paragraph` a paragraph
 * whose `<w:pPr>` is the style's, so `setRunFormat`, {@link setRunColor},
 * `setParagraphSpacing` … all apply to the style (Word's Modify Style).
 * Returns `false` when the document has no style `styleId`.
 */
export function updateStyleFormatting(
  doc: Docx,
  styleId: string,
  edit: {
    readonly run?: (run: WmlRun) => void;
    readonly paragraph?: (paragraph: WmlParagraph) => void;
  },
): boolean {
  const style = stylesPart(doc)?.styles.find((el) => getElementAttr(el, "styleId") === styleId);
  if (!style) return false;
  if (edit.paragraph) {
    const view: WmlParagraph = { kind: "paragraph", children: [], extras: [] };
    const pPr = findChild(style, "pPr");
    if (pPr) view.pPr = pPr;
    edit.paragraph(view);
    replaceChild(
      style,
      "pPr",
      view.pPr && childElementsOf(view.pPr).length > 0 ? view.pPr : undefined,
    );
  }
  if (edit.run) {
    const view: WmlRun = { kind: "run", pieces: [], extras: [] };
    const rPr = findChild(style, "rPr");
    if (rPr) view.rPr = rPr;
    edit.run(view);
    replaceChild(
      style,
      "rPr",
      view.rPr && childElementsOf(view.rPr).length > 0 ? view.rPr : undefined,
    );
  }
  doc.stylesDirty = true;
  doc.dirty = true;
  return true;
}
