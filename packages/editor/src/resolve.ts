/**
 * Effective (style-resolved) formatting, as Word shows it.
 *
 * A run's look is the merge of, from weakest to strongest (ECMA-376
 * §17.7.2): the document defaults, the paragraph style and the styles it is
 * `basedOn`, the run's character style chain, and the run's direct
 * formatting. The canvas renders this, and the ribbon reports it, so a
 * Heading 1 paragraph reads "Calibri Light / 16 / bold" like it does in Word
 * instead of showing nothing because the run itself carries no `<w:rPr>`.
 *
 * Toggle properties (`b`, `i`, `strike`, `caps` …, §17.7.3) toggle rather
 * than override across the style levels — table style, paragraph style,
 * character style — so a bold character style on a bold paragraph style
 * gives regular text. Within one style's `basedOn` chain, and in direct
 * formatting, they are plain values.
 *
 * Simplifications, each one a place where Word can differ:
 * - Table style text formatting (passed in as {@link CellTextFormat}) sits
 *   between the document defaults and the paragraph style, as §17.7.2
 *   orders it.
 */

import {
  childElementsOf,
  type Docx,
  getDocumentSetting,
  getElementAttr,
  getParagraphStyle,
  getRunProp,
  stylesPart,
  type WmlParagraph,
  type WmlRun,
  type XmlElement,
} from "@office-kit/docx";
import { createNumberingResolver } from "./list-numbering.js";
import { resolveThemeColor, type ThemePalette, themePalette, themeRoot } from "./theme-color.js";

/** A border line as resolved for display. */
export interface ResolvedBorder {
  /** ST_Border, e.g. `single`; `none` / `nil` draw nothing. */
  readonly style: string;
  /** Eighths of a point. */
  readonly size: number;
  /** Points between the border and the text. */
  readonly space: number;
  /** Hex RGB (theme colour applied) or `auto`. */
  readonly color: string;
}

/** Shading as resolved for display. */
export interface ResolvedShading {
  /** ST_Shd pattern. */
  readonly pattern: string;
  /** Hex RGB (theme colour applied) or `auto`. */
  readonly fill: string;
  /** Pattern colour, hex RGB or `auto`. */
  readonly color: string;
}

// Word's implicit stops when settings carry no defaultTabStop: every half inch.
const DEFAULT_TAB_STOP_TWIPS = 720;

/** A custom tab stop (`w:tab`, §17.3.1.37) as the canvas lays it out. */
export interface ResolvedTabStop {
  /** Twips from the text column's leading edge. */
  readonly position: number;
  /** `start` / `end` are folded into `left` / `right`; `num` lays out as `left`. */
  readonly align: "left" | "center" | "right" | "decimal";
  /** ST_TabTlc; `none` draws nothing. */
  readonly leader: string;
}

/** On-off run properties beyond bold / italic / strike, by their element name. */
export type RunToggle =
  | "caps"
  | "smallCaps"
  | "dstrike"
  | "outline"
  | "shadow"
  | "emboss"
  | "imprint"
  | "vanish";

const RUN_TOGGLES: ReadonlySet<string> = new Set<RunToggle>([
  "caps",
  "smallCaps",
  "dstrike",
  "outline",
  "shadow",
  "emboss",
  "imprint",
  "vanish",
]);

export interface ResolvedRunFormat {
  /** Font name for Latin text, theme fonts already looked up. */
  readonly font?: string | undefined;
  /** Set when the font comes from the theme, as Word labels it "(Body)" / "(Headings)". */
  readonly fontRole?: "body" | "headings" | undefined;
  /** Font for East Asian text (`w:eastAsia` / `w:eastAsiaTheme`). */
  readonly eastAsiaFont?: string | undefined;
  /** Font for complex-script text (`w:cs` / `w:cstheme`): Arabic, Hebrew, Thai … */
  readonly csFont?: string | undefined;
  readonly sizeHalfPoints?: number | undefined;
  /** Size of complex-script text (`w:szCs`). */
  readonly csSizeHalfPoints?: number | undefined;
  readonly bold: boolean;
  readonly italic: boolean;
  /** Bold / italic of complex-script text (`w:bCs` / `w:iCs`). */
  readonly csBold: boolean;
  readonly csItalic: boolean;
  /**
   * The run is complex script whatever its characters (`w:cs`, §17.3.2.7),
   * or right-to-left (`w:rtl`, §17.3.2.30), which is laid out as complex script.
   */
  readonly complexScript: boolean;
  readonly rtl: boolean;
  readonly strike: boolean;
  /** The other on-off effects that are on. */
  readonly toggles: ReadonlySet<RunToggle>;
  /** `w:u/@w:val`; `"none"` and absent both mean no underline. */
  readonly underline?: string | undefined;
  /** Underline colour, hex RGB (theme colour applied). */
  readonly underlineColor?: string | undefined;
  /** Hex RGB (theme colour applied) or `"auto"`. */
  readonly color?: string | undefined;
  /** `w:color/@w:themeColor`, when the colour comes from the theme. */
  readonly themeColor?: string | undefined;
  readonly highlight?: string | undefined;
  readonly vertAlign?: string | undefined;
  /** Character spacing in twips (`w:spacing`), negative = condensed. */
  readonly spacing?: number | undefined;
  /** Raised (+) / lowered (−) position in half-points (`w:position`). */
  readonly position?: number | undefined;
  /** Horizontal scale in percent (`w:w`). */
  readonly scale?: number | undefined;
  /** Kerning threshold in half-points (`w:kern`). */
  readonly kern?: number | undefined;
  readonly shading?: ResolvedShading | undefined;
  readonly border?: ResolvedBorder | undefined;
  /** The run's character style id. */
  readonly characterStyle?: string | undefined;
  /** Emphasis mark (`w:em`, §17.3.2.12): `dot`, `comma`, `circle` or `underDot`. */
  readonly emphasis?: string | undefined;
  /** East Asian layout (`w:eastAsianLayout`, §17.3.2.10). */
  readonly eastAsianLayout?: ResolvedEastAsianLayout | undefined;
  /** Fit-to-width text (`w:fitText`, §17.3.2.14). */
  readonly fitText?: { readonly width: number; readonly id?: number | undefined } | undefined;
}

/** Word's Asian Layout: combined characters, two lines in one, horizontal-in-vertical. */
export interface ResolvedEastAsianLayout {
  /** Combine Characters, or with brackets Two Lines in One: the text set in two half-size lines. */
  readonly combine: boolean;
  /** ST_CombineBrackets around two-lines-in-one text: `round`, `square`, `angle`, `curly`. */
  readonly brackets?: string | undefined;
  /** Horizontal in vertical (縦中横): upright in vertical text. */
  readonly vert: boolean;
  readonly vertCompress: boolean;
}

/** On-off paragraph properties the Paragraph dialog shows, by element name. */
export type ParagraphToggle =
  | "keepNext"
  | "keepLines"
  | "pageBreakBefore"
  | "widowControl"
  | "suppressLineNumbers"
  | "suppressAutoHyphens"
  | "contextualSpacing"
  | "mirrorIndents"
  | "kinsoku"
  | "wordWrap"
  | "overflowPunct"
  | "topLinePunct"
  | "autoSpaceDE"
  | "autoSpaceDN"
  | "bidi"
  | "snapToGrid";

// Their values when nothing in the hierarchy sets them (§17.3.1): the East
// Asian line-breaking rules and grid snapping are on unless turned off;
// everything else is off.
const PARAGRAPH_TOGGLE_DEFAULTS: Readonly<Record<ParagraphToggle, boolean>> = {
  keepNext: false,
  keepLines: false,
  pageBreakBefore: false,
  widowControl: false,
  suppressLineNumbers: false,
  suppressAutoHyphens: false,
  contextualSpacing: false,
  mirrorIndents: false,
  kinsoku: true,
  wordWrap: true,
  overflowPunct: true,
  topLinePunct: false,
  autoSpaceDE: true,
  autoSpaceDN: true,
  bidi: false,
  snapToGrid: true,
};

export type ParagraphBorderSide = "top" | "left" | "bottom" | "right" | "between" | "bar";

export interface ResolvedParagraphFormat {
  /** `w:jc/@w:val`. */
  readonly alignment?: string | undefined;
  /** Twips. */
  readonly left?: number | undefined;
  readonly right?: number | undefined;
  readonly firstLine?: number | undefined;
  readonly hanging?: number | undefined;
  /**
   * Hundredths of a character (`w:leftChars` …, §17.3.1.12); where set, they
   * take precedence over the twips value of the same indent.
   */
  readonly leftChars?: number | undefined;
  readonly rightChars?: number | undefined;
  readonly firstLineChars?: number | undefined;
  readonly hangingChars?: number | undefined;
  readonly before?: number | undefined;
  readonly after?: number | undefined;
  /** Hundredths of a line (§17.3.1.33); where set, they take precedence over `before` / `after`. */
  readonly beforeLines?: number | undefined;
  readonly afterLines?: number | undefined;
  /** 240ths of a line when `lineRule` is `auto`, twips otherwise. */
  readonly line?: number | undefined;
  readonly lineRule?: string | undefined;
  /** 0–8 for the outline (Navigation pane) levels; 9 or absent is body text. */
  readonly outlineLevel?: number | undefined;
  readonly textAlignment?: string | undefined;
  readonly toggles: Readonly<Record<ParagraphToggle, boolean>>;
  readonly shading?: ResolvedShading | undefined;
  readonly borders: Readonly<Partial<Record<ParagraphBorderSide, ResolvedBorder>>>;
  /** The list the paragraph belongs to, directly or through its style. */
  readonly numbering?: { readonly numId: number; readonly ilvl: number } | undefined;
  /** Custom tab stops, by position; bar tabs are not stops and are left out. */
  readonly tabs: readonly ResolvedTabStop[];
  /**
   * A drop cap frame (`w:framePr w:dropCap`, §17.3.1.11): the paragraph holds
   * the letter, and the next paragraph wraps around it.
   */
  readonly dropCap?: { readonly kind: "drop" | "margin"; readonly hSpace: number } | undefined;
}

/** Paragraph / run properties a table style gives the text of a cell, weakest first. */
export interface CellTextFormat {
  readonly pPr: readonly XmlElement[];
  readonly rPr: readonly XmlElement[];
}

export interface StyleResolver {
  /** settings `w:defaultTabStop` in twips: the interval of the implicit stops (§17.15.1.25). */
  readonly defaultTabStop: number;
  /** A paragraph's formatting; `cell` adds a table style's formatting for text in a cell. */
  paragraph(para: WmlParagraph, cell?: CellTextFormat): ResolvedParagraphFormat;
  /** A run's formatting; without a run, the paragraph mark's (for empty paragraphs). */
  run(para: WmlParagraph, run?: WmlRun, cell?: CellTextFormat): ResolvedRunFormat;
  /**
   * The formatting of a list paragraph's number or bullet: the paragraph
   * mark's run properties with the list level's on top (§17.9.24).
   */
  listLabelRun(para: WmlParagraph, cell?: CellTextFormat): ResolvedRunFormat;
  /** A paragraph or character style's own formatting, for gallery previews. */
  style(styleId: string): { paragraph: ResolvedParagraphFormat; run: ResolvedRunFormat };
  /** The document's theme colours (or Word's default theme). */
  readonly palette: ThemePalette;
  /** The theme's headings / body Latin fonts, if the document has a theme. */
  readonly themeFonts: ThemeFonts;
}

// ST_OnOff false values (§17.17.4); a bare `<w:b/>` means on.
const OFF_VALUES: ReadonlySet<string> = new Set(["0", "false", "off"]);

type Mutable<T> = { -readonly [K in keyof T]: T[K] };

export interface ThemeFonts {
  readonly major?: string | undefined;
  readonly minor?: string | undefined;
  /** The theme's complex-script fonts (`a:cs`); usually empty, leaving the script's default. */
  readonly majorCs?: string | undefined;
  readonly minorCs?: string | undefined;
}

function child(el: XmlElement | undefined, local: string): XmlElement | undefined {
  return el && childElementsOf(el).find((c) => c.name.local === local);
}

function intAttr(el: XmlElement, local: string): number | undefined {
  const raw = getElementAttr(el, local);
  const n = raw === undefined ? Number.NaN : Number(raw);
  return Number.isInteger(n) ? n : undefined;
}

function onOff(el: XmlElement): boolean {
  const val = getElementAttr(el, "val");
  return val === undefined || !OFF_VALUES.has(val);
}

/** An ST_OnOff attribute other than `val`; absent means off. */
function onOffAttr(el: XmlElement, local: string): boolean {
  const val = getElementAttr(el, local);
  return val !== undefined && !OFF_VALUES.has(val);
}

/** `<a:latin typeface>` of the theme's major / minor font, if there is a theme. */
export function readThemeFonts(doc: Docx): ThemeFonts {
  const scheme = child(child(themeRoot(doc), "themeElements"), "fontScheme");
  const typeface = (which: string, script: string): string | undefined => {
    const font = child(child(scheme, which), script);
    // An empty typeface means "the script's default font" (§20.1.4.1.16).
    return (font && getElementAttr(font, "typeface")) || undefined;
  };
  return {
    major: typeface("majorFont", "latin"),
    minor: typeface("minorFont", "latin"),
    majorCs: typeface("majorFont", "cs"),
    minorCs: typeface("minorFont", "cs"),
  };
}

interface Context {
  readonly theme: ThemeFonts;
  readonly palette: ThemePalette;
}

/** A colour attribute set (`val`/`color` + theme attributes) as display RGB. */
function colorOf(el: XmlElement, valName: string, ctx: Context): string | undefined {
  return resolveThemeColor(
    ctx.palette,
    getElementAttr(el, valName),
    getElementAttr(el, "themeColor"),
    getElementAttr(el, "themeTint"),
    getElementAttr(el, "themeShade"),
  );
}

function borderOf(el: XmlElement, ctx: Context): ResolvedBorder {
  return {
    style: getElementAttr(el, "val") ?? "none",
    size: intAttr(el, "sz") ?? 0,
    space: intAttr(el, "space") ?? 0,
    color: colorOf(el, "color", ctx) ?? "auto",
  };
}

function shadingOf(el: XmlElement, ctx: Context): ResolvedShading {
  return {
    pattern: getElementAttr(el, "val") ?? "clear",
    fill:
      resolveThemeColor(
        ctx.palette,
        getElementAttr(el, "fill"),
        getElementAttr(el, "themeFill"),
        getElementAttr(el, "themeFillTint"),
        getElementAttr(el, "themeFillShade"),
      ) ?? "auto",
    color: colorOf(el, "color", ctx) ?? "auto",
  };
}

/** A theme font reference (`minorHAnsi`, `majorEastAsia` …) looked up in the theme. */
function themeFont(
  ref: string,
  theme: ThemeFonts,
): { font: string | undefined; role: "body" | "headings" } {
  const major = ref.startsWith("major");
  return { font: major ? theme.major : theme.minor, role: major ? "headings" : "body" };
}

function applyFonts(out: Mutable<ResolvedRunFormat>, el: XmlElement, theme: ThemeFonts): void {
  // A theme attribute overrides the named font on the same element (§17.3.2.26).
  const asciiTheme = getElementAttr(el, "asciiTheme");
  const ascii = getElementAttr(el, "ascii");
  if (asciiTheme !== undefined) {
    const { font, role } = themeFont(asciiTheme, theme);
    if (font !== undefined) {
      out.font = font;
      out.fontRole = role;
    }
  } else if (ascii !== undefined) {
    out.font = ascii;
    out.fontRole = undefined;
  }
  const eastAsiaTheme = getElementAttr(el, "eastAsiaTheme");
  const eastAsia = getElementAttr(el, "eastAsia");
  if (eastAsiaTheme !== undefined) {
    const { font } = themeFont(eastAsiaTheme, theme);
    if (font !== undefined) out.eastAsiaFont = font;
  } else if (eastAsia !== undefined) {
    out.eastAsiaFont = eastAsia;
  }
  const csTheme = getElementAttr(el, "cstheme");
  const cs = getElementAttr(el, "cs");
  if (csTheme !== undefined) {
    out.csFont = csTheme.startsWith("major") ? theme.majorCs : theme.minorCs;
  } else if (cs !== undefined) {
    out.csFont = cs;
  }
}

// The toggle properties of §17.7.3. `dstrike` is not one of them.
const XOR_TOGGLES: ReadonlySet<string> = new Set([
  "b",
  "bCs",
  "i",
  "iCs",
  "strike",
  "caps",
  "smallCaps",
  "outline",
  "shadow",
  "emboss",
  "imprint",
  "vanish",
]);

function flipToggle(out: Mutable<ResolvedRunFormat>, local: string): void {
  const toggles = out.toggles as Set<RunToggle>;
  if (local === "b") out.bold = !out.bold;
  else if (local === "bCs") out.csBold = !out.csBold;
  else if (local === "i") out.italic = !out.italic;
  else if (local === "iCs") out.csItalic = !out.csItalic;
  else if (local === "strike") out.strike = !out.strike;
  else if (toggles.has(local as RunToggle)) toggles.delete(local as RunToggle);
  else toggles.add(local as RunToggle);
}

/**
 * One level of the style hierarchy (§17.7.2): its run properties, weakest
 * first (a `basedOn` chain, or a table style's conditional formats). Within
 * the level the strongest value of each property wins; a toggle property the
 * level turns on then flips the state built up so far (§17.7.3).
 */
function applyStyleLevel(
  out: Mutable<ResolvedRunFormat>,
  rPrs: readonly (XmlElement | undefined)[],
  ctx: Context,
): void {
  const toggled = new Map<string, boolean>();
  for (const rPr of rPrs) {
    if (!rPr) continue;
    applyRPr(out, rPr, ctx, true);
    for (const el of childElementsOf(rPr)) {
      if (XOR_TOGGLES.has(el.name.local)) toggled.set(el.name.local, onOff(el));
    }
  }
  for (const [local, on] of toggled) if (on) flipToggle(out, local);
}

function applyRPr(
  out: Mutable<ResolvedRunFormat>,
  rPr: XmlElement | undefined,
  ctx: Context,
  skipXorToggles = false,
): void {
  if (!rPr) return;
  const toggles = out.toggles as Set<RunToggle>;
  for (const el of childElementsOf(rPr)) {
    const local = el.name.local;
    if (skipXorToggles && XOR_TOGGLES.has(local)) continue;
    if (RUN_TOGGLES.has(local)) {
      if (onOff(el)) toggles.add(local as RunToggle);
      else toggles.delete(local as RunToggle);
      continue;
    }
    switch (local) {
      case "b":
        out.bold = onOff(el);
        break;
      case "i":
        out.italic = onOff(el);
        break;
      case "bCs":
        out.csBold = onOff(el);
        break;
      case "iCs":
        out.csItalic = onOff(el);
        break;
      case "cs":
        out.complexScript = onOff(el);
        break;
      case "rtl":
        out.rtl = onOff(el);
        break;
      case "szCs": {
        const size = intAttr(el, "val");
        if (size !== undefined) out.csSizeHalfPoints = size;
        break;
      }
      case "strike":
        out.strike = onOff(el);
        break;
      case "u":
        out.underline = getElementAttr(el, "val");
        out.underlineColor =
          getElementAttr(el, "color") === undefined ? undefined : colorOf(el, "color", ctx);
        break;
      case "color":
        out.color = colorOf(el, "val", ctx);
        out.themeColor = getElementAttr(el, "themeColor");
        break;
      case "highlight":
        out.highlight = getElementAttr(el, "val");
        break;
      case "vertAlign":
        out.vertAlign = getElementAttr(el, "val");
        break;
      case "sz": {
        // Universal measures ("12pt") are rare in sz; integer half-points only.
        const size = intAttr(el, "val");
        if (size !== undefined) out.sizeHalfPoints = size;
        break;
      }
      case "spacing":
        out.spacing = intAttr(el, "val");
        break;
      case "position":
        out.position = intAttr(el, "val");
        break;
      case "w":
        out.scale = intAttr(el, "val");
        break;
      case "kern":
        out.kern = intAttr(el, "val");
        break;
      case "em": {
        const val = getElementAttr(el, "val");
        out.emphasis = val === undefined || val === "none" ? undefined : val;
        break;
      }
      case "eastAsianLayout": {
        const brackets = getElementAttr(el, "combineBrackets");
        out.eastAsianLayout = {
          combine: onOffAttr(el, "combine"),
          brackets: brackets === "none" ? undefined : brackets,
          vert: onOffAttr(el, "vert"),
          vertCompress: onOffAttr(el, "vertCompress"),
        };
        break;
      }
      case "fitText": {
        const width = intAttr(el, "val");
        out.fitText = width ? { width, id: intAttr(el, "id") } : undefined;
        break;
      }
      case "shd":
        out.shading = shadingOf(el, ctx);
        break;
      case "bdr":
        out.border = borderOf(el, ctx);
        break;
      case "rFonts":
        applyFonts(out, el, ctx.theme);
        break;
      default:
        break;
    }
  }
}

const PARAGRAPH_TOGGLE_NAMES: ReadonlySet<string> = new Set(Object.keys(PARAGRAPH_TOGGLE_DEFAULTS));
const PARAGRAPH_BORDER_SIDES: ReadonlySet<string> = new Set([
  "top",
  "left",
  "bottom",
  "right",
  "between",
  "bar",
]);
// `w:outlineLvl` 9 means body text (§17.3.1.20).
const BODY_TEXT_LEVEL = 9;

function applyPPr(
  out: Mutable<ResolvedParagraphFormat>,
  pPr: XmlElement | undefined,
  ctx: Context,
): void {
  if (!pPr) return;
  const toggles = out.toggles as Record<ParagraphToggle, boolean>;
  for (const el of childElementsOf(pPr)) {
    const local = el.name.local;
    if (PARAGRAPH_TOGGLE_NAMES.has(local)) {
      toggles[local as ParagraphToggle] = onOff(el);
      continue;
    }
    switch (local) {
      case "jc":
        out.alignment = getElementAttr(el, "val");
        break;
      case "textAlignment":
        out.textAlignment = getElementAttr(el, "val");
        break;
      case "outlineLvl": {
        const level = intAttr(el, "val");
        out.outlineLevel = level === BODY_TEXT_LEVEL ? undefined : level;
        break;
      }
      case "shd":
        out.shading = shadingOf(el, ctx);
        break;
      case "pBdr": {
        // A pBdr replaces the inherited box as a whole.
        const borders: Partial<Record<ParagraphBorderSide, ResolvedBorder>> = {};
        for (const side of childElementsOf(el)) {
          if (PARAGRAPH_BORDER_SIDES.has(side.name.local))
            borders[side.name.local as ParagraphBorderSide] = borderOf(side, ctx);
        }
        out.borders = borders;
        break;
      }
      case "numPr": {
        const numId = child(el, "numId");
        const ilvl = child(el, "ilvl");
        const id = numId && intAttr(numId, "val");
        if (id !== undefined) {
          // numId 0 removes the numbering a style would apply (§17.9.18).
          out.numbering =
            id === 0 ? undefined : { numId: id, ilvl: (ilvl && intAttr(ilvl, "val")) ?? 0 };
        } else if (out.numbering && ilvl) {
          out.numbering = { ...out.numbering, ilvl: intAttr(ilvl, "val") ?? 0 };
        }
        break;
      }
      case "ind": {
        // Each attribute of <w:ind> inherits on its own: a style that sets
        // only `left` keeps the `hanging` from below it. `start` / `end` are
        // the bidi-aware names for `left` / `right` (§17.3.1.12).
        const left = intAttr(el, "left") ?? intAttr(el, "start");
        const right = intAttr(el, "right") ?? intAttr(el, "end");
        const leftChars = intAttr(el, "leftChars") ?? intAttr(el, "startChars");
        const rightChars = intAttr(el, "rightChars") ?? intAttr(el, "endChars");
        const firstLine = intAttr(el, "firstLine");
        const hanging = intAttr(el, "hanging");
        const firstLineChars = intAttr(el, "firstLineChars");
        const hangingChars = intAttr(el, "hangingChars");
        if (left !== undefined) out.left = left;
        if (right !== undefined) out.right = right;
        if (leftChars !== undefined) out.leftChars = leftChars;
        if (rightChars !== undefined) out.rightChars = rightChars;
        // firstLine and hanging exclude each other; the later level replaces both.
        if ([firstLine, hanging, firstLineChars, hangingChars].some((v) => v !== undefined)) {
          out.firstLine = firstLine;
          out.hanging = hanging;
          out.firstLineChars = firstLineChars;
          out.hangingChars = hangingChars;
        }
        break;
      }
      case "tabs":
        out.tabs = mergeTabs(out.tabs, childElementsOf(el));
        break;
      case "framePr": {
        const kind = getElementAttr(el, "dropCap");
        out.dropCap =
          kind === "drop" || kind === "margin"
            ? { kind, hSpace: intAttr(el, "hSpace") ?? 0 }
            : undefined;
        break;
      }
      case "spacing": {
        const before = intAttr(el, "before");
        const after = intAttr(el, "after");
        const line = intAttr(el, "line");
        const beforeLines = intAttr(el, "beforeLines");
        const afterLines = intAttr(el, "afterLines");
        if (before !== undefined) out.before = before;
        if (after !== undefined) out.after = after;
        if (beforeLines !== undefined) out.beforeLines = beforeLines;
        if (afterLines !== undefined) out.afterLines = afterLines;
        if (line !== undefined) {
          out.line = line;
          // A missing lineRule means auto (§17.3.1.33).
          out.lineRule = getElementAttr(el, "lineRule") ?? "auto";
        }
        break;
      }
      default:
        break;
    }
  }
}

function emptyRun(): Mutable<ResolvedRunFormat> {
  return {
    bold: false,
    italic: false,
    strike: false,
    csBold: false,
    csItalic: false,
    complexScript: false,
    rtl: false,
    toggles: new Set(),
  };
}

function emptyParagraph(): Mutable<ResolvedParagraphFormat> {
  return { toggles: { ...PARAGRAPH_TOGGLE_DEFAULTS }, borders: {}, tabs: [] };
}

const TAB_ALIGN: Readonly<Record<string, ResolvedTabStop["align"]>> = {
  left: "left",
  start: "left",
  num: "left",
  center: "center",
  right: "right",
  end: "right",
  decimal: "decimal",
};

/**
 * Tab stops accumulate down the style hierarchy: a level adds its stops, and
 * a `clear` stop removes the inherited one at that position (§17.3.1.37).
 */
function mergeTabs(
  inherited: readonly ResolvedTabStop[],
  tabs: readonly XmlElement[],
): ResolvedTabStop[] {
  const byPosition = new Map(inherited.map((t) => [t.position, t]));
  for (const tab of tabs) {
    const position = intAttr(tab, "pos");
    const val = getElementAttr(tab, "val");
    if (position === undefined || val === undefined) continue;
    const align = TAB_ALIGN[val];
    if (align) {
      byPosition.set(position, {
        position,
        align,
        leader: getElementAttr(tab, "leader") ?? "none",
      });
    } else if (val === "clear") {
      // A `bar` tab, the other value, draws a rule but stops nothing.
      byPosition.delete(position);
    }
  }
  return [...byPosition.values()].toSorted((a, b) => a.position - b.position);
}

/**
 * Build a resolver over the document's current styles. Build a new one after
 * the styles change; within one render the styles are fixed.
 */
export function createStyleResolver(doc: Docx): StyleResolver {
  const part = stylesPart(doc);
  const byId = new Map<string, XmlElement>();
  let defaultParagraphStyle: string | undefined;
  for (const style of part?.styles ?? []) {
    const id = getElementAttr(style, "styleId");
    if (id === undefined) continue;
    byId.set(id, style);
    if (getElementAttr(style, "type") === "paragraph" && getElementAttr(style, "default") === "1") {
      defaultParagraphStyle = id;
    }
  }
  const defaults = part?.docDefaults;
  const defaultRPr = child(child(defaults, "rPrDefault"), "rPr");
  const defaultPPr = child(child(defaults, "pPrDefault"), "pPr");
  const ctx: Context = { theme: readThemeFonts(doc), palette: themePalette(doc) };

  /** The style and its `basedOn` ancestors, weakest (root) first. */
  const chain = (id: string | undefined): XmlElement[] => {
    const out: XmlElement[] = [];
    const seen = new Set<string>();
    let current = id;
    // A basedOn cycle is invalid but occurs in the wild; stop at the repeat.
    while (current !== undefined && !seen.has(current)) {
      seen.add(current);
      const style = byId.get(current);
      if (!style) break;
      out.unshift(style);
      const basedOn = child(style, "basedOn");
      current = basedOn && getElementAttr(basedOn, "val");
    }
    return out;
  };

  // A pStyle that names a missing style falls back to the default style, as in Word.
  const paragraphChain = (id: string | undefined): XmlElement[] =>
    id !== undefined && byId.has(id) ? chain(id) : chain(defaultParagraphStyle);

  const styleRPrs = (styles: readonly XmlElement[]) => styles.map((s) => child(s, "rPr"));

  const levels = createNumberingResolver(doc);
  const layer = (
    styleId: string | undefined,
    pPr: XmlElement | undefined,
    listPPr: XmlElement | undefined,
    cell?: CellTextFormat,
  ) => {
    const out = emptyParagraph();
    applyPPr(out, defaultPPr, ctx);
    for (const pPr of cell?.pPr ?? []) applyPPr(out, pPr, ctx);
    for (const style of paragraphChain(styleId)) applyPPr(out, child(style, "pPr"), ctx);
    applyPPr(out, listPPr, ctx);
    applyPPr(out, pPr, ctx);
    return out;
  };
  // A list level's indents apply between the style and the paragraph's own
  // formatting (§17.9.23), so a list item without direct indents is indented
  // by its level.
  const paragraphOf = (
    styleId: string | undefined,
    pPr: XmlElement | undefined,
    cell?: CellTextFormat,
  ) => {
    const base = layer(styleId, pPr, undefined, cell);
    const list = base.numbering && levels.level(base.numbering.numId, base.numbering.ilvl);
    const listPPr = list && child(list.lvl, "pPr");
    return listPPr ? layer(styleId, pPr, listPPr, cell) : base;
  };

  const runOf = (para: WmlParagraph, run?: WmlRun, cell?: CellTextFormat) => {
    const out = emptyRun();
    applyRPr(out, defaultRPr, ctx);
    applyStyleLevel(out, cell?.rPr ?? [], ctx);
    applyStyleLevel(out, styleRPrs(paragraphChain(getParagraphStyle(para))), ctx);
    if (run) {
      const characterStyle = getRunProp(run, "rStyle").val;
      applyStyleLevel(out, styleRPrs(chain(characterStyle)), ctx);
      applyRPr(out, run.rPr, ctx);
      out.characterStyle = characterStyle;
    }
    return out;
  };

  return {
    palette: ctx.palette,
    themeFonts: ctx.theme,
    defaultTabStop: Number(getDocumentSetting(doc, "defaultTabStop").val) || DEFAULT_TAB_STOP_TWIPS,
    paragraph(para, cell) {
      return paragraphOf(getParagraphStyle(para), para.pPr, cell);
    },
    run: runOf,
    listLabelRun(para, cell) {
      const out = runOf(para, undefined, cell);
      applyRPr(out, child(para.pPr, "rPr"), ctx);
      const numbering = paragraphOf(getParagraphStyle(para), para.pPr, cell).numbering;
      const level = numbering && levels.level(numbering.numId, numbering.ilvl);
      applyRPr(out, child(level?.lvl, "rPr"), ctx);
      return out;
    },
    style(styleId) {
      const style = byId.get(styleId);
      const isCharacter = style !== undefined && getElementAttr(style, "type") === "character";
      const paragraph = paragraphOf(isCharacter ? undefined : styleId, undefined);
      const run = emptyRun();
      applyRPr(run, defaultRPr, ctx);
      applyStyleLevel(run, styleRPrs(paragraphChain(isCharacter ? undefined : styleId)), ctx);
      if (isCharacter) applyStyleLevel(run, styleRPrs(chain(styleId)), ctx);
      return { paragraph, run };
    },
  };
}

/** Page size and margins in twips, for laying out the editing canvas. */
export interface PageGeometry {
  readonly width: number;
  readonly height: number;
  readonly top: number;
  readonly right: number;
  readonly bottom: number;
  readonly left: number;
}

// What Word uses when the body has no `<w:pgSz>` / `<w:pgMar>`: it saves such
// a document with US Letter and 1 in margins (observed in Word for Mac).
const DEFAULT_PAGE: PageGeometry = {
  width: 12240,
  height: 15840,
  top: 1440,
  right: 1440,
  bottom: 1440,
  left: 1440,
};

/** The page of the body's final section (`w:body/w:sectPr`). */
export function pageGeometry(doc: Docx): PageGeometry {
  const sectPr = doc.document.body.sectPr;
  const size = child(sectPr, "pgSz");
  const margins = child(sectPr, "pgMar");
  const read = (el: XmlElement | undefined, local: string, fallback: number): number =>
    (el && intAttr(el, local)) ?? fallback;
  return {
    width: read(size, "w", DEFAULT_PAGE.width),
    height: read(size, "h", DEFAULT_PAGE.height),
    // Negative top/bottom margins mean "text may overlap the header"; the
    // distance from the page edge is the absolute value (§17.6.11).
    top: Math.abs(read(margins, "top", DEFAULT_PAGE.top)),
    right: read(margins, "right", DEFAULT_PAGE.right),
    bottom: Math.abs(read(margins, "bottom", DEFAULT_PAGE.bottom)),
    left: read(margins, "left", DEFAULT_PAGE.left),
  };
}
