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
 * Simplifications, each one a place where Word can differ:
 * - Toggle properties (`b`, `i`, `strike`, …) are treated as plain overrides;
 *   the spec XORs them across style levels (§17.7.3), which only matters
 *   when two levels of the hierarchy both set the same toggle.
 * - Numbering-level run properties are not applied.
 * - Table style text formatting (passed in as {@link CellTextFormat}) sits
 *   between the document defaults and the paragraph style, as §17.7.2
 *   orders it.
 */

import {
  childElementsOf,
  type Docx,
  getElementAttr,
  getParagraphStyle,
  getRawPartRoot,
  getRunProp,
  stylesPart,
  type WmlParagraph,
  type WmlRun,
  type XmlElement,
  xmlPartNames,
} from "@office-kit/docx";

export interface ResolvedRunFormat {
  /** Font name for Latin text, theme fonts already looked up. */
  readonly font?: string | undefined;
  /** Set when the font comes from the theme, as Word labels it "(Body)" / "(Headings)". */
  readonly fontRole?: "body" | "headings" | undefined;
  readonly sizeHalfPoints?: number | undefined;
  readonly bold: boolean;
  readonly italic: boolean;
  readonly strike: boolean;
  /** `w:u/@w:val`; `"none"` and absent both mean no underline. */
  readonly underline?: string | undefined;
  /** Hex RGB or `"auto"`. */
  readonly color?: string | undefined;
  readonly highlight?: string | undefined;
  readonly vertAlign?: string | undefined;
}

export interface ResolvedParagraphFormat {
  /** `w:jc/@w:val`. */
  readonly alignment?: string | undefined;
  /** Twips. */
  readonly left?: number | undefined;
  readonly right?: number | undefined;
  readonly firstLine?: number | undefined;
  readonly hanging?: number | undefined;
  readonly before?: number | undefined;
  readonly after?: number | undefined;
  /** 240ths of a line when `lineRule` is `auto`, twips otherwise. */
  readonly line?: number | undefined;
  readonly lineRule?: string | undefined;
}

/** Paragraph / run properties a table style gives the text of a cell, weakest first. */
export interface CellTextFormat {
  readonly pPr: readonly XmlElement[];
  readonly rPr: readonly XmlElement[];
}

export interface StyleResolver {
  /** A paragraph's formatting; `cell` adds a table style's formatting for text in a cell. */
  paragraph(para: WmlParagraph, cell?: CellTextFormat): ResolvedParagraphFormat;
  /** A run's formatting; without a run, the paragraph mark's (for empty paragraphs). */
  run(para: WmlParagraph, run?: WmlRun, cell?: CellTextFormat): ResolvedRunFormat;
}

// ST_OnOff false values (§17.17.4); a bare `<w:b/>` means on.
const OFF_VALUES: ReadonlySet<string> = new Set(["0", "false", "off"]);
const THEME_PART_DIR = "/word/theme/";

type Mutable<T> = { -readonly [K in keyof T]: T[K] };

interface ThemeFonts {
  readonly major?: string | undefined;
  readonly minor?: string | undefined;
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

/** `<a:latin typeface>` of the theme's major / minor font, if there is a theme. */
function readThemeFonts(doc: Docx): ThemeFonts {
  // The document part has exactly one theme relationship (§14.2.7); Word
  // always stores it under /word/theme/.
  const name = xmlPartNames(doc).find((n) => n.startsWith(THEME_PART_DIR));
  const root = name ? getRawPartRoot(doc, name) : undefined;
  const scheme = child(child(root, "themeElements"), "fontScheme");
  const latin = (which: string): string | undefined => {
    const typeface = child(child(scheme, which), "latin");
    return typeface && getElementAttr(typeface, "typeface");
  };
  return { major: latin("majorFont"), minor: latin("minorFont") };
}

function applyRPr(
  out: Mutable<ResolvedRunFormat>,
  rPr: XmlElement | undefined,
  theme: ThemeFonts,
): void {
  if (!rPr) return;
  for (const el of childElementsOf(rPr)) {
    switch (el.name.local) {
      case "b":
        out.bold = onOff(el);
        break;
      case "i":
        out.italic = onOff(el);
        break;
      case "strike":
        out.strike = onOff(el);
        break;
      case "u":
        out.underline = getElementAttr(el, "val");
        break;
      case "color":
        out.color = getElementAttr(el, "val");
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
      case "rFonts": {
        // An explicit font wins over a theme font on the same element (§17.3.2.26).
        const ascii = getElementAttr(el, "ascii");
        const themeRef = getElementAttr(el, "asciiTheme");
        if (ascii !== undefined) {
          out.font = ascii;
          out.fontRole = undefined;
        } else if (themeRef !== undefined) {
          const major = themeRef.startsWith("major");
          const font = major ? theme.major : theme.minor;
          if (font !== undefined) {
            out.font = font;
            out.fontRole = major ? "headings" : "body";
          }
        }
        break;
      }
      default:
        break;
    }
  }
}

function applyPPr(out: Mutable<ResolvedParagraphFormat>, pPr: XmlElement | undefined): void {
  if (!pPr) return;
  const jc = child(pPr, "jc");
  if (jc) out.alignment = getElementAttr(jc, "val");
  // Each attribute of <w:ind> / <w:spacing> inherits on its own: a style that
  // sets only `after` keeps the `line` from below it.
  const ind = child(pPr, "ind");
  if (ind) {
    // `start` / `end` are the bidi-aware names for `left` / `right` (§17.3.1.12).
    const left = intAttr(ind, "left") ?? intAttr(ind, "start");
    const right = intAttr(ind, "right") ?? intAttr(ind, "end");
    const firstLine = intAttr(ind, "firstLine");
    const hanging = intAttr(ind, "hanging");
    if (left !== undefined) out.left = left;
    if (right !== undefined) out.right = right;
    // firstLine and hanging exclude each other; the later level replaces both.
    if (firstLine !== undefined || hanging !== undefined) {
      out.firstLine = firstLine;
      out.hanging = hanging;
    }
  }
  const spacing = child(pPr, "spacing");
  if (spacing) {
    const before = intAttr(spacing, "before");
    const after = intAttr(spacing, "after");
    const line = intAttr(spacing, "line");
    if (before !== undefined) out.before = before;
    if (after !== undefined) out.after = after;
    if (line !== undefined) {
      out.line = line;
      // A missing lineRule means auto (§17.3.1.33).
      out.lineRule = getElementAttr(spacing, "lineRule") ?? "auto";
    }
  }
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
  const theme = readThemeFonts(doc);

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
  const paragraphChain = (para: WmlParagraph): XmlElement[] => {
    const id = getParagraphStyle(para);
    return id !== undefined && byId.has(id) ? chain(id) : chain(defaultParagraphStyle);
  };

  return {
    paragraph(para, cell) {
      const out: Mutable<ResolvedParagraphFormat> = {};
      applyPPr(out, defaultPPr);
      for (const pPr of cell?.pPr ?? []) applyPPr(out, pPr);
      for (const style of paragraphChain(para)) applyPPr(out, child(style, "pPr"));
      applyPPr(out, para.pPr);
      return out;
    },
    run(para, run, cell) {
      const out: Mutable<ResolvedRunFormat> = { bold: false, italic: false, strike: false };
      applyRPr(out, defaultRPr, theme);
      for (const rPr of cell?.rPr ?? []) applyRPr(out, rPr, theme);
      for (const style of paragraphChain(para)) applyRPr(out, child(style, "rPr"), theme);
      if (run) {
        for (const style of chain(getRunProp(run, "rStyle").val))
          applyRPr(out, child(style, "rPr"), theme);
        applyRPr(out, run.rPr, theme);
      }
      return out;
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
