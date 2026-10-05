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
 * - Table styles and numbering-level run properties are not applied.
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
  type WmlTable,
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

export interface StyleResolver {
  paragraph(para: WmlParagraph): ResolvedParagraphFormat;
  /** A run's formatting; without a run, the paragraph mark's (for empty paragraphs). */
  run(para: WmlParagraph, run?: WmlRun): ResolvedRunFormat;
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
    paragraph(para) {
      const out: Mutable<ResolvedParagraphFormat> = {};
      applyPPr(out, defaultPPr);
      for (const style of paragraphChain(para)) applyPPr(out, child(style, "pPr"));
      applyPPr(out, para.pPr);
      return out;
    },
    run(para, run) {
      const out: Mutable<ResolvedRunFormat> = { bold: false, italic: false, strike: false };
      applyRPr(out, defaultRPr, theme);
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

/** One side of a table or cell border (`<w:top w:val w:sz w:color/>`). */
export interface BorderSpec {
  /** ST_Border, e.g. `single`, `double`, `dotted`; `none` / `nil` draw nothing. */
  readonly style: string;
  /** Eighths of a point. */
  readonly size: number;
  /** Hex RGB or `auto`. */
  readonly color: string;
}

export type BorderSide = "top" | "left" | "bottom" | "right" | "insideH" | "insideV";
export type CellMarginSide = "top" | "left" | "bottom" | "right";

export interface ResolvedTableFormat {
  /** Preferred table width (`w:tblW`); `undefined` lets the content decide. */
  readonly width?: { readonly type: string; readonly value: number } | undefined;
  /** Grid column widths in twips (`w:tblGrid`). */
  readonly columns: readonly number[];
  readonly borders: Readonly<Partial<Record<BorderSide, BorderSpec>>>;
  /** Default cell margins in twips (`w:tblCellMar`). */
  readonly cellMargins: Readonly<Partial<Record<CellMarginSide, number>>>;
  /**
   * Where the table's left edge sits, in twips from the text margin. Word 2013
   * and later (compatibility mode 15) place the border at `w:tblInd`; older
   * modes place the first cell's *text* there, so the border moves left by
   * the cell margin — what Word shows for a document in Compatibility Mode.
   */
  readonly leftEdge: number;
}

const BORDER_SIDES: readonly BorderSide[] = [
  "top",
  "left",
  "bottom",
  "right",
  "insideH",
  "insideV",
];
// `start` / `end` are the bidi-aware names for `left` / `right`.
const SIDE_ALIASES: Readonly<Record<string, BorderSide>> = { start: "left", end: "right" };

function readBorders(container: XmlElement | undefined): Partial<Record<BorderSide, BorderSpec>> {
  const out: Partial<Record<BorderSide, BorderSpec>> = {};
  if (!container) return out;
  for (const el of childElementsOf(container)) {
    const side = SIDE_ALIASES[el.name.local] ?? el.name.local;
    if (!(BORDER_SIDES as readonly string[]).includes(side)) continue;
    out[side as BorderSide] = {
      style: getElementAttr(el, "val") ?? "none",
      size: intAttr(el, "sz") ?? 0,
      color: getElementAttr(el, "color") ?? "auto",
    };
  }
  return out;
}

function readCellMargins(
  container: XmlElement | undefined,
): Partial<Record<CellMarginSide, number>> {
  const out: Partial<Record<CellMarginSide, number>> = {};
  if (!container) return out;
  for (const el of childElementsOf(container)) {
    const side = SIDE_ALIASES[el.name.local] ?? el.name.local;
    // Only twips (dxa) margins; a pct / auto cell margin is not meaningful here.
    const type = getElementAttr(el, "type") ?? "dxa";
    const w = intAttr(el, "w");
    if (
      type === "dxa" &&
      w !== undefined &&
      (side === "top" || side === "left" || side === "bottom" || side === "right")
    ) {
      out[side] = w;
    }
  }
  return out;
}

/** A cell's own `<w:tcBorders>`, which override the table's for that cell. */
export function cellBorders(tcPr: XmlElement | undefined): Partial<Record<BorderSide, BorderSpec>> {
  return readBorders(child(tcPr, "tcBorders"));
}

/**
 * A table's effective width, grid, borders and cell margins: the default
 * table style (Word's "Normal Table"), then the table's style chain, then
 * the table's own `<w:tblPr>`.
 */
export function resolveTable(doc: Docx, table: WmlTable): ResolvedTableFormat {
  const part = stylesPart(doc);
  const styles = new Map<string, XmlElement>();
  let defaultTableStyle: string | undefined;
  for (const style of part?.styles ?? []) {
    const id = getElementAttr(style, "styleId");
    if (id === undefined) continue;
    styles.set(id, style);
    if (getElementAttr(style, "type") === "table" && getElementAttr(style, "default") === "1") {
      defaultTableStyle = id;
    }
  }
  const chainOf = (id: string | undefined): XmlElement[] => {
    const out: XmlElement[] = [];
    const seen = new Set<string>();
    let current = id;
    while (current !== undefined && !seen.has(current)) {
      seen.add(current);
      const style = styles.get(current);
      if (!style) break;
      out.unshift(style);
      const basedOn = child(style, "basedOn");
      current = basedOn && getElementAttr(basedOn, "val");
    }
    return out;
  };
  const styleRef = child(table.tblPr, "tblStyle");
  const styleId = styleRef && getElementAttr(styleRef, "val");
  const levels = [
    ...chainOf(defaultTableStyle).map((s) => child(s, "tblPr")),
    ...(styleId !== undefined && styleId !== defaultTableStyle
      ? chainOf(styleId).map((s) => child(s, "tblPr"))
      : []),
    table.tblPr,
  ];
  const borders: Partial<Record<BorderSide, BorderSpec>> = {};
  const cellMargins: Partial<Record<CellMarginSide, number>> = {};
  let indent = 0;
  for (const tblPr of levels) {
    Object.assign(borders, readBorders(child(tblPr, "tblBorders")));
    Object.assign(cellMargins, readCellMargins(child(tblPr, "tblCellMar")));
    const tblInd = child(tblPr, "tblInd");
    const w =
      tblInd && (getElementAttr(tblInd, "type") ?? "dxa") === "dxa"
        ? intAttr(tblInd, "w")
        : undefined;
    if (w !== undefined) indent = w;
  }
  const modern = compatibilityMode(doc) >= WORD_2013_MODE;
  const tblW = child(table.tblPr, "tblW");
  const widthValue = tblW && intAttr(tblW, "w");
  const columns = table.tblGrid
    ? childElementsOf(table.tblGrid)
        .filter((c) => c.name.local === "gridCol")
        .map((c) => intAttr(c, "w") ?? 0)
    : [];
  return {
    width:
      tblW && widthValue !== undefined
        ? { type: getElementAttr(tblW, "type") ?? "dxa", value: widthValue }
        : undefined,
    columns,
    borders,
    cellMargins,
    leftEdge: modern ? indent : indent - (cellMargins.left ?? 0),
  };
}

const WORD_2013_MODE = 15;
// A document without w:compatSetting compatibilityMode opens in Word as a
// Word 2007 document ("Compatibility Mode").
const DEFAULT_COMPATIBILITY_MODE = 12;
const SETTINGS_PART_SUFFIX = "/settings.xml";

/** `w:compatSetting[@w:name="compatibilityMode"]` from settings.xml. */
export function compatibilityMode(doc: Docx): number {
  const name = xmlPartNames(doc).find(
    (n) => n.startsWith("/word/") && n.endsWith(SETTINGS_PART_SUFFIX),
  );
  const root = name ? getRawPartRoot(doc, name) : undefined;
  const compat = child(root, "compat");
  const setting =
    compat &&
    childElementsOf(compat).find(
      (c) => c.name.local === "compatSetting" && getElementAttr(c, "name") === "compatibilityMode",
    );
  const mode = setting ? Number(getElementAttr(setting, "val")) : Number.NaN;
  return Number.isInteger(mode) ? mode : DEFAULT_COMPATIBILITY_MODE;
}
