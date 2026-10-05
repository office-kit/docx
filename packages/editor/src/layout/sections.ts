/**
 * The document's sections as the page layout needs them: which blocks each
 * one holds, and its `<w:sectPr>` (ECMA-376 Part 1 §17.6) read into plain
 * numbers — page size and margins, columns, vertical alignment, line
 * numbering, page borders, page numbering, note properties, the document
 * grid, and the header/footer part each page type shows (with inheritance
 * from earlier sections, §17.10.5).
 */

import {
  childElementsOf,
  type Docx,
  getDocumentSetting,
  getElementAttr,
  type HeaderFooterType,
  resolveHeaderFooter,
  sectionProperties,
  type XmlElement,
} from "@office-kit/docx";
import type { PageKind, SectionStart, VerticalAlign } from "./paginate.js";

/** Word's defaults for a new document: Letter, 1" margins, ½" header/footer distance. */
const DEFAULT_PAGE = { width: 12240, height: 15840 };
const DEFAULT_MARGINS = {
  top: 1440,
  right: 1440,
  bottom: 1440,
  left: 1440,
  header: 720,
  footer: 720,
  gutter: 0,
};
// `w:cols/@w:space` default: ½ inch (§17.6.4).
const DEFAULT_COLUMN_SPACE = 720;

export interface Margins {
  readonly top: number;
  readonly right: number;
  readonly bottom: number;
  readonly left: number;
  readonly header: number;
  readonly footer: number;
  readonly gutter: number;
  /**
   * A negative top/bottom margin (§17.6.11) is a fixed body edge: a tall
   * header overlaps the text instead of pushing it down.
   */
  readonly fixedTop: boolean;
  readonly fixedBottom: boolean;
}

export interface ColumnSpec {
  /** Twips. */
  readonly width: number;
  /** Twips of space after this column (none after the last). */
  readonly space: number;
}

export interface BorderLine {
  /** ST_Border, e.g. `single`, `double`, or an art border such as `apples`. */
  readonly style: string;
  /** Eighths of a point. */
  readonly size: number;
  /** Points between the border and the text or page edge. */
  readonly space: number;
  /** Hex RGB or `auto`. */
  readonly color: string;
}

export interface PageBorders {
  readonly offsetFrom: "text" | "page";
  readonly display: "allPages" | "firstPage" | "notFirstPage";
  readonly zOrder: "front" | "back";
  readonly top?: BorderLine | undefined;
  readonly left?: BorderLine | undefined;
  readonly bottom?: BorderLine | undefined;
  readonly right?: BorderLine | undefined;
}

export interface LineNumbering {
  readonly countBy: number;
  readonly start: number;
  readonly restart: "newPage" | "newSection" | "continuous";
  /** Twips between the numbers and the text; `undefined` is Word's automatic distance. */
  readonly distance?: number | undefined;
}

export interface NoteProperties {
  readonly numFmt?: string | undefined;
  readonly numStart?: number | undefined;
  readonly numRestart?: "continuous" | "eachSect" | "eachPage" | undefined;
  /** Endnotes: `sectEnd` or `docEnd`; footnotes: `pageBottom` / `beneathText`. */
  readonly pos?: string | undefined;
}

export interface DocumentGrid {
  readonly type: "default" | "lines" | "linesAndChars" | "snapToChars";
  /** Twips between grid lines. */
  readonly linePitch?: number | undefined;
}

export interface PageNumbering {
  readonly format?: string | undefined;
  readonly start?: number | undefined;
}

export interface SectionModel {
  readonly index: number;
  /** First and last body block of the section (inclusive); `last < first` for an empty one. */
  readonly firstBlock: number;
  readonly lastBlock: number;
  readonly start: SectionStart;
  /** Twips. */
  readonly pageWidth: number;
  readonly pageHeight: number;
  readonly margins: Margins;
  readonly columns: readonly ColumnSpec[];
  /** `w:cols/@w:sep`: a line between columns. */
  readonly separator: boolean;
  readonly vAlign: VerticalAlign;
  readonly titlePage: boolean;
  readonly pageNumbers: PageNumbering;
  readonly lineNumbers?: LineNumbering | undefined;
  readonly borders?: PageBorders | undefined;
  readonly grid?: DocumentGrid | undefined;
  readonly footnotePr?: NoteProperties | undefined;
  readonly endnotePr?: NoteProperties | undefined;
  /** Header/footer part shown on each page kind, after inheritance; `undefined` for none. */
  readonly headers: Readonly<Record<PageKind, string | undefined>>;
  readonly footers: Readonly<Record<PageKind, string | undefined>>;
}

function child(el: XmlElement | undefined, local: string): XmlElement | undefined {
  return el && childElementsOf(el).find((c) => c.name.local === local);
}

function num(el: XmlElement | undefined, local: string): number | undefined {
  const raw = el && getElementAttr(el, local);
  if (raw === undefined) return undefined;
  const n = Number(raw);
  return Number.isFinite(n) ? n : undefined;
}

const OFF: ReadonlySet<string> = new Set(["0", "false", "off"]);

function flag(el: XmlElement | undefined, local: string): boolean {
  const c = child(el, local);
  if (!c) return false;
  const val = getElementAttr(c, "val");
  return val === undefined || !OFF.has(val);
}

function onOffAttr(el: XmlElement | undefined, local: string): boolean {
  const raw = el && getElementAttr(el, local);
  return raw !== undefined && !OFF.has(raw);
}

function oneOf<T extends string>(value: string | undefined, allowed: readonly T[], fallback: T): T {
  return allowed.find((a) => a === value) ?? fallback;
}

const STARTS: readonly SectionStart[] = [
  "nextPage",
  "continuous",
  "evenPage",
  "oddPage",
  "nextColumn",
];
const VALIGNS: readonly VerticalAlign[] = ["top", "center", "both", "bottom"];

function readColumns(sectPr: XmlElement | undefined, textWidth: number): ColumnSpec[] {
  const cols = child(sectPr, "cols");
  const count = Math.max(1, Math.floor(num(cols, "num") ?? 1));
  const space = num(cols, "space") ?? DEFAULT_COLUMN_SPACE;
  const explicit = cols ? childElementsOf(cols).filter((c) => c.name.local === "col") : [];
  // Individually listed columns apply unless equalWidth is switched on.
  const equal = explicit.length === 0 || onOffAttr(cols, "equalWidth");
  if (!equal) {
    return explicit.map((c, i) => ({
      width: num(c, "w") ?? 0,
      space: i === explicit.length - 1 ? 0 : (num(c, "space") ?? 0),
    }));
  }
  const width = Math.max(0, (textWidth - space * (count - 1)) / count);
  return Array.from({ length: count }, (_, i) => ({ width, space: i === count - 1 ? 0 : space }));
}

function readBorder(el: XmlElement | undefined): BorderLine | undefined {
  if (!el) return undefined;
  const style = getElementAttr(el, "val") ?? "none";
  if (style === "none" || style === "nil") return undefined;
  return {
    style,
    size: num(el, "sz") ?? 4,
    space: num(el, "space") ?? 0,
    color: getElementAttr(el, "color") ?? "auto",
  };
}

function readBorders(sectPr: XmlElement | undefined): PageBorders | undefined {
  const el = child(sectPr, "pgBorders");
  if (!el) return undefined;
  const borders: PageBorders = {
    offsetFrom: oneOf(getElementAttr(el, "offsetFrom"), ["text", "page"], "text"),
    display: oneOf(
      getElementAttr(el, "display"),
      ["allPages", "firstPage", "notFirstPage"],
      "allPages",
    ),
    zOrder: oneOf(getElementAttr(el, "zOrder"), ["front", "back"], "front"),
    top: readBorder(child(el, "top")),
    left: readBorder(child(el, "left")),
    bottom: readBorder(child(el, "bottom")),
    right: readBorder(child(el, "right")),
  };
  return borders.top || borders.left || borders.bottom || borders.right ? borders : undefined;
}

function readLineNumbers(sectPr: XmlElement | undefined): LineNumbering | undefined {
  const el = child(sectPr, "lnNumType");
  const countBy = num(el, "countBy");
  // No countBy (or 0) means line numbering is off (§17.6.8).
  if (!el || !countBy) return undefined;
  return {
    countBy,
    // `w:start` is the number minus one: start="0" numbers from 1.
    start: (num(el, "start") ?? 0) + 1,
    restart: oneOf(
      getElementAttr(el, "restart"),
      ["newPage", "newSection", "continuous"],
      "newPage",
    ),
    distance: num(el, "distance"),
  };
}

export function readNoteProperties(el: XmlElement | undefined): NoteProperties | undefined {
  if (!el) return undefined;
  const val = (local: string): string | undefined => {
    const c = child(el, local);
    return c && getElementAttr(c, "val");
  };
  const start = val("numStart");
  const restart = val("numRestart");
  return {
    numFmt: val("numFmt"),
    numStart: start === undefined ? undefined : Number(start),
    numRestart:
      restart === "continuous" || restart === "eachSect" || restart === "eachPage"
        ? restart
        : undefined,
    pos: val("pos"),
  };
}

function readGrid(sectPr: XmlElement | undefined): DocumentGrid | undefined {
  const el = child(sectPr, "docGrid");
  if (!el) return undefined;
  return {
    type: oneOf(
      getElementAttr(el, "type"),
      ["default", "lines", "linesAndChars", "snapToChars"],
      "default",
    ),
    linePitch: num(el, "linePitch"),
  };
}

function readSection(
  doc: Docx,
  sectPr: XmlElement | undefined,
  index: number,
  firstBlock: number,
  lastBlock: number,
): SectionModel {
  const pgSz = child(sectPr, "pgSz");
  const pgMar = child(sectPr, "pgMar");
  const pageWidth = num(pgSz, "w") ?? DEFAULT_PAGE.width;
  const pageHeight = num(pgSz, "h") ?? DEFAULT_PAGE.height;
  const top = num(pgMar, "top") ?? DEFAULT_MARGINS.top;
  const bottom = num(pgMar, "bottom") ?? DEFAULT_MARGINS.bottom;
  const margins: Margins = {
    top: Math.abs(top),
    bottom: Math.abs(bottom),
    left: num(pgMar, "left") ?? DEFAULT_MARGINS.left,
    right: num(pgMar, "right") ?? DEFAULT_MARGINS.right,
    header: num(pgMar, "header") ?? DEFAULT_MARGINS.header,
    footer: num(pgMar, "footer") ?? DEFAULT_MARGINS.footer,
    gutter: num(pgMar, "gutter") ?? DEFAULT_MARGINS.gutter,
    fixedTop: top < 0,
    fixedBottom: bottom < 0,
  };
  const textWidth = pageWidth - margins.left - margins.right - margins.gutter;
  const typeEl = child(sectPr, "type");
  const vAlignEl = child(sectPr, "vAlign");
  const pgNumType = child(sectPr, "pgNumType");
  const part = (kind: "header" | "footer", type: HeaderFooterType): string | undefined =>
    resolveHeaderFooter(doc, index, kind, type)?.partName;
  const titlePage = flag(sectPr, "titlePg");
  return {
    index,
    firstBlock,
    lastBlock,
    // The first section always starts a page; its type only matters for later ones.
    start: oneOf(typeEl && getElementAttr(typeEl, "val"), STARTS, "nextPage"),
    pageWidth,
    pageHeight,
    margins,
    columns: readColumns(sectPr, textWidth),
    separator: onOffAttr(child(sectPr, "cols"), "sep"),
    vAlign: oneOf(vAlignEl && getElementAttr(vAlignEl, "val"), VALIGNS, "top"),
    titlePage,
    pageNumbers: {
      format: pgNumType && getElementAttr(pgNumType, "fmt"),
      start: num(pgNumType, "start"),
    },
    lineNumbers: readLineNumbers(sectPr),
    borders: readBorders(sectPr),
    grid: readGrid(sectPr),
    footnotePr: readNoteProperties(child(sectPr, "footnotePr")),
    endnotePr: readNoteProperties(child(sectPr, "endnotePr")),
    headers: {
      default: part("header", "default"),
      first: part("header", "first"),
      even: part("header", "even"),
    },
    footers: {
      default: part("footer", "default"),
      first: part("footer", "first"),
      even: part("footer", "even"),
    },
  };
}

function isSectionBreak(block: Docx["document"]["body"]["blocks"][number]): boolean {
  return (
    block.kind === "paragraph" &&
    !!block.pPr &&
    childElementsOf(block.pPr).some((c) => c.name.local === "sectPr")
  );
}

/** Every section of the document, in order. There is always at least one. */
export function documentSections(doc: Docx): SectionModel[] {
  const all = sectionProperties(doc);
  const blocks = doc.document.body.blocks;
  const out: SectionModel[] = [];
  let first = 0;
  let index = 0;
  blocks.forEach((block, i) => {
    if (!isSectionBreak(block)) return;
    out.push(readSection(doc, all[index], index, first, i));
    index++;
    first = i + 1;
  });
  out.push(readSection(doc, all[index], index, first, blocks.length - 1));
  return out;
}

/** Document-wide settings the page layout reads. */
export interface LayoutSettings {
  readonly evenAndOddHeaders: boolean;
  /** Page colour (`w:background/@w:color`), when the settings show it. */
  readonly pageColor?: string | undefined;
}

const HEX_COLOR = /^[0-9A-Fa-f]{6}$/;

function settingOn(doc: Docx, local: string): boolean {
  const setting = getDocumentSetting(doc, local);
  return setting.present && !OFF.has(setting.val ?? "1");
}

export function layoutSettings(doc: Docx): LayoutSettings {
  // `w:background` is a child of <w:document> before <w:body> (§17.2.1).
  const background = doc.document.extras
    .map((e) => e.node)
    .find((n): n is XmlElement => n.kind === "element" && n.name.local === "background");
  const color = background && getElementAttr(background, "color");
  return {
    evenAndOddHeaders: settingOn(doc, "evenAndOddHeaders"),
    // Word shows the page colour only with displayBackgroundShape (§17.15.1.32).
    pageColor:
      color && HEX_COLOR.test(color) && settingOn(doc, "displayBackgroundShape")
        ? color
        : undefined,
  };
}
