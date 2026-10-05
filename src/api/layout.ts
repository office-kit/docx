/**
 * Section layout: the page setup, columns, line numbering, page borders,
 * vertical alignment and document grid of each section (`<w:sectPr>`,
 * ECMA-376 Part 1 §17.6), and inserting section breaks.
 *
 * Page size, margins and orientation have their own setters
 * (`setPageSize` / `setPageMargins` / `setPageOrientation`); everything else
 * in a section goes through {@link setSectionProperties}.
 */

import {
  MARGINS_NORMAL,
  PAGE_SIZE_LETTER,
  type PageMargins,
  type PageSize,
  type WmlParagraph,
} from "../internal/wordprocessingml/index.js";
import {
  removeOrderedChild,
  SECT_PR_ORDER,
  setOrderedOnOff,
  setOrderedVal,
  upsertOrderedChild,
} from "../internal/wordprocessingml/schema-order.js";
import {
  paragraphSectPr,
  resolveSectionScope,
  sectionIndexOfBlock,
  sectionPropertiesList,
  type SectionScope,
} from "../internal/wordprocessingml/sections.js";
import {
  attrOf,
  intAttrOf,
  isOn,
  onOffChild,
  wmlChild,
  wmlChildren,
  wmlElement,
} from "../internal/wordprocessingml/wml-element.js";
import type { XmlElement } from "../internal/xml/index.js";
import type { Docx } from "./docx.js";

export type { SectionScope };

/** How a section starts relative to the previous one (`w:type`, §17.18.77). */
export type SectionStart = "continuous" | "nextColumn" | "nextPage" | "evenPage" | "oddPage";

/** Vertical alignment of text on the page (`w:vAlign`, §17.18.101). */
export type VerticalAlignment = "top" | "center" | "both" | "bottom";

/** Section text flow (`w:textDirection`, §17.18.93). */
export type SectionTextDirection = "lrTb" | "tbRl" | "btLr" | "lrTbV" | "tbRlV" | "tbLrV";

/** One column of an unequal-width layout (`w:col`). */
export interface ColumnDefinition {
  readonly widthTwips: number;
  /** Space after this column; ignored for the last one. */
  readonly spaceTwips?: number;
}

/** The section's text columns (`w:cols`, §17.6.4). */
export interface SectionColumns {
  /** Number of columns (1–45). */
  readonly count: number;
  /** Space between equal-width columns, in twips. Word's default is 720 (0.5 in). */
  readonly spaceTwips: number;
  /** Draw a vertical line between columns (`w:sep`). */
  readonly separator: boolean;
  /**
   * Per-column widths. Present means unequal widths (`w:equalWidth="0"`);
   * its length must equal `count`.
   */
  readonly columns?: readonly ColumnDefinition[];
}

/** When line numbering restarts (`w:restart`, §17.18.47). */
export type LineNumberRestart = "newPage" | "newSection" | "continuous";

/** Line numbering (`w:lnNumType`, §17.6.8). */
export interface LineNumbering {
  /** The number shown on the first line (Word's "Start at"). */
  readonly start: number;
  /** Show every n-th number (Word's "Count by"). */
  readonly countBy: number;
  /** Distance from the text in twips; `undefined` is Word's "Auto". */
  readonly distanceTwips?: number;
  readonly restart: LineNumberRestart;
}

/** One edge of a page border. */
export interface PageBorder {
  /** A line style from {@link PAGE_BORDER_LINE_STYLES} or an art border from {@link PAGE_BORDER_ART}. */
  readonly style: string;
  /**
   * Width in eighths of a point for line styles (2–96); for art borders, the
   * art's size in points (1–31), as §17.3.4 defines `w:sz` for them.
   */
  readonly size: number;
  /** Distance from the text or page edge in points (0–31). */
  readonly spacePt: number;
  /** Hex RGB or `"auto"`. */
  readonly color: string;
  /** Theme color the RGB value came from, if any (`w:themeColor`). */
  readonly themeColor?: string;
  readonly shadow?: boolean;
  /** 3-D effect (`w:frame`). */
  readonly frame?: boolean;
}

/** Page borders (`w:pgBorders`, §17.6.10). */
export interface PageBorders {
  /** Measure the border distance from the page edge or from the text. */
  readonly offsetFrom: "page" | "text";
  /** Which pages show the border. */
  readonly display: "allPages" | "firstPage" | "notFirstPage";
  /** Draw in front of (`front`) or behind (`back`) intersecting text. */
  readonly zOrder: "front" | "back";
  readonly top?: PageBorder;
  readonly left?: PageBorder;
  readonly bottom?: PageBorder;
  readonly right?: PageBorder;
}

/** Document grid (`w:docGrid`, §17.6.5). */
export interface DocumentGrid {
  readonly type: "default" | "lines" | "linesAndChars" | "snapToChars";
  /** Line pitch in twips. */
  readonly linePitch?: number;
  /**
   * Character pitch as a difference, in 4096ths of a point, from the width
   * of a character of the Normal style's font size.
   */
  readonly charSpace?: number;
}

/** A section's layout as read from its `<w:sectPr>`, with Word's defaults filled in. */
export interface SectionProperties {
  readonly start: SectionStart;
  readonly pageSize: Required<Omit<PageSize, "paperCode">> & Pick<PageSize, "paperCode">;
  readonly margins: PageMargins;
  readonly columns: SectionColumns;
  readonly lineNumbering?: LineNumbering;
  readonly pageBorders?: PageBorders;
  readonly verticalAlignment: VerticalAlignment;
  /** A different first-page header and footer (`w:titlePg`). */
  readonly titlePage: boolean;
  readonly textDirection: SectionTextDirection;
  /** Right-to-left section layout (`w:bidi`). */
  readonly bidi: boolean;
  /** The gutter is on the right (`w:rtlGutter`). */
  readonly rtlGutter: boolean;
  readonly documentGrid?: DocumentGrid;
}

/**
 * Changes for {@link setSectionProperties}. Omitted fields are left alone;
 * `null` removes an optional element (line numbering, page borders, grid).
 */
export interface SectionPropertiesPatch {
  readonly start?: SectionStart;
  readonly columns?: SectionColumns;
  readonly lineNumbering?: LineNumbering | null;
  readonly pageBorders?: PageBorders | null;
  readonly verticalAlignment?: VerticalAlignment;
  readonly titlePage?: boolean;
  readonly textDirection?: SectionTextDirection;
  readonly bidi?: boolean;
  readonly rtlGutter?: boolean;
  readonly documentGrid?: DocumentGrid | null;
}

/** Line styles `w:pgBorders` accepts (the non-art members of ST_Border, §17.18.2). */
export const PAGE_BORDER_LINE_STYLES: readonly string[] = [
  "single",
  "thick",
  "double",
  "dotted",
  "dashed",
  "dotDash",
  "dotDotDash",
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
  "dashSmallGap",
  "dashDotStroked",
  "threeDEmboss",
  "threeDEngrave",
  "outset",
  "inset",
];

/** Art borders (the picture members of ST_Border, §17.18.2), valid only on page borders. */
export const PAGE_BORDER_ART: readonly string[] = [
  "apples",
  "archedScallops",
  "babyPacifier",
  "babyRattle",
  "balloons3Colors",
  "balloonsHotAir",
  "basicBlackDashes",
  "basicBlackDots",
  "basicBlackSquares",
  "basicThinLines",
  "basicWhiteDashes",
  "basicWhiteDots",
  "basicWhiteSquares",
  "basicWideInline",
  "basicWideMidline",
  "basicWideOutline",
  "bats",
  "birds",
  "birdsFlight",
  "cabins",
  "cakeSlice",
  "candyCorn",
  "celticKnotwork",
  "certificateBanner",
  "chainLink",
  "champagneBottle",
  "checkedBarBlack",
  "checkedBarColor",
  "checkered",
  "christmasTree",
  "circlesLines",
  "circlesRectangles",
  "classicalWave",
  "clocks",
  "compass",
  "confetti",
  "confettiGrays",
  "confettiOutline",
  "confettiStreamers",
  "confettiWhite",
  "cornerTriangles",
  "couponCutoutDashes",
  "couponCutoutDots",
  "crazyMaze",
  "creaturesButterfly",
  "creaturesFish",
  "creaturesInsects",
  "creaturesLadyBug",
  "crossStitch",
  "cup",
  "decoArch",
  "decoArchColor",
  "decoBlocks",
  "diamondsGray",
  "doubleD",
  "doubleDiamonds",
  "earth1",
  "earth2",
  "earth3",
  "eclipsingSquares1",
  "eclipsingSquares2",
  "eggsBlack",
  "fans",
  "film",
  "firecrackers",
  "flowersBlockPrint",
  "flowersDaisies",
  "flowersModern1",
  "flowersModern2",
  "flowersPansy",
  "flowersRedRose",
  "flowersRoses",
  "flowersTeacup",
  "flowersTiny",
  "gems",
  "gingerbreadMan",
  "gradient",
  "handmade1",
  "handmade2",
  "heartBalloon",
  "heartGray",
  "hearts",
  "heebieJeebies",
  "holly",
  "houseFunky",
  "hypnotic",
  "iceCreamCones",
  "lightBulb",
  "lightning1",
  "lightning2",
  "mapPins",
  "mapleLeaf",
  "mapleMuffins",
  "marquee",
  "marqueeToothed",
  "moons",
  "mosaic",
  "musicNotes",
  "northwest",
  "ovals",
  "packages",
  "palmsBlack",
  "palmsColor",
  "paperClips",
  "papyrus",
  "partyFavor",
  "partyGlass",
  "pencils",
  "people",
  "peopleWaving",
  "peopleHats",
  "poinsettias",
  "postageStamp",
  "pumpkin1",
  "pushPinNote2",
  "pushPinNote1",
  "pyramids",
  "pyramidsAbove",
  "quadrants",
  "rings",
  "safari",
  "sawtooth",
  "sawtoothGray",
  "scaredCat",
  "seattle",
  "shadowedSquares",
  "sharksTeeth",
  "shorebirdTracks",
  "skyrocket",
  "snowflakeFancy",
  "snowflakes",
  "sombrero",
  "southwest",
  "stars",
  "starsTop",
  "stars3d",
  "starsBlack",
  "starsShadowed",
  "sun",
  "swirligig",
  "tornPaper",
  "tornPaperBlack",
  "trees",
  "triangleParty",
  "triangles",
  "triangle1",
  "triangle2",
  "triangleCircle1",
  "triangleCircle2",
  "shapes1",
  "shapes2",
  "twistedLines1",
  "twistedLines2",
  "vine",
  "waveline",
  "weavingAngles",
  "weavingBraid",
  "weavingRibbon",
  "weavingStrips",
  "whiteFlowers",
  "woodwork",
  "xIllusions",
  "zanyTriangles",
  "zigZag",
  "zigZagStitch",
];

const BORDER_SIDES = ["top", "left", "bottom", "right"] as const;
const SECTION_STARTS: ReadonlySet<string> = new Set([
  "continuous",
  "nextColumn",
  "nextPage",
  "evenPage",
  "oddPage",
]);
const VERTICAL_ALIGNMENTS: ReadonlySet<string> = new Set(["top", "center", "both", "bottom"]);
const TEXT_DIRECTIONS: ReadonlySet<string> = new Set([
  "lrTb",
  "tbRl",
  "btLr",
  "lrTbV",
  "tbRlV",
  "tbLrV",
]);
const GRID_TYPES: ReadonlySet<string> = new Set([
  "default",
  "lines",
  "linesAndChars",
  "snapToChars",
]);
const LINE_NUMBER_RESTARTS: ReadonlySet<string> = new Set(["newPage", "newSection", "continuous"]);
// §17.6.4: Word's default spacing between columns is 0.5 in.
const DEFAULT_COLUMN_SPACE = 720;
const MAX_COLUMNS = 45;
// `w:space` on a page border is in points, 0–31 (§17.18.2 ST_PointMeasure range used by Word).
const MAX_BORDER_SPACE_PT = 31;

function oneOf<T extends string>(
  value: string | undefined,
  allowed: ReadonlySet<string>,
  fallback: T,
): T {
  return value !== undefined && allowed.has(value) ? (value as T) : fallback;
}

/** The number of sections in the document. */
export function sectionCount(doc: Docx): number {
  return sectionPropertiesList(doc.document).length;
}

/** The index of the section that holds body block `blockIndex`. */
export function sectionIndexAt(doc: Docx, blockIndex: number): number {
  return sectionIndexOfBlock(doc.document, blockIndex);
}

function readBorder(el: XmlElement | undefined): PageBorder | undefined {
  const style = attrOf(el, "val");
  if (!el || style === undefined || style === "none" || style === "nil") return undefined;
  const themeColor = attrOf(el, "themeColor");
  return {
    style,
    size: intAttrOf(el, "sz") ?? 4,
    spacePt: intAttrOf(el, "space") ?? 0,
    color: attrOf(el, "color") ?? "auto",
    ...(themeColor === undefined ? {} : { themeColor }),
    ...(isOn(attrOf(el, "shadow")) ? { shadow: true } : {}),
    ...(isOn(attrOf(el, "frame")) ? { frame: true } : {}),
  };
}

function readPageBorders(el: XmlElement | undefined): PageBorders | undefined {
  if (!el) return undefined;
  const sides: Partial<Record<(typeof BORDER_SIDES)[number], PageBorder>> = {};
  for (const side of BORDER_SIDES) {
    const border = readBorder(wmlChild(el, side));
    if (border) sides[side] = border;
  }
  return {
    offsetFrom: attrOf(el, "offsetFrom") === "page" ? "page" : "text",
    display: oneOf(
      attrOf(el, "display"),
      new Set(["allPages", "firstPage", "notFirstPage"]),
      "allPages",
    ),
    zOrder: attrOf(el, "zOrder") === "back" ? "back" : "front",
    ...sides,
  };
}

function readColumns(el: XmlElement | undefined): SectionColumns {
  const count = Math.max(1, intAttrOf(el, "num") ?? 1);
  const spaceTwips = intAttrOf(el, "space") ?? DEFAULT_COLUMN_SPACE;
  const separator = isOn(attrOf(el, "sep"));
  const equalWidth = attrOf(el, "equalWidth");
  const cols = wmlChildren(el, "col");
  // §17.6.4: equalWidth defaults to true, and individual widths only apply when it is off.
  if (equalWidth === undefined || isOn(equalWidth) || cols.length === 0) {
    return { count, spaceTwips, separator };
  }
  return {
    count: cols.length,
    spaceTwips,
    separator,
    columns: cols.map((c): ColumnDefinition => {
      const widthTwips = intAttrOf(c, "w") ?? 0;
      const spaceTwips = intAttrOf(c, "space");
      return spaceTwips === undefined ? { widthTwips } : { widthTwips, spaceTwips };
    }),
  };
}

function readLineNumbering(el: XmlElement | undefined): LineNumbering | undefined {
  if (!el) return undefined;
  const distance = intAttrOf(el, "distance");
  return {
    // Word stores "Start at n" as n - 1 (an absent w:start is "Start at 1").
    start: (intAttrOf(el, "start") ?? 0) + 1,
    countBy: intAttrOf(el, "countBy") ?? 1,
    ...(distance === undefined ? {} : { distanceTwips: distance }),
    restart: oneOf(attrOf(el, "restart"), LINE_NUMBER_RESTARTS, "newPage"),
  };
}

function readGrid(el: XmlElement | undefined): DocumentGrid | undefined {
  if (!el) return undefined;
  const linePitch = intAttrOf(el, "linePitch");
  const charSpace = intAttrOf(el, "charSpace");
  return {
    type: oneOf(attrOf(el, "type"), GRID_TYPES, "default"),
    ...(linePitch === undefined ? {} : { linePitch }),
    ...(charSpace === undefined ? {} : { charSpace }),
  };
}

/** Read a `<w:sectPr>` as typed properties. */
function sectionPropertiesOf(sectPr: XmlElement): SectionProperties {
  const pgSz = wmlChild(sectPr, "pgSz");
  const pgMar = wmlChild(sectPr, "pgMar");
  const width = intAttrOf(pgSz, "w") ?? PAGE_SIZE_LETTER.widthTwips;
  const height = intAttrOf(pgSz, "h") ?? PAGE_SIZE_LETTER.heightTwips;
  const paperCode = intAttrOf(pgSz, "code");
  const lineNumbering = readLineNumbering(wmlChild(sectPr, "lnNumType"));
  const pageBorders = readPageBorders(wmlChild(sectPr, "pgBorders"));
  const documentGrid = readGrid(wmlChild(sectPr, "docGrid"));
  const margin = (local: keyof PageMargins): number =>
    intAttrOf(pgMar, local) ?? (pgMar ? 0 : MARGINS_NORMAL[local]);
  return {
    start: oneOf(attrOf(wmlChild(sectPr, "type"), "val"), SECTION_STARTS, "nextPage"),
    pageSize: {
      widthTwips: width,
      heightTwips: height,
      orientation: attrOf(pgSz, "orient") === "landscape" ? "landscape" : "portrait",
      ...(paperCode === undefined ? {} : { paperCode }),
    },
    margins: {
      top: margin("top"),
      right: margin("right"),
      bottom: margin("bottom"),
      left: margin("left"),
      header: margin("header"),
      footer: margin("footer"),
      gutter: margin("gutter"),
    },
    columns: readColumns(wmlChild(sectPr, "cols")),
    ...(lineNumbering ? { lineNumbering } : {}),
    ...(pageBorders ? { pageBorders } : {}),
    verticalAlignment: oneOf(attrOf(wmlChild(sectPr, "vAlign"), "val"), VERTICAL_ALIGNMENTS, "top"),
    titlePage: onOffChild(sectPr, "titlePg"),
    textDirection: oneOf(attrOf(wmlChild(sectPr, "textDirection"), "val"), TEXT_DIRECTIONS, "lrTb"),
    bidi: onOffChild(sectPr, "bidi"),
    rtlGutter: onOffChild(sectPr, "rtlGutter"),
    ...(documentGrid ? { documentGrid } : {}),
  };
}

/** A section's layout (default: the last section). */
export function getSectionProperties(doc: Docx, section?: number): SectionProperties {
  const [sectPr] = resolveSectionScope(doc.document, section);
  if (!sectPr) throw new RangeError(`Section ${section} does not exist.`);
  return sectionPropertiesOf(sectPr);
}

function requireInt(value: number, min: number, max: number, what: string): number {
  if (!Number.isInteger(value) || value < min || value > max) {
    throw new RangeError(`${what} must be an integer in ${min}..${max}, got ${value}.`);
  }
  return value;
}

function buildColumns(columns: SectionColumns): XmlElement {
  const count = requireInt(columns.count, 1, MAX_COLUMNS, "Column count");
  const space = requireInt(columns.spaceTwips, 0, 31680, "Column spacing");
  const list = columns.columns;
  if (list && list.length !== count) {
    throw new RangeError(`columns lists ${list.length} widths for ${count} columns.`);
  }
  const children = (list ?? []).map((c, i) =>
    wmlElement("col", {
      w: requireInt(c.widthTwips, 1, 31680, "Column width"),
      // The last column's space is meaningless; Word omits it.
      space: i < count - 1 ? (c.spaceTwips ?? space) : undefined,
    }),
  );
  return wmlElement(
    "cols",
    {
      num: count === 1 && !list ? undefined : count,
      space,
      sep: columns.separator ? 1 : undefined,
      equalWidth: list ? 0 : undefined,
    },
    children,
  );
}

function buildLineNumbering(ln: LineNumbering): XmlElement {
  if (!LINE_NUMBER_RESTARTS.has(ln.restart)) throw new RangeError(`Unknown restart ${ln.restart}.`);
  const start = requireInt(ln.start, 1, 32767, "Line number start");
  return wmlElement("lnNumType", {
    countBy: requireInt(ln.countBy, 1, 100, "Count by"),
    start: start === 1 ? undefined : start - 1,
    distance: ln.distanceTwips,
    restart: ln.restart,
  });
}

function buildBorder(side: string, border: PageBorder): XmlElement {
  if (!PAGE_BORDER_LINE_STYLES.includes(border.style) && !PAGE_BORDER_ART.includes(border.style))
    throw new RangeError(`Unknown page border style ${border.style}.`);
  if (border.color !== "auto" && !/^[0-9A-Fa-f]{6}$/.test(border.color)) {
    throw new RangeError(`Border color must be hex RRGGBB or "auto", got ${border.color}.`);
  }
  return wmlElement(side, {
    val: border.style,
    color: border.color,
    themeColor: border.themeColor,
    sz: requireInt(border.size, 1, 96, "Border size"),
    space: requireInt(border.spacePt, 0, MAX_BORDER_SPACE_PT, "Border spacing"),
    shadow: border.shadow ? 1 : undefined,
    frame: border.frame ? 1 : undefined,
  });
}

function buildPageBorders(borders: PageBorders): XmlElement {
  const children = BORDER_SIDES.flatMap((side) => {
    const b = borders[side];
    return b ? [buildBorder(side, b)] : [];
  });
  return wmlElement(
    "pgBorders",
    {
      // Word writes only the non-default values (front / allPages / text).
      zOrder: borders.zOrder === "back" ? "back" : undefined,
      display: borders.display === "allPages" ? undefined : borders.display,
      offsetFrom: borders.offsetFrom === "page" ? "page" : undefined,
    },
    children,
  );
}

function buildGrid(grid: DocumentGrid): XmlElement {
  if (!GRID_TYPES.has(grid.type)) throw new RangeError(`Unknown document grid type ${grid.type}.`);
  return wmlElement("docGrid", {
    type: grid.type === "default" ? undefined : grid.type,
    linePitch: grid.linePitch,
    charSpace: grid.charSpace,
  });
}

function check(value: string, allowed: ReadonlySet<string>, what: string): string {
  if (!allowed.has(value)) throw new RangeError(`Unknown ${what} ${value}.`);
  return value;
}

/**
 * Change section properties of the sections `scope` names (default: the last
 * section). Each element is written at its `CT_SectPr` schema position.
 */
export function setSectionProperties(
  doc: Docx,
  patch: SectionPropertiesPatch,
  scope?: SectionScope,
): void {
  // Validate once up front so a bad value never leaves some sections changed.
  const cols = patch.columns && buildColumns(patch.columns);
  const lnNum = patch.lineNumbering && buildLineNumbering(patch.lineNumbering);
  const borders = patch.pageBorders && buildPageBorders(patch.pageBorders);
  const grid = patch.documentGrid && buildGrid(patch.documentGrid);
  if (patch.start) check(patch.start, SECTION_STARTS, "section start");
  if (patch.verticalAlignment)
    check(patch.verticalAlignment, VERTICAL_ALIGNMENTS, "vertical alignment");
  if (patch.textDirection) check(patch.textDirection, TEXT_DIRECTIONS, "text direction");
  for (const sectPr of resolveSectionScope(doc.document, scope)) {
    const put = (el: XmlElement): void =>
      upsertOrderedChild(sectPr, structuredClone(el), SECT_PR_ORDER);
    if (patch.start) {
      // nextPage is the default (§17.6.22); Word omits it.
      setOrderedVal(
        sectPr,
        "type",
        patch.start === "nextPage" ? undefined : patch.start,
        SECT_PR_ORDER,
      );
    }
    if (cols) put(cols);
    if (lnNum) put(lnNum);
    else if (patch.lineNumbering === null) removeOrderedChild(sectPr, "lnNumType");
    if (borders) put(borders);
    else if (patch.pageBorders === null) removeOrderedChild(sectPr, "pgBorders");
    if (grid) put(grid);
    else if (patch.documentGrid === null) removeOrderedChild(sectPr, "docGrid");
    if (patch.verticalAlignment) {
      const v = patch.verticalAlignment;
      setOrderedVal(sectPr, "vAlign", v === "top" ? undefined : v, SECT_PR_ORDER);
    }
    if (patch.textDirection) {
      const d = patch.textDirection;
      setOrderedVal(sectPr, "textDirection", d === "lrTb" ? undefined : d, SECT_PR_ORDER);
    }
    if (patch.titlePage !== undefined)
      setOrderedOnOff(sectPr, "titlePg", patch.titlePage, SECT_PR_ORDER);
    if (patch.bidi !== undefined) setOrderedOnOff(sectPr, "bidi", patch.bidi, SECT_PR_ORDER);
    if (patch.rtlGutter !== undefined) {
      setOrderedOnOff(sectPr, "rtlGutter", patch.rtlGutter, SECT_PR_ORDER);
    }
  }
  doc.dirty = true;
}

/**
 * Insert a section break after body block `blockIndex`, as Word's Layout ▸
 * Breaks ▸ Section Breaks does at the end of that paragraph. The section
 * that held the block is split in two: the first part keeps a copy of its
 * properties (headers, page setup, columns …), and the part after the break
 * starts as `start` says (`w:type` describes how *its own* section begins).
 *
 * The break goes on the block itself when it is a paragraph without a
 * section break; otherwise a new empty paragraph after it carries it.
 * Returns the index of the new section after the break.
 */
export function insertSectionBreak(
  doc: Docx,
  blockIndex: number,
  start: SectionStart = "nextPage",
): number {
  check(start, SECTION_STARTS, "section start");
  const blocks = doc.document.body.blocks;
  if (!Number.isInteger(blockIndex) || blockIndex < 0 || blockIndex >= blocks.length) {
    throw new RangeError(`Block ${blockIndex} does not exist.`);
  }
  const section = sectionIndexOfBlock(doc.document, blockIndex);
  const following = sectionPropertiesList(doc.document)[section];
  if (!following) throw new RangeError(`Section ${section} does not exist.`);
  let target = blocks[blockIndex];
  if (target?.kind !== "paragraph" || paragraphSectPr(target)) {
    target = { kind: "paragraph", children: [], extras: [] } satisfies WmlParagraph;
    blocks.splice(blockIndex + 1, 0, target);
  }
  const pPr = target.pPr ?? wmlElement("pPr");
  target.pPr = pPr;
  // sectPr is the last child of pPr (only pPrChange follows it, §17.3.1.26).
  const changeIndex = pPr.children.findIndex(
    (c) => c.kind === "element" && c.name.local === "pPrChange",
  );
  const copy = structuredClone(following);
  const children = pPr.children as XmlElement["children"][number][];
  if (changeIndex >= 0) children.splice(changeIndex, 0, copy);
  else children.push(copy);
  setOrderedVal(following, "type", start === "nextPage" ? undefined : start, SECT_PR_ORDER);
  doc.dirty = true;
  return section + 1;
}
