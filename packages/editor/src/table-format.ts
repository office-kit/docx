/**
 * Effective table formatting, as Word lays a table out: the table style
 * hierarchy with its conditional formatting (ECMA-376 §17.7.6), the table's
 * own properties, and each cell's direct formatting.
 *
 * For every cell the table style contributes, from weakest to strongest:
 * the whole table, banded columns, banded rows, the header row, the total
 * row, the first column, the last column, and the four corner cells — each
 * only when the table's Table Style Options (`w:tblLook`) turn it on. Band
 * sizes come from `w:tblStyleRowBandSize` / `w:tblStyleColBandSize`, and the
 * header / total rows and first / last columns are not banded. A region's
 * borders apply to the region's outline (`top`/`left`/`bottom`/`right`) and
 * to the edges between its cells (`insideH`/`insideV`). The cell's own
 * `w:tcPr` wins over all of it.
 */

import {
  childElementsOf,
  type Docx,
  getElementAttr,
  getTableLook,
  stylesPart,
  type TableCellPlacement,
  tableCellPlacements,
  tableColumnWidths,
  type WmlTable,
  type XmlElement,
  getRawPartRoot,
  xmlPartNames,
} from "@office-kit/docx";
import type { CellTextFormat } from "./resolve.js";

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
export type CellEdge = "top" | "left" | "bottom" | "right" | "tl2br" | "tr2bl";
export type CellMarginSide = "top" | "left" | "bottom" | "right";

export interface ResolvedCell {
  readonly placement: TableCellPlacement;
  readonly borders: Readonly<Partial<Record<CellEdge, BorderSpec>>>;
  /** CSS-ready background color (`#RRGGBB`), or undefined for none. */
  readonly background?: string | undefined;
  /** Cell margins in twips. */
  readonly margins: Readonly<Record<CellMarginSide, number>>;
  /** `w:vAlign`: `top` / `center` / `bottom`. */
  readonly verticalAlign?: string | undefined;
  /** `w:textDirection` (ST_TextDirection). */
  readonly textDirection?: string | undefined;
  readonly noWrap: boolean;
  readonly text: CellTextFormat;
}

export interface ResolvedRow {
  /** `w:trHeight` in twips and its rule, if any. */
  readonly height?: { readonly value: number; readonly rule: string } | undefined;
  readonly hidden: boolean;
  /** Repeated at the top of each page (`w:tblHeader`). */
  readonly header: boolean;
  readonly gridBefore: number;
  readonly gridAfter: number;
}

export interface ResolvedTableFormat {
  /** Preferred table width (`w:tblW`); `undefined` lets the content decide. */
  readonly width?: { readonly type: string; readonly value: number } | undefined;
  /** Grid column widths in twips (`w:tblGrid`). */
  readonly columns: readonly number[];
  /** The table-level borders after the style and the table's own `w:tblBorders`. */
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
  /** `w:jc`: `left` / `center` / `right` (and their bidi names). */
  readonly alignment?: string | undefined;
  /** `w:tblCellSpacing` in twips (0 when cells touch). */
  readonly cellSpacing: number;
  readonly layout: "fixed" | "autofit";
  /** `[row][cell]`, parallel to `table.rows[r].cells`. */
  readonly cells: readonly (readonly ResolvedCell[])[];
  readonly rows: readonly ResolvedRow[];
}

const BORDER_SIDES: ReadonlySet<string> = new Set([
  "top",
  "left",
  "bottom",
  "right",
  "insideH",
  "insideV",
  "tl2br",
  "tr2bl",
]);
// `start` / `end` are the bidi-aware names for `left` / `right`.
const SIDE_ALIASES: Readonly<Record<string, string>> = { start: "left", end: "right" };
const HEX = /^[0-9A-Fa-f]{6}$/;
// ST_OnOff false values (§17.17.4).
const OFF: ReadonlySet<string> = new Set(["0", "false", "off"]);
const WORD_2013_MODE = 15;

function child(el: XmlElement | undefined, local: string): XmlElement | undefined {
  return el && childElementsOf(el).find((c) => c.name.local === local);
}

function intAttr(el: XmlElement | undefined, local: string): number | undefined {
  const raw = el && getElementAttr(el, local);
  const n = raw === undefined ? Number.NaN : Number(raw);
  return Number.isFinite(n) ? Math.trunc(n) : undefined;
}

function readBorders(container: XmlElement | undefined): Partial<Record<string, BorderSpec>> {
  const out: Partial<Record<string, BorderSpec>> = {};
  if (!container) return out;
  for (const el of childElementsOf(container)) {
    const side = SIDE_ALIASES[el.name.local] ?? el.name.local;
    if (!BORDER_SIDES.has(side)) continue;
    out[side] = {
      style: getElementAttr(el, "val") ?? "none",
      size: intAttr(el, "sz") ?? 0,
      color: getElementAttr(el, "color") ?? "auto",
    };
  }
  return out;
}

function readMargins(container: XmlElement | undefined): Partial<Record<CellMarginSide, number>> {
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

// Percent shading patterns (ST_Shd pctN): N % of the pattern color over the fill.
const PCT_PATTERN = /^pct(\d+)$/;
// Hatched patterns are approximated by the share of the cell their lines cover.
const HATCH_SHARE: Readonly<Record<string, number>> = {
  horzStripe: 50,
  vertStripe: 50,
  reverseDiagStripe: 50,
  diagStripe: 50,
  horzCross: 50,
  diagCross: 50,
  thinHorzStripe: 25,
  thinVertStripe: 25,
  thinReverseDiagStripe: 25,
  thinDiagStripe: 25,
  thinHorzCross: 25,
  thinDiagCross: 25,
};
const FULL = 100;

function mix(fill: string, pattern: string, percent: number): string {
  const a = [0, 2, 4].map((i) => Number.parseInt(fill.slice(i, i + 2), 16));
  const b = [0, 2, 4].map((i) => Number.parseInt(pattern.slice(i, i + 2), 16));
  return `#${a
    .map((c, i) =>
      Math.round(c + ((b[i] ?? 0) - c) * (percent / FULL))
        .toString(16)
        .padStart(2, "0"),
    )
    .join("")}`;
}

/**
 * The visible color of a `<w:shd>` (§17.3.5): the fill, with the pattern
 * color laid over it at the pattern's density. `undefined` means no shading.
 */
export function shadingColor(shd: XmlElement | undefined): string | undefined {
  if (!shd) return undefined;
  const val = getElementAttr(shd, "val") ?? "clear";
  const fillAttr = getElementAttr(shd, "fill");
  const colorAttr = getElementAttr(shd, "color");
  const fill = fillAttr && HEX.test(fillAttr) ? fillAttr : undefined;
  // An automatic pattern color is black; an automatic fill is none (white paper).
  const pattern = colorAttr && HEX.test(colorAttr) ? colorAttr : "000000";
  if (val === "nil") return undefined;
  if (val === "clear") return fill && `#${fill}`;
  if (val === "solid") return `#${pattern}`;
  const pct = PCT_PATTERN.exec(val);
  const share = pct ? Number(pct[1]) : HATCH_SHARE[val];
  if (share === undefined) return fill && `#${fill}`;
  return mix(fill ?? "FFFFFF", pattern, share);
}

type Region =
  | "wholeTable"
  | "band1Vert"
  | "band2Vert"
  | "band1Horz"
  | "band2Horz"
  | "firstRow"
  | "lastRow"
  | "firstCol"
  | "lastCol"
  | "nwCell"
  | "neCell"
  | "swCell"
  | "seCell";

// The order in which conditional formats are layered (§17.7.6).
const REGION_ORDER: readonly Region[] = [
  "wholeTable",
  "band1Vert",
  "band2Vert",
  "band1Horz",
  "band2Horz",
  "firstRow",
  "lastRow",
  "firstCol",
  "lastCol",
  "nwCell",
  "neCell",
  "swCell",
  "seCell",
];

/** The part of a style one region contributes. */
interface RegionProps {
  readonly pPr?: XmlElement | undefined;
  readonly rPr?: XmlElement | undefined;
  readonly tblPr?: XmlElement | undefined;
  readonly tcPr?: XmlElement | undefined;
}

/** Row / column extent of a region a cell falls in, for outline vs inside edges. */
interface Extent {
  readonly top: number;
  readonly bottom: number;
  readonly left: number;
  readonly right: number;
}

interface Mutable {
  borders: Partial<Record<CellEdge, BorderSpec>>;
  background?: string | undefined;
  margins: Partial<Record<CellMarginSide, number>>;
  verticalAlign?: string | undefined;
  textDirection?: string | undefined;
  noWrap: boolean;
  pPr: XmlElement[];
  rPr: XmlElement[];
}

/**
 * Lay borders over a cell. With an extent (a region of the table), the
 * region's outline sides apply on its edge and `insideH` / `insideV` between
 * its cells; without one (the cell's own tcBorders), sides apply as named.
 */
function applyBorders(
  out: Mutable,
  borders: Partial<Record<string, BorderSpec>>,
  cell: Extent,
  extent: Extent | undefined,
): void {
  const pick = (outer: boolean, side: string, inside: string): BorderSpec | undefined =>
    !extent || outer ? borders[side] : borders[inside];
  const edges: Array<[CellEdge, BorderSpec | undefined]> = [
    ["top", pick(cell.top === extent?.top, "top", "insideH")],
    ["bottom", pick(cell.bottom === extent?.bottom, "bottom", "insideH")],
    ["left", pick(cell.left === extent?.left, "left", "insideV")],
    ["right", pick(cell.right === extent?.right, "right", "insideV")],
    ["tl2br", borders.tl2br],
    ["tr2bl", borders.tr2bl],
  ];
  for (const [edge, spec] of edges) if (spec) out.borders[edge] = spec;
}

/** Lay a region's (or the cell's own) cell properties other than borders over a cell. */
function applyTcPr(out: Mutable, tcPr: XmlElement | undefined): void {
  if (!tcPr) return;
  const shd = child(tcPr, "shd");
  if (shd) out.background = shadingColor(shd);
  Object.assign(out.margins, readMargins(child(tcPr, "tcMar")));
  const vAlign = child(tcPr, "vAlign");
  if (vAlign) out.verticalAlign = getElementAttr(vAlign, "val");
  const dir = child(tcPr, "textDirection");
  if (dir) out.textDirection = getElementAttr(dir, "val");
  const noWrap = child(tcPr, "noWrap");
  if (noWrap) out.noWrap = !OFF.has(getElementAttr(noWrap, "val") ?? "");
}

/** The table style chain (default table style first), resolved through `basedOn`. */
function styleChain(doc: Docx, table: WmlTable): XmlElement[] {
  const styles = new Map<string, XmlElement>();
  let defaultId: string | undefined;
  for (const style of stylesPart(doc)?.styles ?? []) {
    const id = getElementAttr(style, "styleId");
    if (id === undefined || getElementAttr(style, "type") !== "table") continue;
    styles.set(id, style);
    if (getElementAttr(style, "default") === "1") defaultId = id;
  }
  const chainOf = (id: string | undefined): XmlElement[] => {
    const out: XmlElement[] = [];
    const seen = new Set<string>();
    let current = id;
    // A basedOn cycle is invalid but occurs in the wild; stop at the repeat.
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
  const ref = child(table.tblPr, "tblStyle");
  const id = ref && getElementAttr(ref, "val");
  // A missing style id falls back to the default table style, as in Word.
  const own = id !== undefined && styles.has(id) ? chainOf(id) : [];
  const base = chainOf(defaultId);
  return own.length && own[0] === base[0] ? own : [...base, ...own];
}

function regionProps(style: XmlElement, region: Region): RegionProps | undefined {
  if (region === "wholeTable") {
    const extra = childElementsOf(style).find(
      (c) => c.name.local === "tblStylePr" && getElementAttr(c, "type") === "wholeTable",
    );
    return {
      pPr: child(extra, "pPr") ?? child(style, "pPr"),
      rPr: child(extra, "rPr") ?? child(style, "rPr"),
      tblPr: child(extra, "tblPr"),
      tcPr: child(extra, "tcPr") ?? child(style, "tcPr"),
    };
  }
  const el = childElementsOf(style).find(
    (c) => c.name.local === "tblStylePr" && getElementAttr(c, "type") === region,
  );
  return el
    ? {
        pPr: child(el, "pPr"),
        rPr: child(el, "rPr"),
        tblPr: child(el, "tblPr"),
        tcPr: child(el, "tcPr"),
      }
    : undefined;
}

/** Resolve a table's layout and every cell's formatting. */
export function resolveTable(doc: Docx, table: WmlTable): ResolvedTableFormat {
  const chain = styleChain(doc, table);
  const tblPrs = [...chain.map((s) => child(s, "tblPr")), table.tblPr];
  const borders: Partial<Record<string, BorderSpec>> = {};
  const cellMargins: Partial<Record<CellMarginSide, number>> = {};
  let indent = 0;
  let alignment: string | undefined;
  let cellSpacing = 0;
  let rowBand = 1;
  let colBand = 1;
  let tableShading: string | undefined;
  for (const tblPr of tblPrs) {
    Object.assign(borders, readBorders(child(tblPr, "tblBorders")));
    Object.assign(cellMargins, readMargins(child(tblPr, "tblCellMar")));
    const tblInd = child(tblPr, "tblInd");
    if (tblInd && (getElementAttr(tblInd, "type") ?? "dxa") === "dxa") {
      indent = intAttr(tblInd, "w") ?? indent;
    }
    const jc = child(tblPr, "jc");
    if (jc) alignment = getElementAttr(jc, "val");
    const spacing = child(tblPr, "tblCellSpacing");
    if (spacing) cellSpacing = intAttr(spacing, "w") ?? cellSpacing;
    rowBand = intAttr(child(tblPr, "tblStyleRowBandSize"), "val") ?? rowBand;
    colBand = intAttr(child(tblPr, "tblStyleColBandSize"), "val") ?? colBand;
    const shd = child(tblPr, "shd");
    if (shd) tableShading = shadingColor(shd);
  }
  rowBand = Math.max(1, rowBand);
  colBand = Math.max(1, colBand);
  const look = getTableLook(table);
  const placements = tableCellPlacements(table);
  const columns = tableColumnWidths(table);
  const lastRow = table.rows.length - 1;
  const gridCount = Math.max(
    columns.length,
    ...placements.map((row) => {
      const end = row[row.length - 1];
      return end ? end.gridStart + end.gridSpan : 0;
    }),
  );
  // Word formats every leading repeated-header row as the header row.
  let headerRows = look.headerRow ? 1 : 0;
  if (look.headerRow) {
    while (headerRows < table.rows.length && child(table.rows[headerRows]?.trPr, "tblHeader")) {
      headerRows++;
    }
  }
  const hasTotal = look.totalRow && table.rows.length > headerRows;

  const cells = placements.map((row) =>
    row.map((p): ResolvedCell => {
      const bottomRow = p.row + Math.max(1, p.rowSpan) - 1;
      const self: Extent = {
        top: p.row,
        bottom: bottomRow,
        left: p.gridStart,
        right: p.gridStart + p.gridSpan - 1,
      };
      const isFirstCol = look.firstColumn && p.cell === 0;
      const isLastCol = look.lastColumn && p.cell === row.length - 1;
      const inHeader = p.row < headerRows;
      const inTotal = hasTotal && bottomRow === lastRow;
      const regions = new Map<Region, Extent>();
      const whole: Extent = { top: 0, bottom: lastRow, left: 0, right: gridCount - 1 };
      regions.set("wholeTable", whole);
      if (look.bandedColumns && !isFirstCol && !isLastCol) {
        const firstBanded = look.firstColumn ? 1 : 0;
        const band = Math.floor((p.gridStart - firstBanded) / colBand);
        const left = firstBanded + band * colBand;
        regions.set(band % 2 === 0 ? "band1Vert" : "band2Vert", {
          top: 0,
          bottom: lastRow,
          left,
          right: left + colBand - 1,
        });
      }
      if (look.bandedRows && !inHeader && !inTotal) {
        const band = Math.floor((p.row - headerRows) / rowBand);
        const top = headerRows + band * rowBand;
        regions.set(band % 2 === 0 ? "band1Horz" : "band2Horz", {
          top,
          bottom: top + rowBand - 1,
          left: 0,
          right: gridCount - 1,
        });
      }
      if (inHeader) regions.set("firstRow", { ...whole, bottom: headerRows - 1 });
      if (inTotal) regions.set("lastRow", { ...whole, top: lastRow });
      if (isFirstCol) regions.set("firstCol", { ...whole, left: self.left, right: self.right });
      if (isLastCol) regions.set("lastCol", { ...whole, left: self.left, right: self.right });
      if (inHeader && isFirstCol) regions.set("nwCell", self);
      if (inHeader && isLastCol) regions.set("neCell", self);
      if (inTotal && isFirstCol) regions.set("swCell", self);
      if (inTotal && isLastCol) regions.set("seCell", self);

      const out: Mutable = {
        borders: {},
        margins: { ...cellMargins },
        noWrap: false,
        pPr: [],
        rPr: [],
        background: tableShading,
      };
      // The table-level borders: outline on the table's edge, inside elsewhere.
      applyBorders(out, borders, self, whole);
      for (const region of REGION_ORDER) {
        const extent = regions.get(region);
        if (!extent) continue;
        for (const style of chain) {
          const props = regionProps(style, region);
          if (!props) continue;
          if (props.pPr) out.pPr.push(props.pPr);
          if (props.rPr) out.rPr.push(props.rPr);
          applyBorders(out, readBorders(child(props.tblPr, "tblBorders")), self, extent);
          applyBorders(out, readBorders(child(props.tcPr, "tcBorders")), self, extent);
          applyTcPr(out, props.tcPr);
        }
      }
      const cell = table.rows[p.row]?.cells[p.cell];
      applyBorders(out, readBorders(child(cell?.tcPr, "tcBorders")), self, undefined);
      applyTcPr(out, cell?.tcPr);
      return {
        placement: p,
        borders: out.borders,
        background: out.background,
        margins: {
          top: out.margins.top ?? 0,
          left: out.margins.left ?? 0,
          bottom: out.margins.bottom ?? 0,
          right: out.margins.right ?? 0,
        },
        verticalAlign: out.verticalAlign,
        textDirection: out.textDirection,
        noWrap: out.noWrap,
        text: { pPr: out.pPr, rPr: out.rPr },
      };
    }),
  );

  const rows = table.rows.map((row): ResolvedRow => {
    const trHeight = child(row.trPr, "trHeight");
    const value = intAttr(trHeight, "val");
    const hidden = child(row.trPr, "hidden");
    return {
      height:
        value !== undefined && value > 0
          ? { value, rule: (trHeight && getElementAttr(trHeight, "hRule")) ?? "atLeast" }
          : undefined,
      hidden: !!hidden && getElementAttr(hidden, "val") !== "0",
      header: !!child(row.trPr, "tblHeader"),
      gridBefore: intAttr(child(row.trPr, "gridBefore"), "val") ?? 0,
      gridAfter: intAttr(child(row.trPr, "gridAfter"), "val") ?? 0,
    };
  });

  const modern = compatibilityMode(doc) >= WORD_2013_MODE;
  const tblW = child(table.tblPr, "tblW");
  const widthValue = intAttr(tblW, "w");
  const tblLayout = child(table.tblPr, "tblLayout");
  return {
    width:
      tblW && widthValue !== undefined
        ? { type: getElementAttr(tblW, "type") ?? "dxa", value: widthValue }
        : undefined,
    columns,
    borders: pickSides(borders),
    cellMargins,
    leftEdge: modern ? indent : indent - (cellMargins.left ?? 0),
    alignment,
    cellSpacing,
    layout: tblLayout && getElementAttr(tblLayout, "type") === "fixed" ? "fixed" : "autofit",
    cells,
    rows,
  };
}

function pickSides(
  borders: Partial<Record<string, BorderSpec>>,
): Partial<Record<BorderSide, BorderSpec>> {
  const out: Partial<Record<BorderSide, BorderSpec>> = {};
  for (const side of ["top", "left", "bottom", "right", "insideH", "insideV"] as const) {
    const spec = borders[side];
    if (spec) out[side] = spec;
  }
  return out;
}

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
