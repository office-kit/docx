/**
 * DrawingML charts (ECMA-376 Part 1, §21.2): a {@link ChartSpec} written as a
 * `c:chartSpace` part with value caches, and read back from any chart part
 * (Word's own included) as far as the spec can describe it.
 */

import { parseXml, type XmlElement } from "../xml/index.js";
import { readColor, srgbMarkup } from "./color.js";
import { type CellValue, columnName, SHEET_NAME } from "./xlsx.js";
import { A_NS, attr, C_NS, child, descendant, elementChildren, escapeXml, path } from "./xml.js";

/** Insert Chart's chart families (the ECMA-376 `c:` chart types). */
export type ChartKind =
  | "column"
  | "bar"
  | "line"
  | "area"
  | "pie"
  | "doughnut"
  | "scatter"
  | "radar";

/** Clustered, Stacked or 100% Stacked (column / bar / line / area). */
export type ChartGrouping = "clustered" | "stacked" | "percentStacked";

export type LegendPosition = "right" | "left" | "top" | "bottom" | "none";

export interface ChartSeries {
  readonly name: string;
  /** One value per category; `null` leaves a gap. */
  readonly values: ReadonlyArray<number | null>;
  /** `RRGGBB`; omitted means the chart style's color for this series. */
  readonly color?: string;
}

/**
 * A chart's data and the Chart Elements Word's ribbon edits. For scatter
 * charts the categories are the X values.
 */
export interface ChartSpec {
  readonly kind: ChartKind;
  readonly grouping?: ChartGrouping;
  /** Line / scatter / radar: draw point markers. */
  readonly markers?: boolean;
  readonly title?: string;
  readonly legend: LegendPosition;
  readonly dataLabels: boolean;
  readonly categories: readonly string[];
  readonly series: readonly ChartSeries[];
}

const LEGEND_CODE: Readonly<Record<Exclude<LegendPosition, "none">, string>> = {
  right: "r",
  left: "l",
  top: "t",
  bottom: "b",
};
const LEGEND_FROM_CODE: Readonly<Record<string, LegendPosition>> = {
  r: "right",
  l: "left",
  t: "top",
  b: "bottom",
  tr: "right",
};

/** The Office theme's accent colors: the default series colors, in order. */
export const DEFAULT_SERIES_COLORS = [
  "4472C4",
  "ED7D31",
  "A5A5A5",
  "FFC000",
  "5B9BD5",
  "70AD47",
] as const;

const CAT_AX_ID = 111111111;
const VAL_AX_ID = 222222222;

function absRef(col: number, row: number): string {
  return `$${columnName(col)}$${row}`;
}

function strCache(values: readonly string[]): string {
  const pts = values.map((v, i) => `<c:pt idx="${i}"><c:v>${escapeXml(v)}</c:v></c:pt>`).join("");
  return `<c:strCache><c:ptCount val="${values.length}"/>${pts}</c:strCache>`;
}

function numCache(values: ReadonlyArray<number | null>): string {
  const pts = values
    .map((v, i) =>
      v === null || !Number.isFinite(v) ? "" : `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`,
    )
    .join("");
  return `<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>${pts}</c:numCache>`;
}

function seriesColor(spec: ChartSpec, i: number): string {
  return (
    spec.series[i]?.color ?? DEFAULT_SERIES_COLORS[i % DEFAULT_SERIES_COLORS.length] ?? "4472C4"
  );
}

const LINE_WIDTH_EMU = 28575;

function seriesMarkup(spec: ChartSpec, i: number): string {
  const s = spec.series[i];
  if (!s) return "";
  const n = spec.categories.length;
  const col = i + 1;
  const tx = `<c:tx><c:strRef><c:f>${SHEET_NAME}!${absRef(col, 1)}</c:f>${strCache([s.name])}</c:strRef></c:tx>`;
  const color = srgbMarkup(seriesColor(spec, i));
  const valRef = `${SHEET_NAME}!${absRef(col, 2)}:${absRef(col, n + 1)}`;
  const catRef = `${SHEET_NAME}!${absRef(0, 2)}:${absRef(0, n + 1)}`;
  const cat = `<c:cat><c:strRef><c:f>${catRef}</c:f>${strCache(spec.categories)}</c:strRef></c:cat>`;
  const val = `<c:val><c:numRef><c:f>${valRef}</c:f>${numCache(s.values)}</c:numRef></c:val>`;
  const head = `<c:idx val="${i}"/><c:order val="${i}"/>${tx}`;
  const lineSpPr = `<c:spPr><a:ln w="${LINE_WIDTH_EMU}" cap="rnd"><a:solidFill>${color}</a:solidFill><a:round/></a:ln></c:spPr>`;
  const marker = spec.markers
    ? `<c:marker><c:symbol val="circle"/><c:size val="5"/><c:spPr><a:solidFill>${color}</a:solidFill></c:spPr></c:marker>`
    : `<c:marker><c:symbol val="none"/></c:marker>`;
  switch (spec.kind) {
    case "column":
    case "bar":
      return `<c:ser>${head}<c:spPr><a:solidFill>${color}</a:solidFill></c:spPr><c:invertIfNegative val="0"/>${cat}${val}</c:ser>`;
    case "area":
      return `<c:ser>${head}<c:spPr><a:solidFill>${color}</a:solidFill></c:spPr>${cat}${val}</c:ser>`;
    case "line":
      return `<c:ser>${head}${lineSpPr}${marker}${cat}${val}<c:smooth val="0"/></c:ser>`;
    case "radar":
      return `<c:ser>${head}${lineSpPr}${marker}${cat}${val}</c:ser>`;
    case "pie":
    case "doughnut": {
      // Each slice is a data point with its own fill (CT_DPt).
      const points = spec.categories
        .map((_, p) => {
          const fill = srgbMarkup(
            DEFAULT_SERIES_COLORS[p % DEFAULT_SERIES_COLORS.length] ?? "4472C4",
          );
          return `<c:dPt><c:idx val="${p}"/><c:bubble3D val="0"/><c:spPr><a:solidFill>${fill}</a:solidFill><a:ln><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></c:spPr></c:dPt>`;
        })
        .join("");
      return `<c:ser>${head}${points}${cat}${val}</c:ser>`;
    }
    case "scatter": {
      const xNums = spec.categories.map((c) => {
        const v = Number(c);
        return Number.isFinite(v) ? v : null;
      });
      const xVal = `<c:xVal><c:numRef><c:f>${catRef}</c:f>${numCache(xNums)}</c:numRef></c:xVal>`;
      const yVal = `<c:yVal><c:numRef><c:f>${valRef}</c:f>${numCache(s.values)}</c:numRef></c:yVal>`;
      const spPr = spec.markers
        ? `<c:spPr><a:ln w="${LINE_WIDTH_EMU}"><a:noFill/></a:ln></c:spPr>`
        : lineSpPr;
      const scatterMarker = `<c:marker><c:symbol val="circle"/><c:size val="5"/><c:spPr><a:solidFill>${color}</a:solidFill></c:spPr></c:marker>`;
      return `<c:ser>${head}${spPr}${scatterMarker}${xVal}${yVal}<c:smooth val="0"/></c:ser>`;
    }
    default:
      return "";
  }
}

const DATA_LABELS =
  `<c:dLbls><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/>` +
  `<c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>`;

function axisHead(pos: string): string {
  return `<c:delete val="0"/><c:axPos val="${pos}"/>`;
}

function axisTail(crossAx: number): string {
  return `<c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="${crossAx}"/><c:crosses val="autoZero"/>`;
}

function axesMarkup(spec: ChartSpec): string {
  const horizontalBars = spec.kind === "bar";
  const percent = spec.grouping === "percentStacked";
  const scaling = `<c:scaling><c:orientation val="minMax"/></c:scaling>`;
  const between = spec.kind === "area" || spec.kind === "scatter" ? "midCat" : "between";
  const valAx = (id: number, pos: string, crossAx: number, gridlines: boolean): string =>
    `<c:valAx><c:axId val="${id}"/>${scaling}${axisHead(pos)}${gridlines ? "<c:majorGridlines/>" : ""}` +
    `<c:numFmt formatCode="${percent ? "0%" : "General"}" sourceLinked="1"/>${axisTail(crossAx)}<c:crossBetween val="${between}"/></c:valAx>`;
  if (spec.kind === "scatter") {
    return valAx(CAT_AX_ID, "b", VAL_AX_ID, false) + valAx(VAL_AX_ID, "l", CAT_AX_ID, true);
  }
  const catAx =
    `<c:catAx><c:axId val="${CAT_AX_ID}"/>${scaling}${axisHead(horizontalBars ? "l" : "b")}` +
    `<c:numFmt formatCode="General" sourceLinked="1"/>${axisTail(VAL_AX_ID)}<c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>`;
  return catAx + valAx(VAL_AX_ID, horizontalBars ? "b" : "l", CAT_AX_ID, true);
}

function lineGrouping(g: ChartGrouping | undefined): string {
  return g === "stacked" || g === "percentStacked" ? g : "standard";
}

function plotMarkup(spec: ChartSpec): string {
  const sers = spec.series.map((_, i) => seriesMarkup(spec, i)).join("");
  const labels = spec.dataLabels ? DATA_LABELS : "";
  const axIds = `<c:axId val="${CAT_AX_ID}"/><c:axId val="${VAL_AX_ID}"/>`;
  switch (spec.kind) {
    case "column":
    case "bar": {
      const grouping = spec.grouping ?? "clustered";
      const overlap = grouping === "clustered" ? "" : `<c:overlap val="100"/>`;
      return `<c:barChart><c:barDir val="${spec.kind === "bar" ? "bar" : "col"}"/><c:grouping val="${grouping}"/><c:varyColors val="0"/>${sers}${labels}<c:gapWidth val="150"/>${overlap}${axIds}</c:barChart>${axesMarkup(spec)}`;
    }
    case "line":
      return `<c:lineChart><c:grouping val="${lineGrouping(spec.grouping)}"/><c:varyColors val="0"/>${sers}${labels}<c:marker val="1"/>${axIds}</c:lineChart>${axesMarkup(spec)}`;
    case "area":
      return `<c:areaChart><c:grouping val="${lineGrouping(spec.grouping)}"/><c:varyColors val="0"/>${sers}${labels}${axIds}</c:areaChart>${axesMarkup(spec)}`;
    case "pie":
      return `<c:pieChart><c:varyColors val="1"/>${sers}${labels}<c:firstSliceAng val="0"/></c:pieChart>`;
    case "doughnut":
      return `<c:doughnutChart><c:varyColors val="1"/>${sers}${labels}<c:firstSliceAng val="0"/><c:holeSize val="50"/></c:doughnutChart>`;
    case "scatter":
      return `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>${sers}${labels}${axIds}</c:scatterChart>${axesMarkup(spec)}`;
    case "radar":
      return `<c:radarChart><c:radarStyle val="marker"/><c:varyColors val="0"/>${sers}${labels}${axIds}</c:radarChart>${axesMarkup(spec)}`;
    default:
      return "";
  }
}

export function validateChartSpec(spec: ChartSpec): void {
  if (spec.series.length === 0) throw new Error("A chart needs at least one series.");
  if (spec.categories.length === 0) throw new Error("A chart needs at least one category.");
  for (const s of spec.series) {
    if (s.values.length !== spec.categories.length) {
      throw new Error(
        `Series ${JSON.stringify(s.name)} has ${s.values.length} values for ${spec.categories.length} categories.`,
      );
    }
  }
}

/** The `c:chartSpace` part for `spec`; `workbookRelId` links the embedded workbook. */
export function chartSpaceXml(spec: ChartSpec, workbookRelId: string): string {
  validateChartSpec(spec);
  const title =
    spec.title !== undefined
      ? `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${escapeXml(spec.title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>`
      : `<c:autoTitleDeleted val="1"/>`;
  const legend =
    spec.legend === "none"
      ? ""
      : `<c:legend><c:legendPos val="${LEGEND_CODE[spec.legend]}"/><c:overlay val="0"/></c:legend>`;
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
    `<c:chartSpace xmlns:c="${C_NS}" xmlns:a="${A_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
    `<c:date1904 val="0"/><c:roundedCorners val="0"/>` +
    `<c:chart>${title}<c:plotArea><c:layout/>${plotMarkup(spec)}</c:plotArea>${legend}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>` +
    `<c:externalData r:id="${workbookRelId}"><c:autoUpdate val="0"/></c:externalData>` +
    `</c:chartSpace>`
  );
}

/** The embedded workbook's cells: categories in column A, one column per series. */
export function chartSheetRows(spec: ChartSpec): CellValue[][] {
  const header: CellValue[] = ["", ...spec.series.map((s) => s.name)];
  const body = spec.categories.map((cat, r): CellValue[] => {
    const asNumber = Number(cat);
    const first: CellValue =
      spec.kind === "scatter" && cat !== "" && Number.isFinite(asNumber) ? asNumber : cat;
    return [first, ...spec.series.map((s) => s.values[r] ?? null)];
  });
  return [header, ...body];
}

// --- reading --------------------------------------------------------------------

const CHART_ELEMENTS: Readonly<Record<string, ChartKind>> = {
  barChart: "column",
  bar3DChart: "column",
  lineChart: "line",
  line3DChart: "line",
  areaChart: "area",
  area3DChart: "area",
  pieChart: "pie",
  pie3DChart: "pie",
  ofPieChart: "pie",
  doughnutChart: "doughnut",
  scatterChart: "scatter",
  radarChart: "radar",
};

function cachePoints(ref: XmlElement | undefined): string[] {
  if (!ref) return [];
  const cache =
    descendant(ref, C_NS, "strCache") ??
    descendant(ref, C_NS, "numCache") ??
    child(ref, C_NS, "strLit") ??
    child(ref, C_NS, "numLit") ??
    descendant(ref, C_NS, "multiLvlStrCache");
  if (!cache) return [];
  const count = Number(attr(child(cache, C_NS, "ptCount"), "val") ?? "0");
  const out: string[] = Array.from({ length: Number.isFinite(count) ? count : 0 }, () => "");
  // multi-level caches keep their points under c:lvl; the first level is the innermost.
  const holder = child(cache, C_NS, "lvl") ?? cache;
  for (const pt of elementChildren(holder)) {
    if (pt.name.local !== "pt") continue;
    const idx = Number(attr(pt, "idx"));
    if (Number.isInteger(idx) && idx >= 0) out[idx] = textOf(child(pt, C_NS, "v"));
  }
  return out;
}

function textOf(el: XmlElement | undefined): string {
  return el ? el.children.map((c) => (c.kind === "text" ? c.value : "")).join("") : "";
}

function richText(el: XmlElement | undefined): string | undefined {
  const rich = path(el, [C_NS, "tx"], [C_NS, "rich"]);
  if (!rich) return undefined;
  const paragraphs = elementChildren(rich).filter((p) => p.name.local === "p");
  return paragraphs
    .map((p) =>
      elementChildren(p)
        .filter((r) => r.name.local === "r" || r.name.local === "fld")
        .map((r) => child(r, A_NS, "t"))
        .map((t) => (t ? t.children.map((c) => (c.kind === "text" ? c.value : "")).join("") : ""))
        .join(""),
    )
    .join("\n");
}

function boolVal(el: XmlElement | undefined, fallback: boolean): boolean {
  const v = attr(el, "val");
  return v === undefined ? fallback : v === "1" || v === "true";
}

/** Read a chart part back into a {@link ChartSpec} (the first chart group of the plot area). */
export function readChartSpace(xml: string): ChartSpec | undefined {
  const root = parseXml(xml).root;
  const chart = child(root, C_NS, "chart");
  const plot = child(chart, C_NS, "plotArea");
  const group = elementChildren(plot ?? root).find(
    (c) => c.name.uri === C_NS && c.name.local in CHART_ELEMENTS,
  );
  if (!chart || !group) return undefined;
  let kind = CHART_ELEMENTS[group.name.local] ?? "column";
  if (kind === "column" && attr(child(group, C_NS, "barDir"), "val") === "bar") kind = "bar";
  const groupingVal = attr(child(group, C_NS, "grouping"), "val");
  const grouping: ChartGrouping | undefined =
    groupingVal === "stacked" || groupingVal === "percentStacked"
      ? groupingVal
      : kind === "pie" || kind === "doughnut" || kind === "scatter" || kind === "radar"
        ? undefined
        : "clustered";
  const sers = elementChildren(group).filter((c) => c.name.local === "ser");
  const first = sers[0];
  const categories = cachePoints(child(first, C_NS, "cat") ?? child(first, C_NS, "xVal"));
  const series: ChartSeries[] = sers.map((ser, i) => {
    const tx = child(ser, C_NS, "tx");
    const name = cachePoints(tx)[0] ?? textOf(child(tx, C_NS, "v"));
    const raw = cachePoints(child(ser, C_NS, "val") ?? child(ser, C_NS, "yVal"));
    const values = Array.from({ length: Math.max(categories.length, raw.length) }, (_, k) => {
      const s = raw[k];
      const n = s === undefined || s === "" ? Number.NaN : Number(s);
      return Number.isFinite(n) ? n : null;
    });
    const spPr = child(ser, C_NS, "spPr");
    const color =
      readColor(child(spPr, A_NS, "solidFill")) ??
      readColor(path(spPr, [A_NS, "ln"], [A_NS, "solidFill"]));
    const seriesName = name || `Series ${i + 1}`;
    // Pie slices are colored per point; a series color there is not the slice color.
    return color && kind !== "pie" && kind !== "doughnut"
      ? { name: seriesName, values, color: color.hex }
      : { name: seriesName, values };
  });
  const cats = categories.length
    ? categories
    : (series[0]?.values.map((_, k) => String(k + 1)) ?? []);
  const titleEl = child(chart, C_NS, "title");
  const autoDeleted = boolVal(child(chart, C_NS, "autoTitleDeleted"), false);
  // Without its own text, a chart title shows the only series' name.
  const title = titleEl
    ? (richText(titleEl) ?? (series.length === 1 ? series[0]?.name : undefined))
    : !autoDeleted && series.length === 1
      ? series[0]?.name
      : undefined;
  const legendPos = attr(path(chart, [C_NS, "legend"], [C_NS, "legendPos"]), "val");
  const legend = child(chart, C_NS, "legend")
    ? (LEGEND_FROM_CODE[legendPos ?? "r"] ?? "right")
    : "none";
  const groupLabels = child(group, C_NS, "dLbls");
  const serLabels = first ? child(first, C_NS, "dLbls") : undefined;
  const dataLabels = boolVal(child(groupLabels ?? serLabels, C_NS, "showVal"), false);
  const markerSymbol = attr(path(first, [C_NS, "marker"], [C_NS, "symbol"]), "val");
  const markers =
    kind === "line" || kind === "radar"
      ? markerSymbol !== undefined && markerSymbol !== "none"
      : kind === "scatter"
        ? attr(child(group, C_NS, "scatterStyle"), "val") === "marker" ||
          child(path(child(first, C_NS, "spPr"), [A_NS, "ln"]), A_NS, "noFill") !== undefined
        : undefined;
  return {
    kind,
    ...(grouping ? { grouping } : {}),
    ...(markers !== undefined ? { markers } : {}),
    ...(title !== undefined ? { title } : {}),
    legend,
    dataLabels,
    categories: cats,
    series,
  };
}
