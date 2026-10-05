/**
 * Charts on the canvas: a {@link ChartSpec} (read from the chart part's value
 * caches) drawn as SVG in the layout Word uses for its default chart style —
 * title on top, legend at its position, light horizontal gridlines, axis
 * labels in grey.
 */

import type { ChartSpec } from "@office-kit/docx";

const SERIES_COLORS = ["4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47"] as const;
const FONT = "Calibri,Carlito,'Segoe UI',sans-serif";
const TITLE_PX = 18.7; // 14 pt
const LABEL_PX = 12; // 9 pt
const TEXT = "#595959";
const GRID = "#D9D9D9";
const PAD = 10;
const LEGEND_SWATCH = 8;
// Category band share taken by bars (Word's default gap width is 150 %).
const GAP_WIDTH = 1.5;
const MARKER_R = 3;
const CHAR_W = 0.55; // average glyph width / font size, for label sizing

function esc(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

const HEX = /^[0-9A-Fa-f]{6}$/;

function seriesColor(spec: ChartSpec, i: number): string {
  const c = spec.series[i]?.color;
  return `#${c && HEX.test(c) ? c : (SERIES_COLORS[i % SERIES_COLORS.length] ?? "4472C4")}`;
}

function pointColor(i: number): string {
  return `#${SERIES_COLORS[i % SERIES_COLORS.length] ?? "4472C4"}`;
}

function text(
  x: number,
  y: number,
  s: string,
  size: number,
  anchor = "middle",
  extra = "",
): string {
  return `<text x="${r(x)}" y="${r(y)}" font-size="${size}" font-family="${FONT}" fill="${TEXT}" text-anchor="${anchor}" dominant-baseline="middle"${extra}>${esc(s)}</text>`;
}

function r(n: number): number {
  return Math.round(n * 10) / 10;
}

/** Axis ticks at 1 / 2 / 5 × 10ⁿ steps, as Word's automatic scaling picks them. */
function niceScale(min: number, max: number): { min: number; max: number; step: number } {
  if (min === max) max = min + 1;
  const range = max - min;
  const rough = range / 6;
  const mag = 10 ** Math.floor(Math.log10(rough));
  const norm = rough / mag;
  const step = (norm <= 1 ? 1 : norm <= 2 ? 2 : norm <= 5 ? 5 : 10) * mag;
  return { min: Math.floor(min / step) * step, max: Math.ceil(max / step) * step, step };
}

function formatNumber(n: number, percent: boolean): string {
  if (percent) return `${Math.round(n * 100)}%`;
  return String(Math.round(n * 1e6) / 1e6);
}

interface Area {
  x: number;
  y: number;
  w: number;
  h: number;
}

function legendEntries(spec: ChartSpec): Array<{ label: string; color: string }> {
  if (spec.kind === "pie" || spec.kind === "doughnut") {
    return spec.categories.map((c, i) => ({ label: c, color: pointColor(i) }));
  }
  return spec.series.map((s, i) => ({ label: s.name, color: seriesColor(spec, i) }));
}

function drawLegend(spec: ChartSpec, area: Area): { svg: string; plot: Area } {
  if (spec.legend === "none") return { svg: "", plot: area };
  const entries = legendEntries(spec);
  const itemW = (e: { label: string }): number =>
    LEGEND_SWATCH + 4 + e.label.length * LABEL_PX * CHAR_W + 12;
  const row = LABEL_PX + 6;
  let svg = "";
  if (spec.legend === "top" || spec.legend === "bottom") {
    const total = entries.reduce((sum, e) => sum + itemW(e), 0);
    let x = area.x + Math.max(0, (area.w - total) / 2);
    const y = spec.legend === "top" ? area.y + row / 2 : area.y + area.h - row / 2;
    for (const e of entries) {
      svg += `<rect x="${r(x)}" y="${r(y - LEGEND_SWATCH / 2)}" width="${LEGEND_SWATCH}" height="${LEGEND_SWATCH}" fill="${e.color}"/>`;
      svg += text(x + LEGEND_SWATCH + 4, y, e.label, LABEL_PX, "start");
      x += itemW(e);
    }
    const plot =
      spec.legend === "top"
        ? { ...area, y: area.y + row + 4, h: area.h - row - 4 }
        : { ...area, h: area.h - row - 4 };
    return { svg, plot };
  }
  const width = Math.max(...entries.map(itemW));
  const x = spec.legend === "left" ? area.x : area.x + area.w - width;
  let y = area.y + Math.max(0, (area.h - entries.length * row) / 2) + row / 2;
  for (const e of entries) {
    svg += `<rect x="${r(x)}" y="${r(y - LEGEND_SWATCH / 2)}" width="${LEGEND_SWATCH}" height="${LEGEND_SWATCH}" fill="${e.color}"/>`;
    svg += text(x + LEGEND_SWATCH + 4, y, e.label, LABEL_PX, "start");
    y += row;
  }
  const plot =
    spec.legend === "left"
      ? { ...area, x: area.x + width, w: area.w - width }
      : { ...area, w: area.w - width };
  return { svg, plot };
}

/** Per category: the value each series contributes, after stacking / 100 % scaling. */
function stackedValues(spec: ChartSpec): { base: number[][]; top: number[][]; percent: boolean } {
  const n = spec.categories.length;
  const stacked = spec.grouping === "stacked" || spec.grouping === "percentStacked";
  const percent = spec.grouping === "percentStacked";
  const base = spec.series.map(() => Array.from({ length: n }, () => 0));
  const top = spec.series.map((s) => s.values.map((v) => v ?? 0));
  if (!stacked) return { base, top, percent: false };
  for (let c = 0; c < n; c++) {
    const total = percent
      ? spec.series.reduce((sum, s) => sum + Math.abs(s.values[c] ?? 0), 0) || 1
      : 1;
    let pos = 0;
    let neg = 0;
    for (const [i, s] of spec.series.entries()) {
      const v = (s.values[c] ?? 0) / total;
      const row = base[i];
      const topRow = top[i];
      if (!row || !topRow) continue;
      if (v >= 0) {
        row[c] = pos;
        pos += v;
        topRow[c] = pos;
      } else {
        row[c] = neg;
        neg += v;
        topRow[c] = neg;
      }
    }
  }
  return { base, top, percent };
}

function categoryChart(spec: ChartSpec, plot: Area): string {
  const n = Math.max(1, spec.categories.length);
  const horizontal = spec.kind === "bar";
  const { base, top, percent } = stackedValues(spec);
  const all = [...base.flat(), ...top.flat(), 0];
  const scale = niceScale(Math.min(...all), Math.max(...all));
  const tickCount = Math.round((scale.max - scale.min) / scale.step);
  const valueLabelW =
    Math.max(
      ...Array.from(
        { length: tickCount + 1 },
        (_, i) => formatNumber(scale.min + i * scale.step, percent).length,
      ),
    ) *
    LABEL_PX *
    CHAR_W;
  // Reserve the axis label gutters.
  const area: Area = horizontal
    ? {
        x: plot.x + Math.max(...spec.categories.map((c) => c.length)) * LABEL_PX * CHAR_W + 8,
        y: plot.y,
        w: 0,
        h: plot.h - LABEL_PX - 6,
      }
    : { x: plot.x + valueLabelW + 8, y: plot.y + 4, w: 0, h: plot.h - LABEL_PX - 10 };
  area.w = plot.x + plot.w - area.x;
  const valuePos = (v: number): number =>
    horizontal
      ? area.x + ((v - scale.min) / (scale.max - scale.min)) * area.w
      : area.y + area.h - ((v - scale.min) / (scale.max - scale.min)) * area.h;
  const band = (horizontal ? area.h : area.w) / n;
  let svg = "";
  // Gridlines and value labels.
  for (let i = 0; i <= tickCount; i++) {
    const v = scale.min + i * scale.step;
    const p = valuePos(v);
    if (horizontal) {
      svg += `<line x1="${r(p)}" y1="${r(area.y)}" x2="${r(p)}" y2="${r(area.y + area.h)}" stroke="${GRID}" stroke-width="0.75"/>`;
      svg += text(p, area.y + area.h + LABEL_PX / 2 + 4, formatNumber(v, percent), LABEL_PX);
    } else {
      svg += `<line x1="${r(area.x)}" y1="${r(p)}" x2="${r(area.x + area.w)}" y2="${r(p)}" stroke="${GRID}" stroke-width="0.75"/>`;
      svg += text(area.x - 4, p, formatNumber(v, percent), LABEL_PX, "end");
    }
  }
  // Category axis line and labels (bar charts list categories bottom-up, as Word does).
  const catPos = (c: number): number =>
    horizontal ? area.y + area.h - (c + 0.5) * band : area.x + (c + 0.5) * band;
  const zero = valuePos(Math.max(scale.min, Math.min(0, scale.max)));
  svg += horizontal
    ? `<line x1="${r(zero)}" y1="${r(area.y)}" x2="${r(zero)}" y2="${r(area.y + area.h)}" stroke="${GRID}"/>`
    : `<line x1="${r(area.x)}" y1="${r(zero)}" x2="${r(area.x + area.w)}" y2="${r(zero)}" stroke="${GRID}"/>`;
  for (const [c, label] of spec.categories.entries()) {
    svg += horizontal
      ? text(area.x - 4, catPos(c), label, LABEL_PX, "end")
      : text(catPos(c), area.y + area.h + LABEL_PX / 2 + 4, label, LABEL_PX);
  }
  const labels: string[] = [];
  const stacked = spec.grouping === "stacked" || spec.grouping === "percentStacked";
  if (spec.kind === "column" || spec.kind === "bar") {
    const groups = stacked ? 1 : spec.series.length;
    const barW = band / (groups + GAP_WIDTH);
    for (const [i, s] of spec.series.entries()) {
      for (let c = 0; c < n; c++) {
        if (s.values[c] === null || s.values[c] === undefined) continue;
        const a = valuePos(base[i]?.[c] ?? 0);
        const b = valuePos(top[i]?.[c] ?? 0);
        const offset = (stacked ? 0 : i) * barW - (groups * barW) / 2;
        const across = catPos(c) + offset;
        const [x, y, w, h] = horizontal
          ? [Math.min(a, b), across, Math.abs(b - a), barW]
          : [across, Math.min(a, b), barW, Math.abs(b - a)];
        svg += `<rect x="${r(x)}" y="${r(y)}" width="${r(w)}" height="${r(h)}" fill="${seriesColor(spec, i)}"/>`;
        if (spec.dataLabels) {
          const label = String(s.values[c]);
          labels.push(
            horizontal
              ? text(
                  stacked ? x + w / 2 : x + w + 4,
                  y + h / 2,
                  label,
                  LABEL_PX,
                  stacked ? "middle" : "start",
                )
              : text(x + w / 2, stacked ? y + h / 2 : y - LABEL_PX / 2 - 2, label, LABEL_PX),
          );
        }
      }
    }
  } else {
    // Line / area: points at band centers (line) or edges (area), Word's crossBetween.
    const at = (c: number): number =>
      spec.kind === "area" ? area.x + (n === 1 ? area.w / 2 : (c * area.w) / (n - 1)) : catPos(c);
    for (const [i, s] of spec.series.entries()) {
      const pts = Array.from({ length: n }, (_, c) => [at(c), valuePos(top[i]?.[c] ?? 0)] as const);
      const col = seriesColor(spec, i);
      if (spec.kind === "area") {
        const lower = Array.from(
          { length: n },
          (_, c) => [at(c), valuePos(base[i]?.[c] ?? 0)] as const,
        ).toReversed();
        svg += `<path d="M${[...pts, ...lower].map(([x, y]) => `${r(x)},${r(y)}`).join("L")}Z" fill="${col}"/>`;
      } else {
        const segs = pts.filter((_, c) => s.values[c] !== null);
        svg += `<polyline points="${segs.map(([x, y]) => `${r(x)},${r(y)}`).join(" ")}" fill="none" stroke="${col}" stroke-width="2.25" stroke-linejoin="round" stroke-linecap="round"/>`;
        if (spec.markers) {
          for (const [x, y] of segs)
            svg += `<circle cx="${r(x)}" cy="${r(y)}" r="${MARKER_R}" fill="${col}"/>`;
        }
      }
      if (spec.dataLabels) {
        for (const [c, [x, y]] of pts.entries()) {
          const v = s.values[c];
          if (v !== null && v !== undefined)
            labels.push(text(x, y - LABEL_PX / 2 - 3, String(v), LABEL_PX));
        }
      }
    }
  }
  return svg + labels.join("");
}

function pieChart(spec: ChartSpec, plot: Area): string {
  const values = (spec.series[0]?.values ?? []).map((v) => Math.max(0, v ?? 0));
  const total = values.reduce((a, b) => a + b, 0) || 1;
  const cx = plot.x + plot.w / 2;
  const cy = plot.y + plot.h / 2;
  const rad = Math.max(1, Math.min(plot.w, plot.h) / 2 - 4);
  const hole = spec.kind === "doughnut" ? rad * 0.5 : 0;
  let angle = -Math.PI / 2;
  let svg = "";
  const labels: string[] = [];
  for (const [i, v] of values.entries()) {
    const sweep = (v / total) * Math.PI * 2;
    const a2 = angle + sweep;
    const large = sweep > Math.PI ? 1 : 0;
    const p = (a: number, rr: number): string =>
      `${r(cx + rr * Math.cos(a))},${r(cy + rr * Math.sin(a))}`;
    const full = sweep >= Math.PI * 2 - 1e-6;
    const d = full
      ? `M${p(angle, rad)}A${rad},${rad} 0 1 1 ${p(angle + Math.PI, rad)}A${rad},${rad} 0 1 1 ${p(angle, rad)}Z` +
        (hole
          ? `M${p(angle, hole)}A${hole},${hole} 0 1 0 ${p(angle + Math.PI, hole)}A${hole},${hole} 0 1 0 ${p(angle, hole)}Z`
          : "")
      : hole
        ? `M${p(angle, rad)}A${rad},${rad} 0 ${large} 1 ${p(a2, rad)}L${p(a2, hole)}A${hole},${hole} 0 ${large} 0 ${p(angle, hole)}Z`
        : `M${r(cx)},${r(cy)}L${p(angle, rad)}A${rad},${rad} 0 ${large} 1 ${p(a2, rad)}Z`;
    svg += `<path d="${d}" fill="${pointColor(i)}" stroke="#fff" stroke-width="1" fill-rule="evenodd"/>`;
    if (spec.dataLabels && v > 0) {
      const mid = angle + sweep / 2;
      const lr = hole ? (rad + hole) / 2 : rad * 0.65;
      labels.push(
        text(
          cx + lr * Math.cos(mid),
          cy + lr * Math.sin(mid),
          String(spec.series[0]?.values[i] ?? ""),
          LABEL_PX,
          "middle",
          ` fill-opacity="1"`,
        ),
      );
    }
    angle = a2;
  }
  return svg + labels.join("");
}

function scatterChart(spec: ChartSpec, plot: Area): string {
  const xs = spec.categories.map((c, i) =>
    Number.isFinite(Number(c)) && c !== "" ? Number(c) : i + 1,
  );
  const ys = spec.series.flatMap((s) => s.values.filter((v): v is number => v !== null));
  const sx = niceScale(Math.min(0, ...xs), Math.max(...xs));
  const sy = niceScale(Math.min(0, ...ys), Math.max(0, ...ys));
  const yLabelW =
    Math.max(...[sy.min, sy.max].map((v) => formatNumber(v, false).length)) * LABEL_PX * CHAR_W;
  const area: Area = { x: plot.x + yLabelW + 8, y: plot.y + 4, w: 0, h: plot.h - LABEL_PX - 10 };
  area.w = plot.x + plot.w - area.x;
  const X = (v: number): number => area.x + ((v - sx.min) / (sx.max - sx.min)) * area.w;
  const Y = (v: number): number => area.y + area.h - ((v - sy.min) / (sy.max - sy.min)) * area.h;
  let svg = "";
  for (let v = sy.min; v <= sy.max + sy.step / 2; v += sy.step) {
    svg += `<line x1="${r(area.x)}" y1="${r(Y(v))}" x2="${r(area.x + area.w)}" y2="${r(Y(v))}" stroke="${GRID}" stroke-width="0.75"/>`;
    svg += text(area.x - 4, Y(v), formatNumber(v, false), LABEL_PX, "end");
  }
  for (let v = sx.min; v <= sx.max + sx.step / 2; v += sx.step) {
    svg += text(X(v), area.y + area.h + LABEL_PX / 2 + 4, formatNumber(v, false), LABEL_PX);
  }
  for (const [i, s] of spec.series.entries()) {
    const col = seriesColor(spec, i);
    const pts = xs.flatMap((x, k) => {
      const y = s.values[k];
      return y === null || y === undefined ? [] : [[X(x), Y(y), y] as const];
    });
    if (!spec.markers) {
      svg += `<polyline points="${pts.map(([x, y]) => `${r(x)},${r(y)}`).join(" ")}" fill="none" stroke="${col}" stroke-width="2.25"/>`;
    }
    for (const [x, y, v] of pts) {
      svg += `<circle cx="${r(x)}" cy="${r(y)}" r="${MARKER_R}" fill="${col}"/>`;
      if (spec.dataLabels) svg += text(x, y - LABEL_PX / 2 - 3, String(v), LABEL_PX);
    }
  }
  return svg;
}

function radarChart(spec: ChartSpec, plot: Area): string {
  const n = Math.max(3, spec.categories.length);
  const values = spec.series.flatMap((s) => s.values.filter((v): v is number => v !== null));
  const scale = niceScale(Math.min(0, ...values), Math.max(0, ...values));
  const cx = plot.x + plot.w / 2;
  const cy = plot.y + plot.h / 2;
  const rad = Math.max(1, Math.min(plot.w, plot.h) / 2 - LABEL_PX * 1.5);
  const point = (c: number, v: number): readonly [number, number] => {
    const a = -Math.PI / 2 + (c * Math.PI * 2) / n;
    const rr = ((v - scale.min) / (scale.max - scale.min)) * rad;
    return [cx + rr * Math.cos(a), cy + rr * Math.sin(a)];
  };
  let svg = "";
  for (let v = scale.min + scale.step; v <= scale.max + scale.step / 2; v += scale.step) {
    const ring = Array.from({ length: n }, (_, c) => point(c, v));
    svg += `<polygon points="${ring.map(([x, y]) => `${r(x)},${r(y)}`).join(" ")}" fill="none" stroke="${GRID}" stroke-width="0.75"/>`;
  }
  for (let c = 0; c < n; c++) {
    const [x, y] = point(c, scale.max);
    svg += `<line x1="${r(cx)}" y1="${r(cy)}" x2="${r(x)}" y2="${r(y)}" stroke="${GRID}" stroke-width="0.75"/>`;
    const [lx, ly] = point(c, scale.max + (scale.max - scale.min) * 0.12);
    svg += text(lx, ly, spec.categories[c] ?? "", LABEL_PX);
  }
  for (const [i, s] of spec.series.entries()) {
    const col = seriesColor(spec, i);
    const pts = Array.from({ length: n }, (_, c) => point(c, s.values[c] ?? 0));
    svg += `<polygon points="${pts.map(([x, y]) => `${r(x)},${r(y)}`).join(" ")}" fill="none" stroke="${col}" stroke-width="2.25"/>`;
    if (spec.markers)
      for (const [x, y] of pts)
        svg += `<circle cx="${r(x)}" cy="${r(y)}" r="${MARKER_R}" fill="${col}"/>`;
  }
  return svg;
}

/** The SVG for a chart `width` × `height` CSS px, white background as Word's chart area. */
export function chartSvg(spec: ChartSpec, width: number, height: number): string {
  let area: Area = { x: PAD, y: PAD, w: width - 2 * PAD, h: height - 2 * PAD };
  let svg = `<rect width="${width}" height="${height}" fill="#fff" stroke="${GRID}" stroke-width="0.75"/>`;
  if (spec.title !== undefined) {
    svg += text(width / 2, PAD + TITLE_PX / 2, spec.title, TITLE_PX);
    area = { ...area, y: area.y + TITLE_PX + 6, h: area.h - TITLE_PX - 6 };
  }
  const legend = drawLegend(spec, area);
  svg += legend.svg;
  const plot = legend.plot;
  switch (spec.kind) {
    case "pie":
    case "doughnut":
      svg += pieChart(spec, plot);
      break;
    case "scatter":
      svg += scatterChart(spec, plot);
      break;
    case "radar":
      svg += radarChart(spec, plot);
      break;
    default:
      svg += categoryChart(spec, plot);
  }
  return `<svg class="wk-chart" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}" aria-hidden="true">${svg}</svg>`;
}
