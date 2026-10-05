/**
 * Table → HTML for the editing canvas: the table's grid and preferred width,
 * merged cells (`gridSpan` → colspan, `vMerge` → rowspan), row heights,
 * cell margins, vertical alignment and text direction, and every cell's
 * borders and shading after the table style's conditional formatting (see
 * ./table-format.ts). Borders the document leaves out are drawn with the
 * `--wk-gridline` CSS variable, so the page can show Word's dotted
 * gridlines (Table Layout ▸ View Gridlines) without changing the document.
 */

import { type Docx, tableCellBlocks, type WmlParagraph, type WmlTable } from "@office-kit/docx";
import type { CellTextFormat } from "./resolve.js";
import {
  type BorderSpec,
  type ResolvedCell,
  type ResolvedTableFormat,
  resolveTable,
} from "./table-format.js";

/** Where a paragraph sits in a table: its cell (`"row,col"`) and index among the cell's paragraphs. */
export interface CellAnchor {
  readonly coord: string;
  readonly para: number;
}

/** Renders one paragraph; `anchor` places it in a cell, `text` adds the table style's formatting. */
export type ParagraphRenderer = (
  para: WmlParagraph,
  anchor: CellAnchor | undefined,
  text: CellTextFormat | undefined,
) => string;

export interface TableRenderContext {
  readonly doc: Docx;
  /** The table's index in the body, for the `data-wk-*` anchors. */
  readonly block: number;
  readonly paragraph: ParagraphRenderer;
}

const TWIPS_PER_POINT = 20;
const EIGHTHS_PER_POINT = 8;
// ST_TblWidth `pct` is in fiftieths of a percent (§17.18.90).
const PCT_UNITS_PER_PERCENT = 50;
// A zero size still draws Word's thinnest line (¼ pt).
const MIN_BORDER_EIGHTHS = 2;
const HEX_COLOR = /^[0-9A-Fa-f]{6}$/;
const NO_BORDER: ReadonlySet<string> = new Set(["none", "nil"]);
const GRIDLINE = "var(--wk-gridline,none)";

// ST_Border styles CSS can draw; the rest fall back to solid.
const BORDER_STYLE_CSS: Readonly<Record<string, string>> = {
  double: "double",
  triple: "double",
  thinThickSmallGap: "double",
  thickThinSmallGap: "double",
  thinThickThinSmallGap: "double",
  thinThickMediumGap: "double",
  thickThinMediumGap: "double",
  thinThickThinMediumGap: "double",
  thinThickLargeGap: "double",
  thickThinLargeGap: "double",
  thinThickThinLargeGap: "double",
  doubleWave: "double",
  dotted: "dotted",
  dashed: "dashed",
  dashSmallGap: "dashed",
  dotDash: "dashed",
  dotDotDash: "dotted",
  dashDotStroked: "dashed",
  threeDEmboss: "ridge",
  threeDEngrave: "groove",
  outset: "outset",
  inset: "inset",
};
// A double line is two strokes of the border width with one between (§17.18.2).
const DOUBLE_LINE_FACTOR = 3;

function escapeAttr(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;");
}

function colorCss(color: string): string {
  return HEX_COLOR.test(color) ? `#${color}` : "#000";
}

function borderCss(spec: BorderSpec | undefined): string {
  if (!spec || NO_BORDER.has(spec.style)) return GRIDLINE;
  const style = BORDER_STYLE_CSS[spec.style] ?? "solid";
  const factor = style === "double" ? DOUBLE_LINE_FACTOR : 1;
  const width = (Math.max(spec.size, MIN_BORDER_EIGHTHS) * factor) / EIGHTHS_PER_POINT;
  return `${width}pt ${style} ${colorCss(spec.color)}`;
}

/** A diagonal line across the cell as a background gradient. */
function diagonalCss(spec: BorderSpec | undefined, direction: string): string | undefined {
  if (!spec || NO_BORDER.has(spec.style)) return undefined;
  const half = Math.max(spec.size, MIN_BORDER_EIGHTHS) / EIGHTHS_PER_POINT / 2;
  const color = colorCss(spec.color);
  return `linear-gradient(${direction},transparent calc(50% - ${half}pt),${color} calc(50% - ${half}pt),${color} calc(50% + ${half}pt),transparent calc(50% + ${half}pt))`;
}

const pt = (twips: number): string => `${twips / TWIPS_PER_POINT}pt`;

const VERTICAL_ALIGN: Readonly<Record<string, string>> = {
  top: "top",
  center: "middle",
  both: "middle",
  bottom: "bottom",
};

// ST_TextDirection, in the transitional (lrTb …) and strict (lr …) spellings.
const VERTICAL_TOP_DOWN: ReadonlySet<string> = new Set(["tbRl", "tbRlV", "rl", "rlV", "tbLrV", "tb"]);
const VERTICAL_BOTTOM_UP: ReadonlySet<string> = new Set(["btLr", "lr"]);

function textDirectionCss(direction: string | undefined): string | undefined {
  if (direction === undefined) return undefined;
  if (VERTICAL_TOP_DOWN.has(direction)) return "writing-mode:vertical-rl";
  if (VERTICAL_BOTTOM_UP.has(direction)) return "writing-mode:vertical-rl;transform:rotate(180deg)";
  return undefined;
}

function tableCss(format: ResolvedTableFormat, autoLayout: boolean): string {
  const css: string[] = [];
  const width = format.width;
  const total = format.columns.reduce((a, b) => a + b, 0);
  if (width?.type === "dxa" && width.value > 0) css.push(`width:${pt(width.value)}`);
  else if (width?.type === "pct" && width.value > 0) {
    css.push(`width:${width.value / PCT_UNITS_PER_PERCENT}%`);
  } else if (!autoLayout && total > 0) css.push(`width:${pt(total)}`);
  if (autoLayout) css.push("table-layout:auto");
  const align = format.alignment;
  if (align === "center") css.push("margin-left:auto", "margin-right:auto");
  else if (align === "right" || align === "end") css.push("margin-left:auto", "margin-right:0");
  else css.push(`margin-left:${pt(format.leftEdge)}`);
  if (format.cellSpacing > 0) {
    // Word's spacing value insets every cell on each side, so cells sit twice it apart.
    css.push("border-collapse:separate", `border-spacing:${pt(format.cellSpacing * 2)}`);
    const b = format.borders;
    css.push(
      `border-top:${borderCss(b.top)}`,
      `border-bottom:${borderCss(b.bottom)}`,
      `border-left:${borderCss(b.left)}`,
      `border-right:${borderCss(b.right)}`,
    );
  }
  return css.join(";");
}

function cellCss(cell: ResolvedCell, bottom: BorderSpec | undefined): string {
  const b = cell.borders;
  const m = cell.margins;
  const css = [
    `padding:${pt(m.top)} ${pt(m.right)} ${pt(m.bottom)} ${pt(m.left)}`,
    `border-top:${borderCss(b.top)}`,
    `border-bottom:${borderCss(bottom)}`,
    `border-left:${borderCss(b.left)}`,
    `border-right:${borderCss(b.right)}`,
  ];
  const layers = [diagonalCss(b.tl2br, "to top right"), diagonalCss(b.tr2bl, "to bottom right")].filter(
    (l): l is string => l !== undefined,
  );
  if (layers.length) css.push(`background-image:${layers.join(",")}`);
  if (cell.background) css.push(`background-color:${cell.background}`);
  const vAlign = cell.verticalAlign && VERTICAL_ALIGN[cell.verticalAlign];
  if (vAlign) css.push(`vertical-align:${vAlign}`);
  if (cell.noWrap) css.push("white-space:nowrap");
  return css.join(";");
}

/** Empty grid columns before / after a row's cells (`w:gridBefore` / `w:gridAfter`). */
function skip(n: number): string {
  return n > 0 ? `<td class="wk-td-skip" colspan="${n}" contenteditable="false"></td>` : "";
}

/** Strip the editing anchors from a nested table: only the outer cell is addressable. */
function withoutAnchors(html: string): string {
  return html.replace(/ data-wk-[a-z]+="[^"]*"/g, "");
}

/** Render a table (and, read-only, any tables nested in its cells). */
export function renderTable(table: WmlTable, ctx: TableRenderContext): string {
  const format = resolveTable(ctx.doc, table);
  // A merged cell's bottom edge is the bottom of the merge's last row.
  const byGrid = new Map<string, ResolvedCell>();
  format.cells.forEach((row, r) => {
    for (const cell of row) byGrid.set(`${r}:${cell.placement.gridStart}`, cell);
  });
  const autoLayout =
    format.layout === "autofit" &&
    (format.width === undefined || format.width.type === "auto") &&
    table.rows.every((row) =>
      row.cells.every((cell) => {
        const tcW = cell.tcPr?.children.find(
          (c) => c.kind === "element" && c.name.local === "tcW",
        );
        return tcW?.kind === "element" && tcW.attrs.some((a) => a.name.local === "type" && a.value === "auto");
      }),
    );
  const rows = table.rows
    .map((row, r) => {
      const resolvedRow = format.rows[r];
      const resolved = format.cells[r] ?? [];
      const exact = resolvedRow?.height?.rule === "exact" ? resolvedRow.height.value : undefined;
      const cells = row.cells
        .map((cell, c) => {
          const rc = resolved[c];
          if (!rc || rc.placement.rowSpan === 0) return "";
          const p = rc.placement;
          const coord = `${r},${c}`;
          const last = byGrid.get(`${r + p.rowSpan - 1}:${p.gridStart}`) ?? rc;
          const css = cellCss(rc, last.borders.bottom);
          let paraIndex = 0;
          const body = tableCellBlocks(cell)
            .map((b) => {
              if (b.kind === "paragraph") {
                return ctx.paragraph(b, { coord, para: paraIndex++ }, rc.text);
              }
              if (b.kind === "table") {
                return `<div class="wk-nested" contenteditable="false">${withoutAnchors(renderTable(b, ctx))}</div>`;
              }
              return "";
            })
            .join("");
          const box: string[] = [];
          const direction = textDirectionCss(rc.textDirection);
          if (direction) box.push(direction);
          if (exact !== undefined) {
            box.push(`height:${pt(Math.max(0, exact - rc.margins.top - rc.margins.bottom))}`, "overflow:hidden");
          }
          const content = box.length ? `<div class="wk-cell-box" style="${box.join(";")}">${body}</div>` : body;
          const span = [
            p.gridSpan > 1 ? ` colspan="${p.gridSpan}"` : "",
            p.rowSpan > 1 ? ` rowspan="${p.rowSpan}"` : "",
          ].join("");
          return `<td class="wk-td" data-wk-block="${ctx.block}" data-wk-cell="${coord}" data-wk-grid="${p.gridStart}" data-wk-span="${p.gridSpan}"${span} style="${escapeAttr(css)}">${content}</td>`;
        })
        .join("");
      const trCss: string[] = [];
      if (resolvedRow?.height && resolvedRow.height.rule !== "auto") {
        trCss.push(`height:${pt(resolvedRow.height.value)}`);
      }
      if (resolvedRow?.hidden) trCss.push("display:none");
      const cls = resolvedRow?.header ? "wk-tr wk-tr-header" : "wk-tr";
      const style = trCss.length ? ` style="${trCss.join(";")}"` : "";
      return `<tr class="${cls}"${style}>${skip(resolvedRow?.gridBefore ?? 0)}${cells}${skip(resolvedRow?.gridAfter ?? 0)}</tr>`;
    })
    .join("");
  const cols =
    format.columns.length && !autoLayout
      ? `<colgroup>${format.columns.map((w) => `<col style="width:${pt(w)}">`).join("")}</colgroup>`
      : "";
  return `<table class="wk-table" data-wk-block="${ctx.block}" style="${escapeAttr(tableCss(format, autoLayout))}">${cols}<tbody>${rows}</tbody></table>`;
}
