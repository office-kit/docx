/**
 * AST → HTML renderer for the editing canvas.
 *
 * This is a *semantic* renderer: it maps the WordprocessingML AST to editable
 * HTML annotated with `data-wk-*` anchors so DOM selection can be mapped back to
 * document positions. It is not a page-faithful layout engine (that is
 * `@office-kit/docx-preview`'s job); it renders a continuous flow with enough
 * fidelity — bold/italic/underline/color/size/alignment, tables, list markers —
 * to edit against. Elements it does not model render as inert placeholders so
 * the document round-trips losslessly through the library even while shown here.
 */

import {
  type Docx,
  paragraphText,
  type WmlBlock,
  type WmlInline,
  type WmlParagraph,
  type WmlRun,
  type WmlTable,
  type XmlElement,
} from "@office-kit/docx";
import { highlightCss } from "./highlight.js";
import {
  type BorderSpec,
  cellBorders,
  createStyleResolver,
  type ResolvedParagraphFormat,
  type ResolvedRunFormat,
  type ResolvedTableFormat,
  resolveTable,
  type StyleResolver,
} from "./resolve.js";
import {
  createDrawingContext,
  type DrawingRenderContext,
  isDrawingOnlyRun,
  runDrawingsHtml,
} from "./render-drawing.js";
import { WML_NS } from "./wml-ns.js";

const ALIGN_TO_CSS: Record<string, string> = {
  left: "left",
  start: "left",
  center: "center",
  right: "right",
  end: "right",
  both: "justify",
  distribute: "justify",
};

const UNDERLINE_STYLE_CSS: Record<string, string> = {
  double: "text-decoration-style:double",
  dotted: "text-decoration-style:dotted",
  dash: "text-decoration-style:dashed",
  wave: "text-decoration-style:wavy",
  thick: "text-decoration-thickness:2px",
};

const VERT_ALIGN_CSS: Record<string, string> = {
  superscript: "vertical-align:super;font-size:smaller",
  subscript: "vertical-align:sub;font-size:smaller",
};

// When a document font is not installed in the browser (Calibri and Cambria
// ship with Office, not with most systems), fall back to a metric-compatible
// font or at least the right generic family, as Word's font substitution does.
const FALLBACK_FAMILY: Readonly<Record<string, string>> = {
  Calibri: "Carlito,sans-serif",
  Cambria: "Caladea,serif",
  "Times New Roman": "Tinos,'Liberation Serif',serif",
  Arial: "Arimo,'Liberation Sans',sans-serif",
  "Courier New": "Cousine,'Liberation Mono',monospace",
  Georgia: "serif",
  Garamond: "serif",
  "Book Antiqua": "serif",
  Consolas: "monospace",
  "Yu Mincho": "'Hiragino Mincho ProN',serif",
  "MS Mincho": "'Hiragino Mincho ProN',serif",
  "Yu Gothic": "'Hiragino Sans',sans-serif",
  "MS Gothic": "'Hiragino Sans',sans-serif",
};

const TWIPS_PER_POINT = 20;
const AUTO_LINE_UNIT = 240;
// CSS line-height that matches Word's single spacing for its default fonts.
const SINGLE_LINE_HEIGHT = 1.2;

function escapeHtml(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

// Values come straight from the file, and a loaded .docx is untrusted: only
// schema-valid colors become CSS, and font names lose the characters that
// could end the CSS string.
const HEX_COLOR = /^[0-9A-Fa-f]{6}$/;

/**
 * Inline CSS for a run's effective formatting. Every property is written, not
 * only the ones that differ from Normal, because the run sits inside a
 * paragraph element that carries the paragraph mark's look.
 */
function runCss(fmt: ResolvedRunFormat): string {
  const parts: string[] = [
    `font-weight:${fmt.bold ? "bold" : "normal"}`,
    `font-style:${fmt.italic ? "italic" : "normal"}`,
  ];
  const decoration: string[] = [];
  const underlined = fmt.underline !== undefined && fmt.underline !== "none";
  if (underlined) decoration.push("underline");
  if (fmt.strike) decoration.push("line-through");
  parts.push(`text-decoration-line:${decoration.length ? decoration.join(" ") : "none"}`);
  const underlineStyle = underlined ? UNDERLINE_STYLE_CSS[fmt.underline ?? ""] : undefined;
  if (underlineStyle) parts.push(underlineStyle);
  if (fmt.color && HEX_COLOR.test(fmt.color)) parts.push(`color:#${fmt.color}`);
  const highlight = fmt.highlight === undefined ? undefined : highlightCss(fmt.highlight);
  if (highlight) parts.push(`background-color:${highlight}`);
  parts.push(...fontCss(fmt));
  const vertAlign = VERT_ALIGN_CSS[fmt.vertAlign ?? ""];
  if (vertAlign) parts.push(vertAlign);
  return parts.join(";");
}

function fontCss(fmt: ResolvedRunFormat): string[] {
  const parts: string[] = [];
  if (fmt.sizeHalfPoints !== undefined) parts.push(`font-size:${fmt.sizeHalfPoints / 2}pt`);
  if (fmt.font) {
    const font = fmt.font.replace(/['"\\\n\r]/g, "");
    parts.push(`font-family:'${font}',${FALLBACK_FAMILY[font] ?? "sans-serif"}`);
  }
  return parts;
}

function paragraphCss(fmt: ResolvedParagraphFormat, mark: ResolvedRunFormat): string {
  const css = fontCss(mark);
  const align = ALIGN_TO_CSS[fmt.alignment ?? ""];
  if (align) css.push(`text-align:${align}`);
  css.push(`margin-top:${(fmt.before ?? 0) / TWIPS_PER_POINT}pt`);
  css.push(`margin-bottom:${(fmt.after ?? 0) / TWIPS_PER_POINT}pt`);
  if (fmt.left !== undefined) css.push(`margin-left:${fmt.left / TWIPS_PER_POINT}pt`);
  if (fmt.right !== undefined) css.push(`margin-right:${fmt.right / TWIPS_PER_POINT}pt`);
  if (fmt.hanging !== undefined) css.push(`text-indent:${-fmt.hanging / TWIPS_PER_POINT}pt`);
  else if (fmt.firstLine !== undefined)
    css.push(`text-indent:${fmt.firstLine / TWIPS_PER_POINT}pt`);
  if (fmt.line !== undefined && fmt.line > 0) {
    css.push(
      fmt.lineRule === "auto"
        ? `line-height:${(fmt.line / AUTO_LINE_UNIT) * SINGLE_LINE_HEIGHT}`
        : `line-height:${fmt.line / TWIPS_PER_POINT}pt`,
    );
  }
  return css.join(";");
}

/** Concatenate a run's textual pieces (text/tab/break become visible chars). */
function runText(run: WmlRun): string {
  let out = "";
  for (const piece of run.pieces) {
    switch (piece.kind) {
      case "text":
        out += piece.value;
        break;
      case "tab":
        out += "\t";
        break;
      case "break":
        out += "\n";
        break;
      case "noBreakHyphen":
      case "softHyphen":
        out += "‑";
        break;
      default:
        break;
    }
  }
  return out;
}

/**
 * Where a paragraph sits inside a table: its cell (`"row,col"`) and its index
 * among that cell's paragraphs. Without the index, every paragraph of a
 * multi-paragraph cell would map back to the first one.
 */
interface CellAnchor {
  readonly coord: string;
  readonly para: number;
}

function cellAttrs(cell: CellAnchor | undefined): string {
  return cell ? `data-wk-cell="${cell.coord}" data-wk-para="${cell.para}"` : "";
}

function renderRun(
  run: WmlRun,
  style: string,
  block: number,
  inline: number,
  cell?: CellAnchor,
): string {
  const text = runText(run);
  const attrs = [
    `data-wk-block="${block}"`,
    cellAttrs(cell),
    `data-wk-inline="${inline}"`,
    style ? `style="${escapeHtml(style)}"` : "",
  ]
    .filter(Boolean)
    .join(" ");
  // Preserve whitespace/tabs; use a zero-width space for empty runs so the
  // caret has something to land on.
  return `<span class="wk-run" ${attrs}>${escapeHtml(text) || "​"}</span>`;
}

/**
 * Visible text of an unmodelled inline (`<w:hyperlink>`, a field run built in
 * memory, `<w:ins>`, `<w:sdt>`…): its `<w:t>` descendants. `<w:instrText>`
 * (field codes) and `<w:delText>` (deleted revisions) are not visible text.
 */
function rawVisibleText(el: XmlElement): string {
  if (el.name.uri === WML_NS && el.name.local === "t") {
    return el.children
      .map((c) => (c.kind === "text" || c.kind === "cdata" ? c.value : ""))
      .join("");
  }
  if (el.name.uri === WML_NS && (el.name.local === "instrText" || el.name.local === "delText")) {
    return "";
  }
  let out = "";
  for (const child of el.children) if (child.kind === "element") out += rawVisibleText(child);
  return out;
}

/**
 * Render an unmodelled inline read-only. It carries no `data-wk-*` anchor and
 * is not a `.wk-run`, so typing never writes into it and the canvas text sync
 * leaves it alone; the library round-trips the XML untouched.
 */
function renderRawInline(inline: Extract<WmlInline, { kind: "raw" }>): string {
  const text = rawVisibleText(inline.node);
  if (!text) return "";
  const kind = inline.node.name.local === "hyperlink" ? "wk-link" : "wk-inline-raw";
  return `<span class="${kind}" contenteditable="false">${escapeHtml(text)}</span>`;
}

function renderParagraph(
  para: WmlParagraph,
  styles: StyleResolver,
  block: number,
  cell?: CellAnchor,
): string {
  const styleAttr = ` style="${escapeHtml(paragraphCss(styles.paragraph(para), styles.run(para)))}"`;
  // `data-wk-inline` counts runs only (the unit `runAtPath` resolves), so
  // raw inlines are interleaved without consuming an index.
  let runIndex = 0;
  let textRuns = 0;
  let inner = para.children
    .map((child) => {
      if (child.kind !== "run") return renderRawInline(child);
      const inline = runIndex++;
      // A picture-only run gets no editable span: typing there could not be
      // written back into the run, so the caret lives in the text around it.
      if (isDrawingOnlyRun(child)) return pictureObjectsHtml(child, block, inline, cell);
      textRuns++;
      return (
        renderRun(child, runCss(styles.run(para, child)), block, inline, cell) +
        pictureObjectsHtml(child, block, inline, cell)
      );
    })
    .join("");
  if (textRuns === 0) inner = `​${inner}`;
  const cellAttr = cell ? ` ${cellAttrs(cell)}` : "";
  return `<p class="wk-p" data-wk-block="${block}"${cellAttr}${styleAttr}>${inner}</p>`;
}

// ST_Border styles CSS can draw; the many art borders fall back to solid.
const BORDER_STYLE_CSS: Readonly<Record<string, string>> = {
  double: "double",
  dotted: "dotted",
  dashed: "dashed",
  dashSmallGap: "dashed",
  dotDash: "dashed",
  dotDotDash: "dotted",
};
const NO_BORDER: ReadonlySet<string> = new Set(["none", "nil"]);
const EIGHTHS_PER_POINT = 8;
// ST_TblWidth `pct` is in fiftieths of a percent (§17.18.90).
const PCT_UNITS_PER_PERCENT = 50;

function borderCss(spec: BorderSpec | undefined): string {
  if (!spec || NO_BORDER.has(spec.style)) return "none";
  const color = HEX_COLOR.test(spec.color) ? `#${spec.color}` : "#000";
  // A zero size still draws Word's thinnest line (¼ pt).
  const width = Math.max(spec.size, 2) / EIGHTHS_PER_POINT;
  return `${width}pt ${BORDER_STYLE_CSS[spec.style] ?? "solid"} ${color}`;
}

function tableWidthCss(format: ResolvedTableFormat): string {
  const width = format.width;
  if (width?.type === "dxa" && width.value > 0) return `width:${width.value / TWIPS_PER_POINT}pt`;
  if (width?.type === "pct" && width.value > 0)
    return `width:${width.value / PCT_UNITS_PER_PERCENT}%`;
  return "";
}

/**
 * A table as Word lays it out: its preferred width and grid columns, and only
 * the borders the document defines (directly or through its table style) —
 * a table without borders shows none, as in Word with gridlines hidden.
 */
function renderTable(table: WmlTable, doc: Docx, styles: StyleResolver, block: number): string {
  const format = resolveTable(doc, table);
  const margins = format.cellMargins;
  const padding = [margins.top ?? 0, margins.right ?? 0, margins.bottom ?? 0, margins.left ?? 0]
    .map((twips) => `${twips / TWIPS_PER_POINT}pt`)
    .join(" ");
  const lastRow = table.rows.length - 1;
  const rows = table.rows
    .map((row, r) => {
      const lastCol = row.cells.length - 1;
      const cells = row.cells
        .map((cell, c) => {
          const coord = `${r},${c}`;
          const own = cellBorders(cell.tcPr);
          const outer = format.borders;
          const css = [
            `padding:${padding}`,
            `border-top:${borderCss(own.top ?? (r === 0 ? outer.top : outer.insideH))}`,
            `border-bottom:${borderCss(own.bottom ?? (r === lastRow ? outer.bottom : outer.insideH))}`,
            `border-left:${borderCss(own.left ?? (c === 0 ? outer.left : outer.insideV))}`,
            `border-right:${borderCss(own.right ?? (c === lastCol ? outer.right : outer.insideV))}`,
          ].join(";");
          const body = cell.paragraphs
            .map((p, para) => renderParagraph(p, styles, block, { coord, para }))
            .join("");
          return `<td class="wk-td" data-wk-block="${block}" data-wk-cell="${coord}" style="${escapeHtml(css)}">${body}</td>`;
        })
        .join("");
      return `<tr class="wk-tr">${cells}</tr>`;
    })
    .join("");
  const cols = format.columns.length
    ? `<colgroup>${format.columns.map((w) => `<col style="width:${w / TWIPS_PER_POINT}pt">`).join("")}</colgroup>`
    : "";
  const tableCss = [tableWidthCss(format), `margin-left:${format.leftEdge / TWIPS_PER_POINT}pt`]
    .filter(Boolean)
    .join(";");
  const styleAttr = ` style="${tableCss}"`;
  return `<table class="wk-table" data-wk-block="${block}"${styleAttr}>${cols}<tbody>${rows}</tbody></table>`;
}

function renderBlock(blockNode: WmlBlock, doc: Docx, styles: StyleResolver, block: number): string {
  switch (blockNode.kind) {
    case "paragraph":
      return renderParagraph(blockNode, styles, block);
    case "table":
      return renderTable(blockNode, doc, styles, block);
    default:
      // Raw / unmodelled block: show a non-editable marker; the library still
      // round-trips the underlying XML.
      return `<div class="wk-raw" data-wk-block="${block}" contenteditable="false">⟨preserved content⟩</div>`;
  }
}

// Drawing indices and z-order for the render pass in progress. Set for the
// duration of renderDocumentHtml so the per-run hook needs no extra parameter
// threaded through every block / paragraph / run function.
let drawingContext: DrawingRenderContext | undefined;

function pictureObjectsHtml(run: WmlRun, block: number, inline: number, cell?: CellAnchor): string {
  if (!drawingContext || !run.pieces.some((p) => p.kind === "drawing")) return "";
  const at = cell
    ? { block, cell: parseCellCoord(cell.coord), para: cell.para, inline }
    : { block, inline };
  return runDrawingsHtml(drawingContext, run, at);
}

function parseCellCoord(coord: string): { row: number; col: number } {
  const [row = 0, col = 0] = coord.split(",").map(Number);
  return { row, col };
}

/** Render the whole document body to an HTML string for the canvas. */
export function renderDocumentHtml(doc: Docx): string {
  const styles = createStyleResolver(doc);
  drawingContext = createDrawingContext(doc);
  try {
    return doc.document.body.blocks.map((b, i) => renderBlock(b, doc, styles, i)).join("");
  } finally {
    drawingContext = undefined;
  }
}

/** Plain-text extraction of a paragraph (used for tests / accessibility). */
export function paragraphPlainText(para: WmlParagraph): string {
  return paragraphText(para);
}
