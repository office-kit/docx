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
  type XmlElement,
} from "@office-kit/docx";
import { highlightCss } from "./highlight.js";
import { type CellAnchor, renderTable } from "./render-table.js";
import {
  type CellTextFormat,
  createStyleResolver,
  type ResolvedParagraphFormat,
  type ResolvedRunFormat,
  type StyleResolver,
} from "./resolve.js";
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
  text?: CellTextFormat,
): string {
  const styleAttr = ` style="${escapeHtml(paragraphCss(styles.paragraph(para, text), styles.run(para, undefined, text)))}"`;
  // `data-wk-inline` counts runs only (the unit `runAtPath` resolves), so
  // raw inlines are interleaved without consuming an index.
  let runIndex = 0;
  let inner = para.children
    .map((child) =>
      child.kind === "run"
        ? renderRun(child, runCss(styles.run(para, child, text)), block, runIndex++, cell)
        : renderRawInline(child),
    )
    .join("");
  if (runIndex === 0) inner = `​${inner}`;
  const cellAttr = cell ? ` ${cellAttrs(cell)}` : "";
  return `<p class="wk-p" data-wk-block="${block}"${cellAttr}${styleAttr}>${inner}</p>`;
}

function renderBlock(blockNode: WmlBlock, doc: Docx, styles: StyleResolver, block: number): string {
  switch (blockNode.kind) {
    case "paragraph":
      return renderParagraph(blockNode, styles, block);
    case "table":
      return renderTable(blockNode, {
        doc,
        block,
        paragraph: (para, anchor, text) => renderParagraph(para, styles, block, anchor, text),
      });
    default:
      // Raw / unmodelled block: show a non-editable marker; the library still
      // round-trips the underlying XML.
      return `<div class="wk-raw" data-wk-block="${block}" contenteditable="false">⟨preserved content⟩</div>`;
  }
}

/** Render the whole document body to an HTML string for the canvas. */
export function renderDocumentHtml(doc: Docx): string {
  const styles = createStyleResolver(doc);
  return doc.document.body.blocks.map((b, i) => renderBlock(b, doc, styles, i)).join("");
}

/** Plain-text extraction of a paragraph (used for tests / accessibility). */
export function paragraphPlainText(para: WmlParagraph): string {
  return paragraphText(para);
}
