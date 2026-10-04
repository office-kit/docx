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
  getParagraphAlignment,
  getParagraphStyle,
  getRunFormat,
  paragraphText,
  type RunFormatting,
  type WmlBlock,
  type WmlInline,
  type WmlParagraph,
  type WmlRun,
  type WmlTable,
  type XmlElement,
} from "@office-kit/docx";
import { highlightCss } from "./highlight.js";
import { WML_NS } from "./wml-ns.js";

/**
 * Visual formatting for the built-in paragraph styles so headings read as
 * headings on the canvas (Word/Google-Docs-like hierarchy). This is display
 * only — the underlying `pStyle` is what round-trips; here we just make the
 * document look like a document instead of a flat wall of text.
 */
const STYLE_LOOK: Record<string, string> = {
  Title: "font-size:26pt;font-weight:700;margin:0 0 8px;line-height:1.15",
  Subtitle: "font-size:15pt;color:#666;margin:0 0 12px",
  Heading1: "font-size:20pt;font-weight:600;color:#1a1a1a;margin:18px 0 6px;line-height:1.2",
  Heading2: "font-size:16pt;font-weight:600;color:#1a1a1a;margin:14px 0 4px;line-height:1.2",
  Heading3: "font-size:13pt;font-weight:600;color:#333;margin:12px 0 4px",
  Heading4: "font-size:12pt;font-weight:600;font-style:italic;color:#333;margin:10px 0 4px",
  Heading5: "font-size:11pt;font-weight:600;color:#444;margin:10px 0 4px",
  Heading6: "font-size:11pt;font-weight:600;font-style:italic;color:#555;margin:10px 0 4px",
  Quote: "font-style:italic;color:#555;border-left:3px solid #ccc;padding-left:12px;margin:8px 0",
};

const ALIGN_TO_CSS: Record<string, string> = {
  left: "left",
  center: "center",
  right: "right",
  both: "justify",
  distribute: "justify",
};

function escapeHtml(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

// Values come straight from the file (`getRunFormat` does not validate), and a
// loaded .docx is untrusted: only schema-valid colors become CSS, and font
// names lose the characters that could end the CSS string.
const HEX_COLOR = /^[0-9A-Fa-f]{6}$/;

/** Inline CSS for a run's direct formatting. */
function runStyle(fmt: RunFormatting): string {
  const parts: string[] = [];
  if (fmt.bold) parts.push("font-weight:bold");
  if (fmt.italic) parts.push("font-style:italic");
  const decoration: string[] = [];
  if (fmt.underline && fmt.underline !== "none") decoration.push("underline");
  if (fmt.strike) decoration.push("line-through");
  if (decoration.length) parts.push(`text-decoration:${decoration.join(" ")}`);
  if (fmt.color && HEX_COLOR.test(fmt.color)) parts.push(`color:#${fmt.color}`);
  const highlight = fmt.highlight === undefined ? undefined : highlightCss(fmt.highlight);
  if (highlight) parts.push(`background-color:${highlight}`);
  if (fmt.fontSizeHalfPoints) parts.push(`font-size:${fmt.fontSizeHalfPoints / 2}pt`);
  if (fmt.font) parts.push(`font-family:'${fmt.font.replace(/['"\\\n\r]/g, "")}'`);
  return parts.join(";");
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

function renderRun(run: WmlRun, block: number, inline: number, cell?: CellAnchor): string {
  const fmt = getRunFormat(run);
  const style = runStyle(fmt);
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

function renderParagraph(para: WmlParagraph, block: number, cell?: CellAnchor): string {
  const align = getParagraphAlignment(para);
  const alignCss = align ? ALIGN_TO_CSS[align] : undefined;
  const look = STYLE_LOOK[getParagraphStyle(para) ?? ""];
  const css = [look, alignCss ? `text-align:${alignCss}` : ""].filter(Boolean).join(";");
  const styleAttr = css ? ` style="${css}"` : "";
  // `data-wk-inline` counts runs only (the unit `runAtPath` resolves), so
  // raw inlines are interleaved without consuming an index.
  let runIndex = 0;
  let inner = para.children
    .map((child) =>
      child.kind === "run" ? renderRun(child, block, runIndex++, cell) : renderRawInline(child),
    )
    .join("");
  if (runIndex === 0) inner = `​${inner}`;
  const cellAttr = cell ? ` ${cellAttrs(cell)}` : "";
  return `<p class="wk-p" data-wk-block="${block}"${cellAttr}${styleAttr}>${inner}</p>`;
}

function renderTable(table: WmlTable, block: number): string {
  const rows = table.rows
    .map((row, r) => {
      const cells = row.cells
        .map((cell, c) => {
          const coord = `${r},${c}`;
          const body = cell.paragraphs
            .map((p, para) => renderParagraph(p, block, { coord, para }))
            .join("");
          return `<td class="wk-td" data-wk-block="${block}" data-wk-cell="${coord}">${body}</td>`;
        })
        .join("");
      return `<tr class="wk-tr">${cells}</tr>`;
    })
    .join("");
  return `<table class="wk-table" data-wk-block="${block}"><tbody>${rows}</tbody></table>`;
}

function renderBlock(blockNode: WmlBlock, block: number): string {
  switch (blockNode.kind) {
    case "paragraph":
      return renderParagraph(blockNode, block);
    case "table":
      return renderTable(blockNode, block);
    default:
      // Raw / unmodelled block: show a non-editable marker; the library still
      // round-trips the underlying XML.
      return `<div class="wk-raw" data-wk-block="${block}" contenteditable="false">⟨preserved content⟩</div>`;
  }
}

/** Render the whole document body to an HTML string for the canvas. */
export function renderDocumentHtml(doc: { document: { body: { blocks: WmlBlock[] } } }): string {
  return doc.document.body.blocks.map((b, i) => renderBlock(b, i)).join("");
}

/** Plain-text extraction of a paragraph (used for tests / accessibility). */
export function paragraphPlainText(para: WmlParagraph): string {
  return paragraphText(para);
}
