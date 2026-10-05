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
  contentControlBlocks,
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
  fieldAttrs,
  fieldResultRuns,
  fieldRunRoles,
  fieldType,
  renderFieldStructureRun,
  renderSymbolRun,
  simpleFieldInstruction,
} from "./render-fields.js";
import { isMathElement, renderMath } from "./render-math.js";
import { noteMarksHtml, ownNoteMarkHtml } from "./render-notes.js";
import {
  type CellTextFormat,
  createStyleResolver,
  type ResolvedParagraphFormat,
  type ResolvedRunFormat,
} from "./resolve.js";
import {
  listLabelHtml,
  paragraphBoxCss,
  type RenderResolver,
  renderResolver,
  runEffectsCss,
  specialRunHtml,
  underlineStyleCss,
} from "./render-format.js";
import {
  createDrawingContext,
  type DrawingRenderContext,
  isDrawingOnlyRun,
  runDrawingsHtml,
} from "./render-drawing.js";
import { paragraphFloatsHtml, runObjectsHtml } from "./render-vml.js";
import type { DocPosition } from "./selection.js";
import { WML_NS } from "./wml-ns.js";
import { deletedRunText, type ReviewDecorations, reviewDecorations } from "./render-revisions.js";

const ALIGN_TO_CSS: Record<string, string> = {
  left: "left",
  start: "left",
  center: "center",
  right: "right",
  end: "right",
  both: "justify",
  distribute: "justify",
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
// Scripts Word formats with the complex-script properties (§17.3.2.7):
// Hebrew, Arabic, Syriac, Thaana, N'Ko, the Indic scripts, Thai, Lao and
// the Hebrew / Arabic presentation forms.
const COMPLEX_SCRIPT = /[\u0590-\u08FF\u0900-\u0DFF\u0E00-\u0EFF\uFB1D-\uFDFF\uFE70-\uFEFF]/;

/**
 * Whether a run is laid out with its complex-script formatting: it says so
 * (`w:cs` / `w:rtl`), or its text is in a complex script. A run mixing scripts
 * is one span on the canvas, so it takes the formatting of its script-specific
 * characters; Word would format each part separately.
 */
export function isComplexScriptRun(fmt: ResolvedRunFormat, text: string): boolean {
  return fmt.complexScript || fmt.rtl || COMPLEX_SCRIPT.test(text);
}

export function runCss(fmt: ResolvedRunFormat, complexScript = false): string {
  const bold = complexScript ? fmt.csBold : fmt.bold;
  const italic = complexScript ? fmt.csItalic : fmt.italic;
  const parts: string[] = [
    `font-weight:${bold ? "bold" : "normal"}`,
    `font-style:${italic ? "italic" : "normal"}`,
  ];
  if (fmt.rtl) parts.push("direction:rtl", "unicode-bidi:embed");
  const decoration: string[] = [];
  const underlined = fmt.underline !== undefined && fmt.underline !== "none";
  if (underlined) decoration.push("underline");
  const doubleStrike = fmt.toggles.has("dstrike");
  if (fmt.strike || doubleStrike) decoration.push("line-through");
  parts.push(`text-decoration-line:${decoration.length ? decoration.join(" ") : "none"}`);
  const underlineStyle = underlined ? underlineStyleCss(fmt.underline) : undefined;
  if (underlineStyle) parts.push(underlineStyle);
  else if (doubleStrike) parts.push("text-decoration-style:double");
  if (fmt.color && HEX_COLOR.test(fmt.color)) parts.push(`color:#${fmt.color}`);
  const highlight = fmt.highlight === undefined ? undefined : highlightCss(fmt.highlight);
  if (highlight) parts.push(`background-color:${highlight}`);
  parts.push(...fontCss(fmt, complexScript));
  const vertAlign = VERT_ALIGN_CSS[fmt.vertAlign ?? ""];
  if (vertAlign) parts.push(vertAlign);
  parts.push(...runEffectsCss(fmt));
  return parts.join(";");
}

const GENERIC_FAMILIES: ReadonlySet<string> = new Set(["serif", "sans-serif", "monospace"]);

const cssFontName = (font: string): string => `'${font.replace(/['"\\\n\r]/g, "")}'`;

function fontCss(fmt: ResolvedRunFormat, complexScript = false): string[] {
  const parts: string[] = [];
  const size = complexScript ? (fmt.csSizeHalfPoints ?? fmt.sizeHalfPoints) : fmt.sizeHalfPoints;
  if (size !== undefined) parts.push(`font-size:${size / 2}pt`);
  // The browser picks a font per glyph down the list, which approximates
  // Word's per-script fonts: Latin text in the Latin font (or its
  // metric-compatible stand-in), CJK text in the East Asian font,
  // complex-script text in the complex-script font.
  const primary = complexScript ? (fmt.csFont ?? fmt.font) : fmt.font;
  const others = [fmt.eastAsiaFont, complexScript ? undefined : fmt.csFont].filter(
    (f): f is string => !!f && f !== primary,
  );
  if (primary || others.length > 0) {
    // A generic family always matches, so it can only come last.
    const fallback = (primary && FALLBACK_FAMILY[primary]?.split(",")) || ["sans-serif"];
    const generic = fallback.filter((f) => GENERIC_FAMILIES.has(f));
    const list = [
      ...(primary ? [cssFontName(primary)] : []),
      ...fallback.filter((f) => !GENERIC_FAMILIES.has(f)),
      ...new Set(others.map(cssFontName)),
      ...(generic.length > 0 ? generic : ["sans-serif"]),
    ];
    parts.push(`font-family:${list.join(",")}`);
  }
  return parts;
}

export function paragraphCss(fmt: ResolvedParagraphFormat, mark: ResolvedRunFormat): string {
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
  css.push(...paragraphBoxCss(fmt));
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
  review: ReviewDecorations,
  cell?: CellAnchor,
  extraAttrs = "",
  extraClass = "",
): string {
  const deco = review.run(run);
  // A tracked deletion: shown (per the markup mode) but never editable.
  if (deco.deleted) {
    return `<span class="wk-del-run${deco.classes}" contenteditable="false"${deco.attrs} style="${escapeHtml(style)}">${escapeHtml(deletedRunText(run))}</span>`;
  }
  const text = runText(run);
  const attrs = [
    `data-wk-block="${block}"`,
    cellAttrs(cell),
    `data-wk-inline="${inline}"`,
    extraAttrs,
    style ? `style="${escapeHtml(style)}"` : "",
  ]
    .filter(Boolean)
    .join(" ");
  // Preserve whitespace/tabs; use a zero-width space for empty runs so the
  // caret has something to land on.
  return `${ownNoteMarkHtml(run)}<span class="wk-run${extraClass}${deco.classes}" ${attrs}${deco.attrs}>${specialRunHtml(run) ?? (escapeHtml(text) || "​")}</span>${noteMarksHtml(run)}`;
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
  const kind = inline.node.name.local === "hyperlink" ? "wk-link" : "wk-inline-raw";
  const instr = simpleFieldInstruction(inline.node);
  const field =
    instr === undefined ? "" : ` ${fieldAttrs({ type: fieldType(instr), instruction: instr })}`;
  const span = `<span class="${kind}" contenteditable="false"${field}>${escapeHtml(text)}</span>`;
  return `${text ? span : ""}${noteMarksHtml(inline)}`;
}

function renderParagraph(
  para: WmlParagraph,
  doc: Docx,
  styles: RenderResolver,
  review: ReviewDecorations,
  block: number,
  cell?: CellAnchor,
  text?: CellTextFormat,
): string {
  const labelText = styles.listLabel(para);
  const label =
    labelText === undefined
      ? ""
      : listLabelHtml(
          labelText,
          styles.paragraph(para, text),
          runCss(styles.listLabelRun(para, text)),
        );
  const styleAttr = ` style="${escapeHtml(paragraphCss(styles.paragraph(para, text), styles.run(para, undefined, text)))}"`;
  // `data-wk-inline` counts runs only (the unit `runAtPath` resolves), so
  // raw inlines are interleaved without consuming an index.
  let runIndex = 0;
  let textRuns = 0;
  const fields = fieldResultRuns(para);
  // Fields, symbols and equations (Insert tab): see render-fields / render-math.
  const fieldRoles = fieldRunRoles(para);
  let mathIndex = 0;
  const mathAnchor = (): string =>
    `data-wk-math-block="${block}" data-wk-math="${mathIndex++}"${cell ? ` data-wk-math-cell="${cell.coord}" data-wk-math-para="${cell.para}"` : ""}`;
  let inner = para.children
    .map((child, i) => {
      if (child.kind !== "run") {
        review.inline(child.node);
        return isMathElement(child.node)
          ? renderMath(child.node, mathAnchor())
          : renderRawInline(child);
      }
      const role = fieldRoles.get(child);
      const fmt = styles.run(para, child, text);
      const css = runCss(fmt, isComplexScriptRun(fmt, runText(child)));
      const special = renderFieldStructureRun(child, role) ?? renderSymbolRun(child, css);
      if (special !== undefined) {
        runIndex++;
        return special;
      }
      const inline = runIndex++;
      // A picture-only run gets no editable span: typing there could not be
      // written back into the run, so the caret lives in the text around it.
      const objects =
        pictureObjectsHtml(child, block, inline, cell) +
        runObjectsHtml(child, doc, positionOf(block, inline, cell));
      if (isDrawingOnlyRun(child)) return objects;
      textRuns++;
      const field = fields.get(i);
      return (
        renderRun(
          child,
          css,
          block,
          inline,
          review,
          cell,
          field ? fieldAttrs(field) : "",
          role === "result" ? " wk-fresult" : "",
        ) + objects
      );
    })
    .join("");
  if (textRuns === 0) inner = `​${inner}`;
  // Floating objects go first so their static position is the paragraph's top.
  inner = label + paragraphFloatsHtml(para, doc, (i) => positionOf(block, i, cell)) + inner;
  const cellAttr = cell ? ` ${cellAttrs(cell)}` : "";
  const deco = review.paragraph(para);
  return `<p class="wk-p${deco.classes}" data-wk-block="${block}"${cellAttr}${styleAttr}${deco.attrs}>${inner}</p>`;
}

function renderBlock(
  blockNode: WmlBlock,
  doc: Docx,
  styles: RenderResolver,
  review: ReviewDecorations,
  block: number,
): string {
  switch (blockNode.kind) {
    case "paragraph":
      return renderParagraph(blockNode, doc, styles, review, block);
    case "table":
      return renderTable(blockNode, {
        doc,
        block,
        paragraph: (para, anchor, text) =>
          renderParagraph(para, doc, styles, review, block, anchor, text),
      });
    default: {
      // A block content control (cover page, watermark, TOC) shows its content,
      // read-only: the blocks are a parsed copy, so edits there could not be saved.
      const inner = contentControlBlocks(blockNode);
      if (inner) {
        const html = inner.map((b) => renderBlock(b, doc, styles, review, block)).join("");
        return `<div class="wk-sdt" data-wk-block="${block}" contenteditable="false">${html}</div>`;
      }
      // Raw / unmodelled block: show a non-editable marker; the library still
      // round-trips the underlying XML.
      return `<div class="wk-raw" data-wk-block="${block}" contenteditable="false">⟨preserved content⟩</div>`;
    }
  }
}

// Drawing indices and z-order for the render pass in progress. Set for the
// duration of renderDocumentHtml so the per-run hook needs no extra parameter
// threaded through every block / paragraph / run function.
let drawingContext: DrawingRenderContext | undefined;

function pictureObjectsHtml(run: WmlRun, block: number, inline: number, cell?: CellAnchor): string {
  if (!drawingContext || !run.pieces.some((p) => p.kind === "drawing")) return "";
  return runDrawingsHtml(drawingContext, run, positionOf(block, inline, cell));
}

function positionOf(block: number, inline: number, cell?: CellAnchor): DocPosition {
  if (!cell) return { block, inline };
  const [row = 0, col = 0] = cell.coord.split(",").map(Number);
  return { block, cell: { row, col }, para: cell.para, inline };
}

/** Render the whole document body to an HTML string for the canvas. */
export function renderDocumentHtml(doc: Docx): string {
  drawingContext = createDrawingContext(doc);
  try {
    return renderBlocksHtml(doc, doc.document.body.blocks);
  } finally {
    drawingContext = undefined;
  }
}

/**
 * Render a block list — the body's or a story's — one top-level element per
 * block, each tagged with its `data-wk-block` index.
 */
export function renderBlocksHtml(doc: Docx, blocks: readonly WmlBlock[]): string {
  const styles = renderResolver(doc, createStyleResolver(doc));
  const review = reviewDecorations(doc);
  return blocks.map((b, i) => renderBlock(b, doc, styles, review, i)).join("");
}

/** Plain-text extraction of a paragraph (used for tests / accessibility). */
export function paragraphPlainText(para: WmlParagraph): string {
  return paragraphText(para);
}
