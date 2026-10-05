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
  wrappedRuns,
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
import { documentSections } from "./layout/sections.js";
import {
  eastAsianTypographyCss,
  gridLetterSpacing,
  lineHeightCss,
  type GridPitch,
  gridPitch,
  paragraphMetrics,
} from "./render-grid.js";
import {
  combinedTextHtml,
  listLabelHtml,
  paragraphBoxCss,
  type RenderResolver,
  renderResolver,
  runEffectsCss,
  specialRunHtml,
  tabStopsAttr,
  textHtml,
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

const JA_SERIF = "YuMincho,'Hiragino Mincho ProN','Noto Serif JP','Noto Serif CJK JP',serif";
const JA_SANS = "YuGothic,'Hiragino Sans','Noto Sans JP','Noto Sans CJK JP',sans-serif";
const ZH_SERIF = "'Songti SC','Noto Serif SC','Noto Serif CJK SC',serif";
const ZH_SANS = "'PingFang SC','Noto Sans SC','Noto Sans CJK SC',sans-serif";

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
  // Japanese Word writes these fonts under their Japanese names. macOS has
  // YuMincho / YuGothic and Hiragino, other systems the Noto CJK fonts.
  "Yu Mincho": JA_SERIF,
  游明朝: JA_SERIF,
  "游明朝 Demibold": JA_SERIF,
  "MS Mincho": JA_SERIF,
  "ＭＳ 明朝": JA_SERIF,
  "ＭＳ Ｐ明朝": JA_SERIF,
  "Hiragino Mincho ProN": JA_SERIF,
  "ヒラギノ明朝 ProN": JA_SERIF,
  "Hiragino Kaku Gothic ProN": JA_SANS,
  "ヒラギノ角ゴ ProN": JA_SANS,
  "Yu Gothic": JA_SANS,
  "Yu Gothic Light": JA_SANS,
  游ゴシック: JA_SANS,
  "游ゴシック Light": JA_SANS,
  "MS Gothic": JA_SANS,
  "ＭＳ ゴシック": JA_SANS,
  "ＭＳ Ｐゴシック": JA_SANS,
  Meiryo: JA_SANS,
  メイリオ: JA_SANS,
  SimSun: ZH_SERIF,
  宋体: ZH_SERIF,
  DengXian: ZH_SANS,
  等线: ZH_SANS,
  "Microsoft YaHei": ZH_SANS,
  微软雅黑: ZH_SANS,
};

const TWIPS_PER_POINT = 20;
// w:w, a run at its natural width (§17.3.2.43).
const FULL_SCALE = 100;
// w:sz when nothing sets it (§17.3.2.38).
const DEFAULT_SIZE_HALF_POINTS = 20;

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
    const list: string[] = primary ? [cssFontName(primary)] : [];
    for (const f of fallback) if (!GENERIC_FAMILIES.has(f)) list.push(f);
    // An East Asian font missing from the system needs its stand-ins too,
    // or CJK text falls through to the Latin font's generic family.
    for (const font of others) {
      list.push(cssFontName(font));
      for (const f of FALLBACK_FAMILY[font]?.split(",") ?? []) {
        if (!GENERIC_FAMILIES.has(f)) list.push(f);
      }
    }
    list.push(...(generic.length > 0 ? generic : ["sans-serif"]));
    parts.push(`font-family:${[...new Set(list)].join(",")}`);
  }
  return parts;
}

const HORIZONTAL_SIDES = { before: "top", after: "bottom", left: "left", right: "right" } as const;
const VERTICAL_SIDES = { before: "right", after: "left", left: "top", right: "bottom" } as const;

export function paragraphCss(
  fmt: ResolvedParagraphFormat,
  mark: ResolvedRunFormat,
  pitch: GridPitch = {},
  vertical = false,
): string {
  const css = fontCss(mark);
  const fontSize = (mark.sizeHalfPoints ?? DEFAULT_SIZE_HALF_POINTS) / 2;
  const align = ALIGN_TO_CSS[fmt.alignment ?? ""];
  if (align) css.push(`text-align:${align}`);
  // Distributed text spreads every line, the last one too, by character.
  if (fmt.alignment === "distribute" || fmt.alignment === "thaiDistribute") {
    css.push("text-align-last:justify", "text-justify:inter-character");
  }
  const m = paragraphMetrics(fmt, pitch, fontSize);
  // In vertical text (tbRl) lines follow each other right to left, so space
  // before is on the right and the left indent at the top. Physical sides
  // rather than logical ones: a bidi paragraph's own direction must not flip them.
  const side = vertical ? VERTICAL_SIDES : HORIZONTAL_SIDES;
  css.push(`margin-${side.before}:${m.before}pt`, `margin-${side.after}:${m.after}pt`);
  if (m.left !== undefined) css.push(`margin-${side.left}:${m.left}pt`);
  if (m.right !== undefined) css.push(`margin-${side.right}:${m.right}pt`);
  if (m.hanging !== undefined) css.push(`text-indent:${-m.hanging}pt`);
  else if (m.firstLine !== undefined) css.push(`text-indent:${m.firstLine}pt`);
  const lineHeight = lineHeightCss(fmt, pitch, mark);
  if (lineHeight) css.push(lineHeight);
  else if (fmt.line !== undefined && fmt.line > 0 && fmt.lineRule !== "auto") {
    css.push(`line-height:${fmt.line / TWIPS_PER_POINT}pt`);
  }
  const spacing = gridLetterSpacing(fmt, pitch, fontSize);
  if (spacing !== undefined) css.push(`letter-spacing:${spacing}pt`);
  css.push(...eastAsianTypographyCss(fmt));
  css.push(...paragraphBoxCss(fmt));
  if (fmt.dropCap) css.push(...dropCapCss(fmt.dropCap));
  return css.join(";");
}

/**
 * A drop cap floats at the start of the next paragraph's first lines; in the
 * margin it takes no room from them and sits `hSpace` left of the text.
 */
function dropCapCss(dropCap: NonNullable<ResolvedParagraphFormat["dropCap"]>): string[] {
  const space = dropCap.hSpace / TWIPS_PER_POINT;
  return dropCap.kind === "drop"
    ? ["float:left", `margin-right:${space}pt`]
    : [
        "float:left",
        "width:0",
        "display:flex",
        "justify-content:flex-end",
        "position:relative",
        `left:${-space}pt`,
      ];
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
  innerHtml?: string,
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
  return `${ownNoteMarkHtml(run)}<span class="wk-run${extraClass}${deco.classes}" ${attrs}${deco.attrs}>${innerHtml ?? specialRunHtml(run) ?? (textHtml(text) || "​")}</span>${noteMarksHtml(run)}`;
}

/**
 * Visible text of an unmodelled inline (`<w:hyperlink>`, a field run built in
 * memory, `<w:ins>`, `<w:sdt>`…): its `<w:t>` and `<w:tab>` descendants. `<w:instrText>`
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
  // Only reached inside runs: a paragraph's tab stops (`w:tabs/w:tab`) sit in pPr.
  if (el.name.uri === WML_NS && el.name.local === "tab") return "\t";
  let out = "";
  for (const child of el.children) if (child.kind === "element") out += rawVisibleText(child);
  return out;
}

/**
 * Render an unmodelled inline read-only. It carries no `data-wk-*` anchor and
 * is not a `.wk-run`, so typing never writes into it and the canvas text sync
 * leaves it alone; the library round-trips the XML untouched.
 */
function renderRawInline(
  inline: Extract<WmlInline, { kind: "raw" }>,
  runStyle: (run: WmlRun) => string,
): string {
  // A hyperlink looks like its runs (Word's Hyperlink style, or none in a
  // table of contents), so wrapped runs keep their own formatting.
  const runs = wrappedRuns(inline);
  const text = runs ? runs.map(runText).join("") : rawVisibleText(inline.node);
  const kind = inline.node.name.local === "hyperlink" ? "wk-link" : "wk-inline-raw";
  const instr = simpleFieldInstruction(inline.node);
  const field =
    instr === undefined ? "" : ` ${fieldAttrs({ type: fieldType(instr), instruction: instr })}`;
  const span = `<span class="${kind}" contenteditable="false"${field}>${runs ? runs.map((run) => wrappedRunHtml(run, runStyle)).join("") : textHtml(text)}</span>`;
  return `${text ? span : ""}${noteMarksHtml(inline)}`;
}

function wrappedRunHtml(run: WmlRun, runStyle: (run: WmlRun) => string): string {
  const text = runText(run);
  return text ? `<span style="${escapeHtml(runStyle(run))}">${textHtml(text)}</span>` : "";
}

/**
 * Characters per Fit Text region: runs sharing a `w:fitText/@w:id` are one
 * region, spread together over its width (§17.3.2.14).
 */
function fitTextRegionChars(
  para: WmlParagraph,
  styles: RenderResolver,
  text: CellTextFormat | undefined,
): Map<number, number> {
  const out = new Map<number, number>();
  for (const child of para.children) {
    if (child.kind !== "run") continue;
    const id = styles.run(para, child, text).fitText?.id;
    if (id !== undefined) out.set(id, (out.get(id) ?? 0) + [...runText(child)].length);
  }
  return out;
}

/** A run's part of its Fit Text region's width, by its share of the characters. */
function fitTextShare(
  fmt: ResolvedRunFormat,
  run: WmlRun,
  regionChars: ReadonlyMap<number, number>,
): ResolvedRunFormat {
  const fit = fmt.fitText;
  const total = fit?.id === undefined ? undefined : regionChars.get(fit.id);
  if (!fit || !total) return fmt;
  const width = (fit.width * [...runText(run)].length) / total;
  return { ...fmt, fitText: { ...fit, width } };
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
  const paragraphFmt = styles.paragraph(para, text);
  const labelText = styles.listLabel(para);
  const label =
    labelText === undefined
      ? ""
      : listLabelHtml(labelText, paragraphFmt, runCss(styles.listLabelRun(para, text)));
  const { pitch, vertical } = sectionFlowOf(block);
  const styleAttr = ` style="${escapeHtml(paragraphCss(paragraphFmt, styles.run(para, undefined, text), pitch, vertical))}"`;
  // `data-wk-inline` counts runs only (the unit `runAtPath` resolves), so
  // raw inlines are interleaved without consuming an index.
  let runIndex = 0;
  let textRuns = 0;
  const fields = fieldResultRuns(para);
  // Fields, symbols and equations (Insert tab): see render-fields / render-math.
  const fieldRoles = fieldRunRoles(para);
  const fitTextChars = fitTextRegionChars(para, styles, text);
  let mathIndex = 0;
  const mathAnchor = (): string =>
    `data-wk-math-block="${block}" data-wk-math="${mathIndex++}"${cell ? ` data-wk-math-cell="${cell.coord}" data-wk-math-para="${cell.para}"` : ""}`;
  let inner = para.children
    .map((child, i) => {
      if (child.kind !== "run") {
        review.inline(child.node);
        return isMathElement(child.node)
          ? renderMath(child.node, mathAnchor())
          : renderRawInline(child, (run) => {
              const fmt = styles.run(para, run, text);
              return runCss(fmt, isComplexScriptRun(fmt, runText(run)));
            });
      }
      const role = fieldRoles.get(child);
      const fmt = fitTextShare(styles.run(para, child, text), child, fitTextChars);
      const lineHeight = lineHeightCss(paragraphFmt, pitch, fmt);
      const css =
        runCss(fmt, isComplexScriptRun(fmt, runText(child))) + (lineHeight ? `;${lineHeight}` : "");
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
      const scaled =
        fmt.scale !== undefined && fmt.scale !== FULL_SCALE ? ` data-wk-scale="${fmt.scale}"` : "";
      const combined = fmt.eastAsianLayout?.combine
        ? combinedTextHtml(runText(child), fmt.eastAsianLayout)
        : undefined;
      return (
        renderRun(
          child,
          css,
          block,
          inline,
          review,
          cell,
          (field ? fieldAttrs(field) : "") + scaled,
          role === "result" ? " wk-fresult" : "",
          combined,
        ) + objects
      );
    })
    .join("");
  if (textRuns === 0) inner = `​${inner}`;
  // Floating objects go first so their static position is the paragraph's top.
  inner = label + paragraphFloatsHtml(para, doc, (i) => positionOf(block, i, cell)) + inner;
  const cellAttr = cell ? ` ${cellAttrs(cell)}` : "";
  const tabsAttr = inner.includes('class="wk-tab"')
    ? tabStopsAttr(paragraphFmt, styles.defaultTabStop)
    : "";
  // The canvas puts a drop cap with the paragraph it drops into (it floats there).
  const dropCapAttr = paragraphFmt.dropCap ? " data-wk-dropcap" : "";
  const deco = review.paragraph(para);
  return `<p class="wk-p${deco.classes}" data-wk-block="${block}"${cellAttr}${styleAttr}${tabsAttr}${dropCapAttr}${deco.attrs}>${inner}</p>`;
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
  const sections = documentSections(doc);
  const base = baseFontSize(doc);
  const flowOfBlock: SectionFlow[] = [];
  for (const s of sections) {
    const flow = { pitch: gridPitch(s.grid, base), vertical: s.vertical };
    for (let b = s.firstBlock; b <= s.lastBlock; b++) flowOfBlock[b] = flow;
  }
  sectionFlowOf = (block) => flowOfBlock[block] ?? HORIZONTAL;
  try {
    return renderBlocksHtml(doc, doc.document.body.blocks);
  } finally {
    drawingContext = undefined;
    sectionFlowOf = () => HORIZONTAL;
  }
}

/** What a block takes from its section: the grid, and the text direction. */
interface SectionFlow {
  readonly pitch: GridPitch;
  readonly vertical: boolean;
}
const HORIZONTAL: SectionFlow = { pitch: {}, vertical: false };
// The section flow of a body block, for the render in progress. Stories
// (headers, notes) are rendered on their own: no grid, horizontal.
let sectionFlowOf: (block: number) => SectionFlow = () => HORIZONTAL;

/** The Normal style's font size in points: the size a character grid is built on. */
function baseFontSize(doc: Docx): number {
  const size = createStyleResolver(doc).style("Normal").run.sizeHalfPoints;
  return (size ?? DEFAULT_SIZE_HALF_POINTS) / 2;
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
