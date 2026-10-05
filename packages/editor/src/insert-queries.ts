/**
 * Read-only queries the Insert-tab dialogs need: the selected text (Link's
 * "Text to display"), the link at the caret (Edit / Remove Link), and the
 * targets a link or cross-reference can point at (headings, bookmarks, notes,
 * numbered items, captions).
 */

import {
  bookmarks,
  complexFields,
  type Docx,
  equationLinear,
  getParagraphNumbering,
  getParagraphStyle,
  type HyperlinkInfo,
  paragraphEquations,
  paragraphHyperlinks,
  paragraphText,
  runTextLength,
  type WmlParagraph,
  type WmlRun,
  type XmlElement,
} from "@office-kit/docx";
import { absoluteOffset } from "./char-offset.js";
import { paragraphAt, paragraphsInRange } from "./doc-access.js";
import { parseCell } from "./dom-selection.js";
import type { EditorModel } from "./model.js";
import { type DocPosition, orderSelection } from "./selection.js";

/** The selected text (paragraphs joined by newlines); empty for a caret. */
export function selectionText(model: EditorModel): string {
  const sel = model.selection;
  if (!sel) return "";
  const ordered = orderSelection(sel);
  if (ordered.collapsed) return "";
  const paras = paragraphsInRange(model.doc, ordered);
  return paras
    .map((p, i) => {
      const text = paragraphText(p);
      const from = i === 0 ? absoluteOffset(p, ordered.start) : 0;
      const to = i === paras.length - 1 ? absoluteOffset(p, ordered.end) : text.length;
      return text.slice(from, to);
    })
    .join("\n");
}

/**
 * The hyperlink at the caret: one the selection covers, or one directly
 * beside a collapsed caret. `index` is its position among the paragraph's
 * links (what the Edit / Remove Link commands take).
 */
export function hyperlinkAtCaret(
  model: EditorModel,
): { index: number; link: HyperlinkInfo } | undefined {
  const sel = model.selection;
  if (!sel) return undefined;
  const { start, end } = orderSelection(sel);
  const para = paragraphAt(model.doc, start);
  if (!para) return undefined;
  const links = paragraphHyperlinks(model.doc, para);
  if (links.length === 0) return undefined;
  const from = absoluteOffset(para, start);
  const to =
    paragraphAt(model.doc, end) === para ? absoluteOffset(para, end) : Number.POSITIVE_INFINITY;
  const offsets = rawInlineOffsets(para);
  for (const [index, link] of links.entries()) {
    const at = offsets.get(link.element);
    if (at !== undefined && at >= from && at <= to) return { index, link };
  }
  return undefined;
}

/**
 * Where each raw inline (a `w:hyperlink` …) sits in the paragraph's caret
 * offsets. Offsets count run text only, so a link occupies a single point
 * between the runs around it.
 */
function rawInlineOffsets(para: WmlParagraph): Map<XmlElement, number> {
  const out = new Map<XmlElement, number>();
  let cursor = 0;
  for (const c of para.children) {
    if (c.kind === "run") cursor += runTextLength(c);
    else out.set(c.node, cursor);
  }
  return out;
}

export interface DocumentTarget {
  /** What the dialog lists. */
  readonly label: string;
  readonly paragraph: WmlParagraph;
  /** Body block index of the paragraph (Go To / caret placement). */
  readonly block: number;
}

export interface HeadingTarget extends DocumentTarget {
  readonly level: number;
}

const HEADING_STYLE = /^Heading(\d)$/;

/** Top-level paragraphs in a Heading style, in order. */
export function documentHeadings(doc: Docx): HeadingTarget[] {
  const out: HeadingTarget[] = [];
  for (const [block, b] of doc.document.body.blocks.entries()) {
    if (b.kind !== "paragraph") continue;
    const m = HEADING_STYLE.exec(getParagraphStyle(b) ?? "");
    if (!m?.[1]) continue;
    out.push({ label: paragraphText(b), paragraph: b, block, level: Number(m[1]) });
  }
  return out;
}

export interface BookmarkTarget extends DocumentTarget {
  readonly name: string;
  readonly hidden: boolean;
}

/** Bookmarks in document order; hidden ones (`_Ref…`, `_Toc…`) are flagged. */
export function documentBookmarks(doc: Docx): BookmarkTarget[] {
  const blockOf = new Map(doc.document.body.blocks.map((b, i) => [b, i] as const));
  return bookmarks(doc).map((b) => ({
    label: b.name,
    name: b.name,
    hidden: b.name.startsWith("_"),
    paragraph: b.paragraph,
    block: blockOf.get(b.paragraph) ?? 0,
  }));
}

export interface NoteTarget extends DocumentTarget {
  readonly run: WmlRun;
  readonly number: number;
}

/** Footnote or endnote reference marks in the body, numbered as Word shows them. */
export function documentNotes(doc: Docx, kind: "footnote" | "endnote"): NoteTarget[] {
  const local = `${kind}Reference`;
  const out: NoteTarget[] = [];
  for (const [block, b] of doc.document.body.blocks.entries()) {
    if (b.kind !== "paragraph") continue;
    for (const c of b.children) {
      if (
        c.kind !== "run" ||
        !c.pieces.some((p) => p.kind === "raw" && p.node.name.local === local)
      )
        continue;
      const number = out.length + 1;
      const context = paragraphText(b).trim().slice(0, CONTEXT_LENGTH);
      out.push({ label: `${number} ${context}`, paragraph: b, block, run: c, number });
    }
  }
  return out;
}

// Characters of paragraph text shown beside a note number in the list.
const CONTEXT_LENGTH = 50;

/** Paragraphs in a list (numbered items). */
export function documentNumberedItems(doc: Docx): DocumentTarget[] {
  const out: DocumentTarget[] = [];
  for (const [block, b] of doc.document.body.blocks.entries()) {
    if (b.kind === "paragraph" && getParagraphNumbering(b)) {
      out.push({ label: paragraphText(b), paragraph: b, block });
    }
  }
  return out;
}

export interface EquationRef {
  /** A position in the equation's paragraph (what `editEquationCommand` takes). */
  readonly at: DocPosition;
  /** Index among that paragraph's equations. */
  readonly index: number;
  /** The equation in linear format, for the edit dialog. */
  readonly linear: string;
}

/**
 * The equation a rendered `.wk-math` element shows (double-click to edit),
 * from the `data-wk-math-*` anchors `renderDocumentHtml` writes.
 */
export function equationAtElement(model: EditorModel, el: Element): EquationRef | undefined {
  const block = Number(el.getAttribute("data-wk-math-block"));
  const index = Number(el.getAttribute("data-wk-math"));
  if (!Number.isInteger(block) || !Number.isInteger(index)) return undefined;
  const cell = parseCell(el.getAttribute("data-wk-math-cell"));
  const para = Number(el.getAttribute("data-wk-math-para"));
  const at: DocPosition = {
    block,
    ...(cell ? { cell, para: Number.isInteger(para) ? para : 0 } : {}),
  };
  const paragraph = paragraphAt(model.doc, at);
  const element = paragraph ? paragraphEquations(paragraph)[index] : undefined;
  return element ? { at, index, linear: equationLinear(element) } : undefined;
}

/** Caption paragraphs: those holding a `SEQ <label>` field (Figure, Table, Equation). */
export function documentCaptions(doc: Docx, label: string): DocumentTarget[] {
  const blockOf = new Map(doc.document.body.blocks.map((b, i) => [b, i] as const));
  const wanted = label.toUpperCase();
  const seen = new Set<WmlParagraph>();
  const out: DocumentTarget[] = [];
  for (const f of complexFields(doc)) {
    if (f.type !== "SEQ" || seen.has(f.paragraph)) continue;
    const id = f.instruction.trim().split(/\s+/)[1]?.toUpperCase();
    if (id !== wanted) continue;
    seen.add(f.paragraph);
    out.push({
      label: paragraphText(f.paragraph),
      paragraph: f.paragraph,
      block: blockOf.get(f.paragraph) ?? 0,
    });
  }
  return out;
}
