/**
 * Recording tracked changes while the user edits (Review ▸ Track Changes).
 *
 * The structural commands (typing, deleting, Enter, Backspace at a paragraph
 * start) and the formatting paths call into here when tracking is on; the
 * revisions themselves are written by `@office-kit/docx` (`insertTrackedText`,
 * `deleteTrackedText`, `trackParagraphMark`, `trackRunFormatChange`,
 * `trackParagraphFormatChange`).
 */

import {
  deleteTrackedText,
  documentProtection,
  getDocumentSetting,
  insertTrackedText,
  isolateParagraphRunRange,
  paragraphMarkRevision,
  runTextLength,
  trackParagraphFormatChange,
  trackParagraphMark,
  trackRunFormatChange,
  type RevisionAuthor,
  type WmlParagraph,
  type WmlRun,
  type XmlElement,
  type XmlNode,
} from "@office-kit/docx";
import { absoluteOffset, positionAtOffset } from "./char-offset.js";
import { cellAt, paragraphAt, paragraphsInRange } from "./doc-access.js";
import type { EditorModel } from "./model.js";
import { caretAt, type DocPosition, orderSelection } from "./selection.js";
import { WML_NS } from "./wml-ns.js";

/** The person edits are attributed to (Word ▸ Preferences ▸ User Information). */
export interface Reviewer {
  readonly author: string;
  readonly initials: string;
}

const DEFAULT_REVIEWER: Reviewer = { author: "Author", initials: "A" };
const reviewers = new WeakMap<EditorModel, Reviewer>();

/** Set who this editor's tracked changes and comments are attributed to. */
export function setReviewer(model: EditorModel, reviewer: Reviewer): void {
  reviewers.set(model, reviewer);
}

export function reviewerOf(model: EditorModel): Reviewer {
  return reviewers.get(model) ?? DEFAULT_REVIEWER;
}

function who(model: EditorModel): RevisionAuthor {
  return { author: reviewerOf(model).author };
}

const OFF_VALUES: ReadonlySet<string> = new Set(["0", "false", "off"]);

/** Whether the document has Track Changes switched on (`w:trackRevisions`). */
export function trackRevisionsSetting(model: EditorModel): boolean {
  const setting = getDocumentSetting(model.doc, "trackRevisions");
  return setting.present && !OFF_VALUES.has(setting.val ?? "");
}

/**
 * Whether edits are recorded now: Track Changes is on, or the document is
 * protected so that only tracked changes are allowed.
 */
export function isTrackingRevisions(model: EditorModel): boolean {
  if (trackRevisionsSetting(model)) return true;
  const protection = documentProtection(model.doc);
  return !!protection?.enforced && protection.edit === "trackedChanges";
}

function samePara(a: DocPosition, b: DocPosition): boolean {
  return (
    a.block === b.block &&
    a.cell?.row === b.cell?.row &&
    a.cell?.col === b.cell?.col &&
    (a.para ?? 0) === (b.para ?? 0)
  );
}

function base(pos: DocPosition): DocPosition {
  return { block: pos.block, ...(pos.cell ? { cell: pos.cell, para: pos.para ?? 0 } : {}) };
}

/**
 * Mark the selected range deleted and collapse the caret to its start. Every
 * paragraph mark inside the range is marked deleted too, so accepting joins
 * the paragraphs as deleting them would.
 */
export function trackedDeleteSelection(model: EditorModel): DocPosition | undefined {
  const sel = model.selection;
  if (!sel) return undefined;
  const { start, end, collapsed } = orderSelection(sel);
  if (collapsed) return sel.focus;
  const doc = model.doc;
  const author = who(model);
  const first = paragraphAt(doc, start);
  if (!first) throw new Error("The selection is not inside a paragraph.");
  const from = absoluteOffset(first, start);
  if (samePara(start, end)) {
    deleteTrackedText(doc, first, from, absoluteOffset(first, end), author);
  } else {
    const paras = rangeParagraphs(model, start, end);
    for (const [i, para] of paras.entries()) {
      const isFirst = i === 0;
      const isLast = i === paras.length - 1;
      const total = paraLength(para);
      const s = isFirst ? from : 0;
      const e = isLast ? Math.min(absoluteOffset(para, end), total) : total;
      deleteTrackedText(doc, para, s, e, author);
      if (!isLast) trackParagraphMark(doc, para, "del", author);
    }
  }
  const caret = positionAtOffset(first, from, base(start));
  model.setSelection(caretAt(caret));
  return caret;
}

function paraLength(para: WmlParagraph): number {
  let n = 0;
  for (const child of para.children) if (child.kind === "run") n += runTextLength(child);
  return n;
}

/** The paragraphs from `start`'s to `end`'s, in order (within one cell, or top-level). */
function rangeParagraphs(model: EditorModel, start: DocPosition, end: DocPosition): WmlParagraph[] {
  if (start.cell || end.cell) {
    const cell = cellAt(model.doc, start);
    const sameCell =
      cell !== undefined &&
      start.block === end.block &&
      start.cell?.row === end.cell?.row &&
      start.cell?.col === end.cell?.col;
    if (!sameCell) {
      throw new Error("Deleting a selection that crosses table cells is not supported yet.");
    }
    return cell.paragraphs.slice(start.para ?? 0, (end.para ?? 0) + 1);
  }
  return paragraphsInRange(model.doc, { start, end, collapsed: false });
}

/**
 * Type `text` at the caret as a tracked insertion, replacing (marking deleted)
 * a range selection first. Line breaks start new paragraphs whose marks are
 * recorded as inserted.
 */
export function trackedInsertText(model: EditorModel, text: string): void {
  const start = trackedDeleteSelection(model);
  if (!start) throw new Error("Insert text needs a caret position.");
  const [firstLine = "", ...rest] = text.split(/\r\n|\r|\n/);
  insertAtCaret(model, firstLine);
  for (const line of rest) {
    trackedSplit(model);
    insertAtCaret(model, line);
  }
}

function insertAtCaret(model: EditorModel, text: string): void {
  const pos = model.selection?.focus;
  const para = pos && paragraphAt(model.doc, pos);
  if (!pos || !para) throw new Error("The caret is not inside a paragraph.");
  const end = insertTrackedText(model.doc, para, absoluteOffset(para, pos), text, who(model));
  model.setSelection(caretAt(positionAtOffset(para, end, base(pos))));
}

/** The paragraph right before the caret's (same cell, or the previous top-level block). */
function previousParagraph(
  model: EditorModel,
  pos: DocPosition,
):
  | {
      para: WmlParagraph;
      pos: DocPosition;
    }
  | undefined {
  if (pos.cell) {
    const index = (pos.para ?? 0) - 1;
    const para = cellAt(model.doc, pos)?.paragraphs[index];
    return para && { para, pos: { block: pos.block, cell: pos.cell, para: index } };
  }
  const block = model.doc.document.body.blocks[pos.block - 1];
  return block?.kind === "paragraph" ? { para: block, pos: { block: pos.block - 1 } } : undefined;
}

/**
 * Enter with Track Changes on, after the structural split ran: the mark that
 * now ends the paragraph before the caret is new, so it is recorded as
 * inserted.
 */
export function recordSplit(model: EditorModel): void {
  const pos = model.selection?.focus;
  const prev = pos && previousParagraph(model, pos);
  if (prev) trackParagraphMark(model.doc, prev.para, "ins", who(model));
}

/** Split at the caret as Enter does with tracking on (used for multi-line typing). */
function trackedSplit(model: EditorModel): void {
  const pos = model.selection?.focus;
  if (!pos) return;
  const doc = model.doc;
  const para = paragraphAt(doc, pos);
  if (!para) throw new Error("The caret is not inside a paragraph.");
  const offset = absoluteOffset(para, pos);
  // The text before the caret moves to a new paragraph in front; its mark is
  // the new one, the original mark stays with the text after the caret.
  const head: WmlParagraph = {
    kind: "paragraph",
    ...(para.pPr ? { pPr: structuredClone(para.pPr) } : {}),
    children: [],
    extras: [],
  };
  isolateParagraphRunRange(para, 0, offset);
  const split = splitIndex(para, offset);
  head.children = para.children.slice(0, split);
  para.children = para.children.slice(split);
  if (!head.children.some((c) => c.kind === "run")) {
    head.children.push({ kind: "run", pieces: [], extras: [] });
  }
  if (!para.children.some((c) => c.kind === "run")) {
    para.children.push({ kind: "run", pieces: [], extras: [] });
  }
  trackParagraphMark(doc, head, "ins", who(model));
  let next: DocPosition;
  if (pos.cell) {
    const cell = cellAt(doc, pos);
    if (!cell) throw new Error("The caret is not inside a table cell.");
    const index = pos.para ?? 0;
    cell.paragraphs.splice(index, 0, head);
    next = { block: pos.block, cell: pos.cell, para: index + 1 };
  } else {
    doc.document.body.blocks.splice(pos.block, 0, head);
    next = { block: pos.block + 1 };
  }
  model.setSelection(caretAt({ ...next, inline: 0, offset: 0 }));
}

function splitIndex(para: WmlParagraph, offset: number): number {
  let cursor = 0;
  for (const [i, child] of para.children.entries()) {
    if (child.kind !== "run") continue;
    const len = runTextLength(child);
    if (cursor + len > offset) return i;
    cursor += len;
  }
  return para.children.length;
}

/**
 * Backspace at a paragraph start (or Delete at its end) with tracking on:
 * the mark between the paragraphs is recorded as deleted and the caret moves
 * to the end of the previous paragraph. Returns false — after removing the
 * insertion record — when that mark is the author's own pending insertion,
 * so the caller joins the paragraphs for real.
 */
export function trackedMergeBack(model: EditorModel): boolean {
  const pos = model.selection?.focus;
  const prev = pos && previousParagraph(model, pos);
  if (!pos || !prev) return true;
  const mark = paragraphMarkRevision(prev.para);
  if (mark?.kind === "ins" && mark.author === reviewerOf(model).author) {
    removeMarkRecord(prev.para, "ins");
    return false;
  }
  trackParagraphMark(model.doc, prev.para, "del", who(model));
  model.setSelection(
    caretAt(positionAtOffset(prev.para, Number.POSITIVE_INFINITY, base(prev.pos))),
  );
  return true;
}

function removeMarkRecord(para: WmlParagraph, local: "ins" | "del"): void {
  const rPr = para.pPr?.children.find(
    (c): c is XmlElement => c.kind === "element" && c.name.uri === WML_NS && c.name.local === "rPr",
  );
  if (!rPr) return;
  const list = rPr.children as XmlNode[];
  const at = list.findIndex(
    (c) => c.kind === "element" && c.name.uri === WML_NS && c.name.local === local,
  );
  if (at >= 0) list.splice(at, 1);
}

/**
 * Backspace (`direction` -1) or Delete (+1) at a collapsed caret with
 * tracking on: the character before / after the caret is marked deleted
 * (already-deleted text is skipped, since it has no width), or at a
 * paragraph edge the paragraph mark is.
 */
export function trackedDeleteChar(model: EditorModel, direction: -1 | 1): void {
  const pos = model.selection?.focus;
  const para = pos && paragraphAt(model.doc, pos);
  if (!pos || !para) return;
  const offset = absoluteOffset(para, pos);
  if (direction < 0 && offset === 0) {
    trackedMergeBack(model);
    return;
  }
  if (direction > 0 && offset >= paraLength(para)) {
    const next: DocPosition | undefined = pos.cell
      ? { block: pos.block, cell: pos.cell, para: (pos.para ?? 0) + 1 }
      : { block: pos.block + 1 };
    if (!paragraphAt(model.doc, next)) return;
    trackParagraphMark(model.doc, para, "del", who(model));
    return;
  }
  const from = direction < 0 ? offset - 1 : offset;
  deleteTrackedText(model.doc, para, from, from + 1, who(model));
  model.setSelection(caretAt(positionAtOffset(para, from, base(pos))));
}

/**
 * Apply a run formatting edit and record it as a tracked change when
 * tracking is on (`<w:rPrChange>` with the run's properties from before).
 */
export function withRunFormatTracking(model: EditorModel, run: WmlRun, apply: () => void): void {
  if (!isTrackingRevisions(model)) {
    apply();
    return;
  }
  const before = run.rPr && structuredClone(run.rPr);
  apply();
  trackRunFormatChange(model.doc, run, before, who(model));
}

function propsKey(el: XmlElement | undefined): string {
  return JSON.stringify(el ?? null);
}

/**
 * Snapshot the selected paragraphs' properties before a formatting command,
 * and return the step that records `<w:pPrChange>` for each one the command
 * changed. `undefined` when tracking is off.
 */
export function beginParagraphFormatTracking(model: EditorModel): (() => void) | undefined {
  const sel = model.selection;
  if (!sel || !isTrackingRevisions(model)) return undefined;
  const paras = paragraphsInRange(model.doc, orderSelection(sel));
  const before = new Map(paras.map((p) => [p, p.pPr && structuredClone(p.pPr)] as const));
  return () => {
    for (const [para, old] of before) {
      if (propsKey(para.pPr) !== propsKey(old))
        trackParagraphFormatChange(model.doc, para, old, who(model));
    }
  };
}
