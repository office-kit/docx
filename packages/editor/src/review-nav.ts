/**
 * Where tracked changes and comments are in the document, as caret positions:
 * what Review ▸ Previous / Next and Accept / Reject "this change" work from.
 */

import {
  type Docx,
  runTextLength,
  type WmlParagraph,
  type WmlRun,
  type XmlElement,
} from "@office-kit/docx";
import type { EditorModel } from "./model.js";
import { type CellCoord, type DocPosition, orderSelection } from "./selection.js";
import { WML_NS } from "./wml-ns.js";

/** A revision or comment anchor and the caret position it starts at. */
export interface ReviewMark {
  readonly kind: "revision" | "comment";
  /** The revision's or comment's `w:id`. */
  readonly id: string;
  readonly position: DocPosition;
}

const PROPERTY_CHANGES: ReadonlySet<string> = new Set(["rPrChange", "pPrChange"]);

function wAttr(el: { readonly attrs: XmlElement["attrs"] }, local: string): string | undefined {
  return el.attrs.find((a) => a.name.uri === WML_NS && a.name.local === local)?.value;
}

function child(el: XmlElement | undefined, local: string): XmlElement | undefined {
  return el?.children.find(
    (c): c is XmlElement => c.kind === "element" && c.name.uri === WML_NS && c.name.local === local,
  );
}

/** Revision ids recorded on a run: its insertion / deletion and a formatting change. */
function runRevisionIds(run: WmlRun): string[] {
  const ids: string[] = [];
  const own = run.revision && wAttr(run.revision, "id");
  if (own !== undefined) ids.push(own);
  for (const c of run.rPr?.children ?? []) {
    if (c.kind === "element" && c.name.uri === WML_NS && PROPERTY_CHANGES.has(c.name.local)) {
      const id = wAttr(c, "id");
      if (id !== undefined) ids.push(id);
    }
  }
  return ids;
}

/** Revision ids on the paragraph itself: a property change and its mark. */
function paragraphRevisionIds(para: WmlParagraph): string[] {
  const ids: string[] = [];
  const change = child(para.pPr, "pPrChange");
  const mark = child(para.pPr, "rPr");
  for (const el of [change, child(mark, "ins"), child(mark, "del"), child(mark, "rPrChange")]) {
    const id = el && wAttr(el, "id");
    if (id !== undefined) ids.push(id);
  }
  return ids;
}

interface Located {
  readonly para: WmlParagraph;
  readonly base: DocPosition;
}

function locatedParagraphs(doc: Docx): Located[] {
  const out: Located[] = [];
  for (const [block, node] of doc.document.body.blocks.entries()) {
    if (node.kind === "paragraph") out.push({ para: node, base: { block } });
    else if (node.kind === "table") {
      for (const [row, r] of node.rows.entries()) {
        for (const [col, c] of r.cells.entries()) {
          const cell: CellCoord = { row, col };
          for (const [para, p] of c.paragraphs.entries()) {
            out.push({ para: p, base: { block, cell, para } });
          }
        }
      }
    }
  }
  return out;
}

/** Every revision and comment start, in document order (a split revision once). */
export function reviewMarks(doc: Docx): ReviewMark[] {
  const out: ReviewMark[] = [];
  const seen = new Set<string>();
  const add = (kind: ReviewMark["kind"], id: string, position: DocPosition): void => {
    const key = `${kind}:${id}`;
    if (seen.has(key)) return;
    seen.add(key);
    out.push({ kind, id, position });
  };
  for (const { para, base } of locatedParagraphs(doc)) {
    let inline = 0;
    for (const c of para.children) {
      if (c.kind === "run") {
        for (const id of runRevisionIds(c)) add("revision", id, { ...base, inline, offset: 0 });
        inline++;
        continue;
      }
      const node = c.node;
      if (node.name.uri === WML_NS && node.name.local === "commentRangeStart") {
        const id = wAttr(node, "id");
        if (id !== undefined) add("comment", id, { ...base, inline, offset: 0 });
      }
    }
    const runs = para.children.filter((x): x is WmlRun => x.kind === "run");
    const last = runs.at(-1);
    const end = {
      ...base,
      inline: Math.max(runs.length - 1, 0),
      offset: last ? runTextLength(last) : 0,
    };
    for (const id of paragraphRevisionIds(para)) add("revision", id, end);
  }
  return out;
}

function comparePositions(a: DocPosition, b: DocPosition): number {
  const order = orderSelection({ anchor: a, focus: b });
  if (order.collapsed) return 0;
  return order.start === a ? -1 : 1;
}

/**
 * The next (`direction` 1) or previous (-1) mark of `kind` after / before the
 * caret, wrapping around the document as Word does. `undefined` when there is
 * none.
 */
export function adjacentReviewMark(
  model: EditorModel,
  kind: ReviewMark["kind"],
  direction: 1 | -1,
): ReviewMark | undefined {
  const marks = reviewMarks(model.doc).filter((m) => m.kind === kind);
  const caret = model.selection ? orderSelection(model.selection) : undefined;
  if (!caret) return direction > 0 ? marks[0] : marks.at(-1);
  if (direction > 0) {
    return marks.find((m) => comparePositions(m.position, caret.end) > 0) ?? marks[0];
  }
  return marks.findLast((m) => comparePositions(m.position, caret.start) < 0) ?? marks.at(-1);
}

/**
 * The revisions "this change" means: those the selection covers, or with a
 * caret the change the caret is in (or right after), falling back to the
 * caret paragraph's own property / mark changes.
 */
export function revisionIdsAtSelection(model: EditorModel): string[] {
  const sel = model.selection;
  if (!sel) return [];
  const { start, end, collapsed } = orderSelection(sel);
  const ids = new Set<string>();
  for (const { para, base } of locatedParagraphs(model.doc)) {
    const from = comparePositions({ ...base, inline: Number.MAX_SAFE_INTEGER }, start);
    const to = comparePositions({ ...base, inline: -1 }, end);
    if (from < 0 || to > 0) continue;
    const runs = para.children.filter((x): x is WmlRun => x.kind === "run");
    for (const [inline, run] of runs.entries()) {
      const runStart = { ...base, inline, offset: 0 };
      const runEnd = { ...base, inline, offset: runTextLength(run) };
      const touches = collapsed
        ? comparePositions(runStart, start) <= 0 && comparePositions(runEnd, start) >= 0
        : comparePositions(runEnd, start) > 0 && comparePositions(runStart, end) < 0;
      // A zero-width deleted run sits between characters: take it when the
      // caret is right at it.
      if (touches) for (const id of runRevisionIds(run)) ids.add(id);
    }
    if (!collapsed || ids.size === 0) for (const id of paragraphRevisionIds(para)) ids.add(id);
  }
  return [...ids];
}
