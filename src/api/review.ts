/**
 * Review: tracked changes (reading, recording, accepting and rejecting),
 * comments, proofing language, Word-compatible word counts and the
 * accessibility checker — the library side of Word's Review tab.
 */

import {
  serializeXml,
  type XmlAttr,
  type XmlElement,
  type XmlNode,
} from "../internal/xml/index.js";
import {
  buildComment,
  listRevisions,
  nextAnnotationId,
  parseParagraph,
  resolveRevisions,
  revisionAttrs,
  revisionDate,
  revisionElement,
  type RevisionInfo,
  setParagraphMarkRevision,
  toDeletedPiece,
  WML_NS,
  wmlAttrValue,
  type WmlBlock,
  type WmlParagraph,
  type WmlRun,
} from "../internal/wordprocessingml/index.js";
import {
  commentsPart,
  coreProperties,
  type Docx,
  endnotesPart,
  footnotesPart,
  isolateParagraphRunRange,
  outline,
  runTextLength,
} from "./docx.js";

export type { RevisionInfo, RevisionKind } from "../internal/wordprocessingml/index.js";

// --- shared helpers -------------------------------------------------------------

/** Who makes a tracked change or comment, and when (ISO 8601; now when omitted). */
export interface RevisionAuthor {
  readonly author: string;
  readonly date?: string;
}

function dateOf(who: RevisionAuthor): string {
  return revisionDate(who.date);
}

function isWml(node: XmlNode | undefined, local: string): node is XmlElement {
  return node?.kind === "element" && node.name.uri === WML_NS && node.name.local === local;
}

/** Whether an element (already known to be one) is `w:<local>`. */
function named(el: XmlElement, local: string): boolean {
  return el.name.uri === WML_NS && el.name.local === local;
}

function childEl(el: XmlElement | undefined, local: string): XmlElement | undefined {
  return el?.children.find((c): c is XmlElement => isWml(c, local));
}

function mutableChildren(el: XmlElement): XmlNode[] {
  // XmlElement's child list is readonly to discourage sharing; these helpers
  // edit one properties element in place, as the builders do.
  return el.children as XmlNode[];
}

function sameXml(a: XmlElement | undefined, b: XmlElement | undefined): boolean {
  const text = (el: XmlElement | undefined): string =>
    el ? serializeXml({ prologue: [], root: el, epilogue: [] }) : "";
  return text(a) === text(b);
}

function propsElement(local: string, children: XmlNode[] = []): XmlElement {
  return revisionElement(local, [], children);
}

function allParagraphs(blocks: readonly WmlBlock[]): Array<{ para: WmlParagraph; block: number }> {
  const out: Array<{ para: WmlParagraph; block: number }> = [];
  for (const [block, node] of blocks.entries()) {
    if (node.kind === "paragraph") out.push({ para: node, block });
    else if (node.kind === "table") {
      for (const row of node.rows) {
        for (const cell of row.cells) for (const para of cell.paragraphs) out.push({ para, block });
      }
    }
  }
  return out;
}

/** A fresh `w:id` for a revision or comment (unique across annotation types). */
function newId(doc: Docx): number {
  return nextAnnotationId(
    doc.document.body.blocks,
    doc.document.body.sectPr,
    commentsPart(doc)?.comments ?? [],
  );
}

// --- reading and resolving revisions --------------------------------------------

/**
 * Every tracked change in the body, in document order: insertions and
 * deletions (text and paragraph marks), formatting changes and table-row
 * insertions / deletions.
 */
export function revisions(doc: Docx): RevisionInfo[] {
  return listRevisions(doc.document);
}

/**
 * Accept the tracked changes with the given `w:id`s (Review ▸ Accept). Returns
 * how many were resolved.
 */
export function acceptRevisions(doc: Docx, ids: Iterable<string>): number {
  const wanted = new Set(ids);
  const n = resolveRevisions(doc.document, "accept", (id) => wanted.has(id));
  if (n > 0) doc.dirty = true;
  return n;
}

/**
 * Reject the tracked changes with the given `w:id`s (Review ▸ Reject): an
 * insertion is removed, a deletion restored, a formatting change undone.
 */
export function rejectRevisions(doc: Docx, ids: Iterable<string>): number {
  const wanted = new Set(ids);
  const n = resolveRevisions(doc.document, "reject", (id) => wanted.has(id));
  if (n > 0) doc.dirty = true;
  return n;
}

// --- recording revisions --------------------------------------------------------

function authorOf(run: WmlRun): string | undefined {
  return run.revision && wmlAttrValue(run.revision, "author");
}

function isSimpleText(run: WmlRun): boolean {
  return run.pieces.every((p) => p.kind === "text");
}

function simpleText(run: WmlRun): string {
  return run.pieces.map((p) => (p.kind === "text" ? p.value : "")).join("");
}

function setText(run: WmlRun, text: string): void {
  run.pieces = text
    ? [{ kind: "text", value: text, preserveSpace: /^\s|\s$|\s\s/.test(text) }]
    : [];
}

/** Whether `run` is a tracked insertion `who` made (typing into it stays one insertion). */
function ownInsertion(run: WmlRun, who: RevisionAuthor): boolean {
  return run.revision?.kind === "ins" && authorOf(run) === who.author;
}

/** The run's properties for new text typed next to it, without its revision record. */
function inheritedRPr(run: WmlRun | undefined): XmlElement | undefined {
  if (!run?.rPr) return undefined;
  const copy = structuredClone(run.rPr);
  const list = mutableChildren(copy);
  const change = list.findIndex((c) => isWml(c, "rPrChange"));
  if (change >= 0) list.splice(change, 1);
  return copy;
}

/**
 * Insert `text` at character `offset` of `para` as a tracked insertion by
 * `who` (Track Changes on, typing). Text typed next to (or inside) the same
 * author's insertion extends it, as in Word. Returns the offset after the
 * inserted text. Throws when `offset` falls inside a run that cannot be split
 * (one holding a tab, break or field), like plain insertion does.
 */
export function insertTrackedText(
  doc: Docx,
  para: WmlParagraph,
  offset: number,
  text: string,
  who: RevisionAuthor,
): number {
  if (text === "") return offset;
  // Inside (or at the end of) the author's own insertion: just add the text.
  let cursor = 0;
  for (const child of para.children) {
    if (child.kind !== "run") continue;
    const len = runTextLength(child);
    if (
      ownInsertion(child, who) &&
      isSimpleText(child) &&
      offset > cursor &&
      offset <= cursor + len
    ) {
      const current = simpleText(child);
      const at = offset - cursor;
      setText(child, current.slice(0, at) + text + current.slice(at));
      doc.dirty = true;
      return offset + text.length;
    }
    cursor += len;
  }
  isolateParagraphRunRange(para, 0, offset);
  // The insertion goes before the first run that extends past `offset`, so
  // after any zero-length deleted runs there: new text follows deleted text,
  // as Word places it.
  let index = para.children.length;
  let before: WmlRun | undefined;
  cursor = 0;
  for (const [i, child] of para.children.entries()) {
    if (child.kind !== "run") continue;
    const len = runTextLength(child);
    if (cursor + len > offset) {
      if (cursor !== offset) {
        throw new Error("Cannot insert text inside a run containing tabs, breaks, or fields.");
      }
      index = i;
      break;
    }
    cursor += len;
    if (child.revision?.kind !== "del") before = child;
  }
  const next = para.children[index];
  if (next?.kind === "run" && ownInsertion(next, who) && isSimpleText(next)) {
    setText(next, text + simpleText(next));
    doc.dirty = true;
    return offset + text.length;
  }
  const source = before ?? para.children.find((c): c is WmlRun => c.kind === "run");
  const rPr = inheritedRPr(source);
  const run: WmlRun = {
    kind: "run",
    ...(rPr ? { rPr } : {}),
    revision: {
      kind: "ins",
      attrs: revisionAttrs(newId(doc), who.author, dateOf(who)),
    },
    pieces: [],
    extras: [],
  };
  setText(run, text);
  para.children.splice(index, 0, run);
  doc.dirty = true;
  return offset + text.length;
}

/**
 * Mark characters `[start, end)` of `para` as a tracked deletion by `who`.
 * Text that is the author's own pending insertion is removed outright (Word
 * does not record deleting what you just typed); everything else stays in the
 * document as `<w:del>` / `<w:delText>`. Returns the number of characters
 * marked or removed.
 */
export function deleteTrackedText(
  doc: Docx,
  para: WmlParagraph,
  start: number,
  end: number,
  who: RevisionAuthor,
): number {
  const runs = new Set(isolateParagraphRunRange(para, start, end));
  if (runs.size === 0) return 0;
  let id: number | undefined;
  let count = 0;
  const next: typeof para.children = [];
  for (const child of para.children) {
    if (child.kind !== "run" || !runs.has(child)) {
      next.push(child);
      continue;
    }
    count += runTextLength(child);
    // Deleting a pending insertion — anyone's — takes it out: the text was
    // never part of the original, so there is nothing to restore on reject.
    if (child.revision?.kind === "ins") continue;
    id ??= newId(doc);
    next.push({
      ...child,
      pieces: child.pieces.map(toDeletedPiece),
      revision: { kind: "del", attrs: revisionAttrs(id, who.author, dateOf(who)) },
    });
  }
  para.children = next;
  doc.dirty = true;
  return count;
}

/** The tracked insertion / deletion recorded on a paragraph's mark, if any. */
export function paragraphMarkRevision(
  para: WmlParagraph,
): { kind: "ins" | "del"; author?: string } | undefined {
  const mark = childEl(para.pPr, "rPr");
  for (const kind of ["del", "ins"] as const) {
    const marker = childEl(mark, kind);
    if (marker) {
      const author = wmlAttrValue(marker, "author");
      return author === undefined ? { kind } : { kind, author };
    }
  }
  return undefined;
}

function ensurePPr(para: WmlParagraph): XmlElement {
  para.pPr ??= propsElement("pPr");
  return para.pPr;
}

/**
 * Record a tracked insertion or deletion of the paragraph mark (pressing Enter
 * / joining paragraphs with Track Changes on): `<w:ins>` / `<w:del>` first in
 * the mark's `<w:rPr>` (CT_ParaRPr order). A deletion over the author's own
 * inserted mark should instead remove the mark — {@link paragraphMarkRevision}
 * tells the caller which case it is in.
 */
export function trackParagraphMark(
  doc: Docx,
  para: WmlParagraph,
  kind: "ins" | "del",
  who: RevisionAuthor,
): void {
  setParagraphMarkRevision(para, kind, revisionAttrs(newId(doc), who.author, dateOf(who)));
  doc.dirty = true;
}

/** The change element and the properties it recorded, from a pre-edit copy of a properties element. */
function originalProperties(
  before: XmlElement | undefined,
  changeLocal: string,
  ignored: ReadonlySet<string>,
): { change: XmlElement | undefined; children: XmlNode[] } {
  const change = childEl(before, changeLocal);
  if (change) {
    const recorded = change.children.find((c): c is XmlElement => c.kind === "element");
    return { change, children: recorded ? [...recorded.children] : [] };
  }
  return {
    change: undefined,
    children: (before?.children ?? []).filter(
      (c) => !(c.kind === "element" && ignored.has(c.name.local)),
    ),
  };
}

function recordPropertyChange(
  doc: Docx,
  current: XmlElement,
  before: XmlElement | undefined,
  changeLocal: "rPrChange" | "pPrChange",
  propsLocal: "rPr" | "pPr",
  ignored: ReadonlySet<string>,
  who: RevisionAuthor,
): void {
  const original = originalProperties(before, changeLocal, ignored);
  const list = mutableChildren(current);
  const at = list.findIndex((c) => isWml(c, changeLocal));
  if (at >= 0) list.splice(at, 1);
  const now = list.filter((c) => !(c.kind === "element" && ignored.has(c.name.local)));
  // Back to how it was: no change left to record.
  if (sameXml(propsElement(propsLocal, now), propsElement(propsLocal, original.children))) return;
  const attrs: readonly XmlAttr[] =
    original.change?.attrs ?? revisionAttrs(newId(doc), who.author, dateOf(who));
  // The change element is last in both CT_RPr and CT_PPr.
  list.push(revisionElement(changeLocal, attrs, [propsElement(propsLocal, original.children)]));
}

const RPR_CHANGE_IGNORED: ReadonlySet<string> = new Set([
  "rPrChange",
  "ins",
  "del",
  "moveFrom",
  "moveTo",
]);
const PPR_CHANGE_IGNORED: ReadonlySet<string> = new Set(["pPrChange", "rPr", "sectPr"]);

/**
 * Record a tracked formatting change on `run`, given a copy of its `<w:rPr>`
 * from before the edit (`structuredClone(run.rPr)`): writes `<w:rPrChange>`
 * holding the original properties, keeps an earlier record's originals when
 * the run was already changed, and drops the record when the run is back to
 * its original formatting. A pending insertion records nothing — its
 * formatting is part of the insertion.
 */
export function trackRunFormatChange(
  doc: Docx,
  run: WmlRun,
  before: XmlElement | undefined,
  who: RevisionAuthor,
): void {
  if (run.revision) return;
  run.rPr ??= propsElement("rPr");
  recordPropertyChange(doc, run.rPr, before, "rPrChange", "rPr", RPR_CHANGE_IGNORED, who);
  if (run.rPr.children.length === 0) delete run.rPr;
  doc.dirty = true;
}

/**
 * Record a tracked paragraph-formatting change, given a copy of the
 * paragraph's `<w:pPr>` from before the edit: `<w:pPrChange>` with the
 * original paragraph properties (the mark's rPr and sectPr are not part of
 * it, CT_PPrBase).
 */
export function trackParagraphFormatChange(
  doc: Docx,
  para: WmlParagraph,
  before: XmlElement | undefined,
  who: RevisionAuthor,
): void {
  const pPr = ensurePPr(para);
  recordPropertyChange(doc, pPr, before, "pPrChange", "pPr", PPR_CHANGE_IGNORED, who);
  doc.dirty = true;
}

// --- comments -------------------------------------------------------------------

export interface CommentInfo {
  readonly id: number;
  readonly author: string;
  readonly initials?: string;
  readonly date?: string;
  /** The comment's text, paragraphs joined with `\n`. */
  readonly text: string;
  /** Top-level block holding the comment's anchor (its range start or reference). */
  readonly block?: number;
}

function commentText(comment: XmlElement): string {
  return comment.children
    .filter((c): c is XmlElement => isWml(c, "p"))
    .map((p) => {
      let out = "";
      const walk = (el: XmlElement): void => {
        if (named(el, "t")) {
          for (const c of el.children) if (c.kind === "text") out += c.value;
          return;
        }
        for (const c of el.children) if (c.kind === "element") walk(c);
      };
      walk(p);
      return out;
    })
    .join("\n");
}

function anchorId(node: XmlNode): string | undefined {
  if (node.kind !== "element") return undefined;
  if (named(node, "commentRangeStart") || named(node, "commentReference")) {
    return wmlAttrValue(node, "id");
  }
  for (const child of node.children) {
    const id = anchorId(child);
    if (id !== undefined) return id;
  }
  return undefined;
}

/** Every comment in `comments.xml`, with the block its anchor is in. */
export function comments(doc: Docx): CommentInfo[] {
  const anchors = new Map<string, number>();
  for (const { para, block } of allParagraphs(doc.document.body.blocks)) {
    for (const child of para.children) {
      const id =
        child.kind === "raw"
          ? anchorId(child.node)
          : child.pieces
              .map((p) => (p.kind === "raw" ? anchorId(p.node) : undefined))
              .find((x) => x !== undefined);
      if (id !== undefined && !anchors.has(id)) anchors.set(id, block);
    }
  }
  return (commentsPart(doc)?.comments ?? []).map((comment) => {
    const id = wmlAttrValue(comment, "id") ?? "";
    const initials = wmlAttrValue(comment, "initials");
    const date = wmlAttrValue(comment, "date");
    const block = anchors.get(id);
    return {
      id: Number(id),
      author: wmlAttrValue(comment, "author") ?? "",
      ...(initials === undefined ? {} : { initials }),
      ...(date === undefined ? {} : { date }),
      text: commentText(comment),
      ...(block === undefined ? {} : { block }),
    };
  });
}

function findComment(doc: Docx, id: number): XmlElement | undefined {
  return commentsPart(doc)?.comments.find((c) => wmlAttrValue(c, "id") === String(id));
}

/** Replace a comment's text (Edit Comment). Returns false when there is no such comment. */
export function setCommentText(doc: Docx, id: number, text: string): boolean {
  const comment = findComment(doc, id);
  if (!comment) return false;
  const rebuilt = buildComment({ id, author: "", text });
  mutableChildren(comment).splice(0, comment.children.length, ...rebuilt.children);
  doc.commentsDirty = true;
  return true;
}

/** Whether an inline is one of comment `id`'s anchors (range start / end / reference run). */
function isAnchorOf(node: XmlNode, id: string): boolean {
  if (
    (isWml(node, "commentRangeStart") ||
      isWml(node, "commentRangeEnd") ||
      isWml(node, "commentReference")) &&
    wmlAttrValue(node, "id") === id
  ) {
    return true;
  }
  return isWml(node, "r") && node.children.some((c) => isAnchorOf(c, id));
}

/**
 * Delete one comment (Review ▸ Delete): its entry in `comments.xml` and its
 * range markers and reference mark in the body. Returns false when there is
 * no such comment.
 */
export function removeComment(doc: Docx, id: number): boolean {
  const part = commentsPart(doc);
  const comment = findComment(doc, id);
  if (!part || !comment) return false;
  part.comments.splice(part.comments.indexOf(comment), 1);
  doc.commentsDirty = true;
  const key = String(id);
  for (const { para } of allParagraphs(doc.document.body.blocks)) {
    para.children = para.children.filter((child) =>
      child.kind === "raw"
        ? !isAnchorOf(child.node, key)
        : !(
            child.pieces.length > 0 &&
            child.pieces.every((p) => p.kind === "raw" && isAnchorOf(p.node, key))
          ),
    );
  }
  doc.dirty = true;
  return true;
}

// --- proofing language ----------------------------------------------------------

/** `<w:lang>` (§17.3.2.20): languages for Latin, East Asian and complex-script text. */
export interface RunLanguage {
  readonly latin?: string;
  readonly eastAsia?: string;
  readonly bidi?: string;
}

// ST_Lang: an RFC 4646 / BCP 47 tag ("en-US", "ja-JP") or a two-byte hex LCID.
const LANGUAGE_TAG = /^(?:[A-Za-z]{2,8}(?:-[A-Za-z0-9]{1,8})*|[0-9A-Fa-f]{4})$/;
const LANG_ATTRS = { latin: "val", eastAsia: "eastAsia", bidi: "bidi" } as const;

/** The run's direct `<w:lang>` settings. */
export function getRunLanguage(run: WmlRun): RunLanguage {
  const lang = childEl(run.rPr, "lang");
  if (!lang) return {};
  const out: { latin?: string; eastAsia?: string; bidi?: string } = {};
  for (const [key, local] of Object.entries(LANG_ATTRS)) {
    const value = wmlAttrValue(lang, local);
    if (value !== undefined) out[key as keyof RunLanguage] = value;
  }
  return out;
}

/**
 * Set the run's proofing languages (Review ▸ Language ▸ Set Proofing
 * Language). Keys left out are cleared; an empty object removes `<w:lang>`.
 * Throws a `RangeError` for a value that is not a language tag (ST_Lang).
 */
export function setRunLanguage(run: WmlRun, language: RunLanguage): void {
  for (const value of Object.values(language)) {
    if (value !== undefined && !LANGUAGE_TAG.test(value)) {
      throw new RangeError(`Not a language tag (ST_Lang): ${JSON.stringify(value)}.`);
    }
  }
  const attrs: XmlAttr[] = [];
  for (const [key, local] of Object.entries(LANG_ATTRS)) {
    const value = language[key as keyof RunLanguage];
    if (value !== undefined) {
      attrs.push({ name: { uri: WML_NS, local, prefix: "w" }, value, isNamespaceDecl: false });
    }
  }
  run.rPr ??= propsElement("rPr");
  const list = mutableChildren(run.rPr);
  const at = list.findIndex((c) => isWml(c, "lang"));
  if (at >= 0) list.splice(at, 1);
  if (attrs.length > 0) {
    // Before the trailing change record, which must stay last.
    const change = list.findIndex((c) => isWml(c, "rPrChange"));
    list.splice(change < 0 ? list.length : change, 0, revisionElement("lang", attrs));
  }
}

// --- word count -----------------------------------------------------------------

export interface WordCount {
  readonly words: number;
  /** Characters excluding spaces. */
  readonly characters: number;
  readonly charactersWithSpaces: number;
  /** Paragraphs with any text in them (Word does not count empty ones). */
  readonly paragraphs: number;
}

export interface WordCountOptions {
  /** Word's "Include textboxes, footnotes and endnotes" (default false). */
  readonly includeTextboxesAndNotes?: boolean;
}

// East Asian ideographs and kana count one word per character in Word.
const EAST_ASIAN = /[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}]/u;
const SPACE = /\s/u;

/** Visible text of a raw subtree: `<w:t>`, not deleted text or field codes. */
function visibleXmlText(el: XmlElement): string {
  if (named(el, "t")) return el.children.map((c) => (c.kind === "text" ? c.value : "")).join("");
  if (named(el, "tab")) return "\t";
  // Text boxes are counted on their own (Word's "Include textboxes" option).
  if (["del", "delText", "instrText", "txbxContent"].some((local) => named(el, local))) return "";
  let out = "";
  for (const child of el.children) if (child.kind === "element") out += visibleXmlText(child);
  return out;
}

/** A paragraph's counted text: what Word shows, without deletions or field codes. */
function countedText(para: WmlParagraph): string {
  let out = "";
  for (const child of para.children) {
    if (child.kind === "raw") {
      out += visibleXmlText(child.node);
      continue;
    }
    if (child.revision?.kind === "del") continue;
    for (const piece of child.pieces) {
      if (piece.kind === "text") out += piece.value;
      else if (piece.kind === "tab" || piece.kind === "break") out += " ";
    }
  }
  return out;
}

function textboxParagraphs(node: XmlNode, out: WmlParagraph[]): void {
  if (node.kind !== "element") return;
  if (named(node, "txbxContent")) {
    for (const child of node.children) if (isWml(child, "p")) out.push(parseParagraph(child));
    return;
  }
  for (const child of node.children) textboxParagraphs(child, out);
}

function noteParagraphs(notes: readonly XmlElement[]): WmlParagraph[] {
  // The separator notes (w:type separator / continuationSeparator / …) are
  // layout, not content.
  return notes
    .filter((note) => wmlAttrValue(note, "type") === undefined)
    .flatMap((note) => note.children.filter((c) => isWml(c, "p")).map((p) => parseParagraph(p)));
}

interface Totals {
  words: number;
  chars: number;
  withSpaces: number;
  paragraphs: number;
}

function countInto(totals: Totals, text: string): void {
  let inWord = false;
  let any = false;
  for (const ch of text) {
    totals.withSpaces++;
    if (SPACE.test(ch)) {
      inWord = false;
      continue;
    }
    any = true;
    totals.chars++;
    if (EAST_ASIAN.test(ch)) {
      totals.words++;
      inWord = false;
    } else if (!inWord) {
      totals.words++;
      inWord = true;
    }
  }
  if (any) totals.paragraphs++;
}

/**
 * Word's Word Count (Review ▸ Word Count): words, characters with and without
 * spaces, and non-empty paragraphs, by Word's rules — a word is a run of
 * non-space characters, except that every East Asian ideograph or kana counts
 * as one; deleted text and field codes are not counted. Pages and lines depend
 * on layout and are not computed here.
 */
export function wordCount(doc: Docx, options: WordCountOptions = {}): WordCount {
  const paragraphs = allParagraphs(doc.document.body.blocks).map(({ para }) => para);
  if (options.includeTextboxesAndNotes) {
    const boxes: WmlParagraph[] = [];
    for (const para of paragraphs) {
      for (const child of para.children) {
        if (child.kind === "raw") textboxParagraphs(child.node, boxes);
        else
          for (const piece of child.pieces) {
            if (piece.kind === "drawing" || piece.kind === "pict" || piece.kind === "raw") {
              textboxParagraphs(piece.node, boxes);
            }
          }
      }
    }
    paragraphs.push(
      ...boxes,
      ...noteParagraphs(footnotesPart(doc)?.footnotes ?? []),
      ...noteParagraphs(endnotesPart(doc)?.footnotes ?? []),
    );
  }
  const totals: Totals = { words: 0, chars: 0, withSpaces: 0, paragraphs: 0 };
  for (const para of paragraphs) countInto(totals, countedText(para));
  return {
    words: totals.words,
    characters: totals.chars,
    charactersWithSpaces: totals.withSpaces,
    paragraphs: totals.paragraphs,
  };
}

// --- accessibility --------------------------------------------------------------

export type AccessibilityIssueKind =
  | "missingAltText"
  | "missingTableHeader"
  | "skippedHeadingLevel"
  | "lowContrast"
  | "repeatedBlankParagraphs"
  | "missingTitle";

/** Word's Accessibility Checker groups: errors, warnings and tips. */
export type AccessibilitySeverity = "error" | "warning" | "tip";

export interface AccessibilityIssue {
  readonly kind: AccessibilityIssueKind;
  readonly severity: AccessibilitySeverity;
  /** Top-level block the issue is in (absent for document-wide issues). */
  readonly block?: number;
  /** What was found: the picture's name, the heading text, the colors … */
  readonly detail: string;
}

const SEVERITY: Readonly<Record<AccessibilityIssueKind, AccessibilitySeverity>> = {
  missingAltText: "error",
  missingTableHeader: "error",
  skippedHeadingLevel: "warning",
  lowContrast: "warning",
  repeatedBlankParagraphs: "tip",
  missingTitle: "tip",
};

// WCAG 2.x contrast thresholds, which Word's checker applies: 4.5:1, or 3:1
// for large text (18 pt, or 14 pt bold).
const MIN_CONTRAST = 4.5;
const MIN_CONTRAST_LARGE = 3;
const LARGE_HALF_POINTS = 36;
const LARGE_BOLD_HALF_POINTS = 28;
const BLANK_RUN_LIMIT = 3;
const HEX6 = /^[0-9A-Fa-f]{6}$/;
const HIGHLIGHT_HEX: Readonly<Record<string, string>> = {
  yellow: "FFFF00",
  green: "00FF00",
  cyan: "00FFFF",
  magenta: "FF00FF",
  blue: "0000FF",
  red: "FF0000",
  darkBlue: "000080",
  darkCyan: "008080",
  darkGreen: "008000",
  darkMagenta: "800080",
  darkRed: "800000",
  darkYellow: "808000",
  darkGray: "808080",
  lightGray: "C0C0C0",
  black: "000000",
  white: "FFFFFF",
};

function luminance(hex: string): number {
  const channel = (i: number): number => {
    const c = Number.parseInt(hex.slice(i, i + 2), 16) / 255;
    return c <= 0.039_28 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4;
  };
  return 0.2126 * channel(0) + 0.7152 * channel(2) + 0.0722 * channel(4);
}

/** WCAG contrast ratio of two `RRGGBB` colors. */
export function contrastRatio(a: string, b: string): number {
  const [hi, lo] = [luminance(a), luminance(b)].toSorted((x, y) => y - x);
  return ((hi ?? 0) + 0.05) / ((lo ?? 0) + 0.05);
}

/** An attribute of a properties element's child (`<w:color w:val>` …). */
function propAttr(props: XmlElement | undefined, local: string, attr = "val"): string | undefined {
  const el = childEl(props, local);
  return el && wmlAttrValue(el, attr);
}

function shadingFill(props: XmlElement | undefined): string | undefined {
  const fill = propAttr(props, "shd", "fill");
  return fill && HEX6.test(fill) ? fill : undefined;
}

function docPrOf(node: XmlNode): XmlElement | undefined {
  if (node.kind !== "element") return undefined;
  if (node.name.local === "docPr") return node;
  for (const child of node.children) {
    const found = docPrOf(child);
    if (found) return found;
  }
  return undefined;
}

function headingLevels(doc: Docx): Map<WmlParagraph, number> {
  return new Map(outline(doc).map(({ paragraph, level }) => [paragraph, level]));
}

/**
 * Check a document for the problems Word's Accessibility Checker reports:
 * pictures without alternative text, tables without a header row, skipped
 * heading levels, text colours with too little contrast against their
 * background, runs of blank paragraphs used for spacing, and a missing
 * document title.
 */
export function checkAccessibility(doc: Docx): AccessibilityIssue[] {
  const issues: AccessibilityIssue[] = [];
  const add = (kind: AccessibilityIssueKind, detail: string, block?: number): void => {
    issues.push({
      kind,
      severity: SEVERITY[kind],
      detail,
      ...(block === undefined ? {} : { block }),
    });
  };
  const headings = headingLevels(doc);
  let lastLevel = 0;
  let blanks = 0;
  for (const [block, node] of doc.document.body.blocks.entries()) {
    if (node.kind === "table") {
      blanks = 0;
      const header = childEl(node.rows[0]?.trPr, "tblHeader");
      if (node.rows.length > 1 && (!header || wmlAttrValue(header, "val") === "0")) {
        add(
          "missingTableHeader",
          `${node.rows.length} × ${node.rows[0]?.cells.length ?? 0}`,
          block,
        );
      }
    }
    if (node.kind !== "paragraph") continue;
    const level = headings.get(node);
    if (level !== undefined) {
      if (level > lastLevel + 1) add("skippedHeadingLevel", `${lastLevel} → ${level}`, block);
      lastLevel = level;
    }
    const empty =
      countedText(node).trim() === "" &&
      !node.children.some(
        (c) => c.kind === "run" && c.pieces.some((p) => p.kind === "drawing" || p.kind === "pict"),
      );
    blanks = empty ? blanks + 1 : 0;
    if (blanks === BLANK_RUN_LIMIT) add("repeatedBlankParagraphs", String(BLANK_RUN_LIMIT), block);
    const paragraphFill = shadingFill(node.pPr);
    for (const child of node.children) {
      if (child.kind !== "run") continue;
      for (const piece of child.pieces) {
        if (piece.kind !== "drawing") continue;
        const docPr = docPrOf(piece.node);
        const descr = docPr && docPr.attrs.find((a) => a.name.local === "descr")?.value;
        if (!descr?.trim()) {
          add(
            "missingAltText",
            docPr?.attrs.find((a) => a.name.local === "name")?.value ?? "",
            block,
          );
        }
      }
      const color = propAttr(child.rPr, "color");
      if (!color || !HEX6.test(color)) continue;
      const highlight = propAttr(child.rPr, "highlight");
      const background =
        (highlight && HIGHLIGHT_HEX[highlight]) ??
        shadingFill(child.rPr) ??
        paragraphFill ??
        "FFFFFF";
      const size = Number(propAttr(child.rPr, "sz") ?? 0);
      const bold = childEl(child.rPr, "b") !== undefined;
      const large = size >= LARGE_HALF_POINTS || (bold && size >= LARGE_BOLD_HALF_POINTS);
      const ratio = contrastRatio(color, background);
      if (ratio < (large ? MIN_CONTRAST_LARGE : MIN_CONTRAST)) {
        add("lowContrast", `#${color} / #${background} (${ratio.toFixed(2)}:1)`, block);
      }
    }
  }
  if (!coreProperties(doc).title?.trim()) add("missingTitle", "");
  return issues;
}
