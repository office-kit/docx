import type { XmlAttr, XmlElement, XmlNode } from "../xml/index.js";
import { WML_NS } from "./namespaces.js";
import type {
  WmlDocument,
  WmlInline,
  WmlParagraph,
  WmlRawInline,
  WmlRun,
  WmlRunPiece,
  WmlRunRevision,
  WmlTable,
  WmlTableRow,
} from "./types.js";

/**
 * Tracked changes (ECMA-376 Part 1 §17.13.5): reading them, and accepting or
 * rejecting all of them or a chosen few by `w:id`.
 *
 * Revisions covered:
 * - run content: `<w:ins>` / `<w:del>` (lifted onto runs as `WmlRun.revision`,
 *   or still raw when the wrapper holds more than runs);
 * - the paragraph mark: `<w:ins>` / `<w:del>` in `<w:pPr><w:rPr>`;
 * - property changes: `<w:rPrChange>`, `<w:pPrChange>`, `<w:sectPrChange>`,
 *   `<w:tblPrChange>`, `<w:trPrChange>`, `<w:tcPrChange>`;
 * - table rows: `<w:ins>` / `<w:del>` in `<w:trPr>`.
 */

export type RevisionMode = "accept" | "reject";

export type RevisionKind =
  | "insert"
  | "delete"
  | "paragraphInsert"
  | "paragraphDelete"
  | "format"
  | "rowInsert"
  | "rowDelete";

/** One tracked change, as Word's Reviewing pane lists it. */
export interface RevisionInfo {
  /** The revision's `w:id`. Runs split from one insertion share it. */
  readonly id: string;
  readonly kind: RevisionKind;
  readonly author?: string;
  readonly date?: string;
  /** The inserted / deleted text (empty for formatting and paragraph marks). */
  readonly text: string;
  /** Index of the top-level body block that holds (the start of) the change. */
  readonly block: number;
}

/** Revision-marker children of a paragraph mark's `<w:rPr>` (CT_ParaRPr), first in order. */
const PARA_MARK_MARKERS: ReadonlySet<string> = new Set(["ins", "del", "moveFrom", "moveTo"]);
/** Children of `<w:pPr>` that a `<w:pPrChange>` does not describe (CT_PPrBase excludes them). */
const PPR_TRAILERS: ReadonlySet<string> = new Set(["rPr", "sectPr"]);
const PROPERTY_CHANGES = [
  "rPrChange",
  "pPrChange",
  "sectPrChange",
  "tblPrChange",
  "trPrChange",
  "tcPrChange",
] as const;
const PROPERTY_CHANGE_SET: ReadonlySet<string> = new Set(PROPERTY_CHANGES);

function isWml(node: XmlNode, local: string): node is XmlElement {
  return node.kind === "element" && node.name.uri === WML_NS && node.name.local === local;
}

/** Whether an element (already known to be one) is `w:<local>`. */
function named(el: XmlElement, local: string): boolean {
  return el.name.uri === WML_NS && el.name.local === local;
}

/** A `w:`-namespaced attribute value. */
export function wmlAttrValue(el: { readonly attrs: readonly XmlAttr[] }, local: string) {
  return el.attrs.find((a) => a.name.uri === WML_NS && a.name.local === local)?.value;
}

function childEl(el: XmlElement | undefined, local: string): XmlElement | undefined {
  return el?.children.find((c): c is XmlElement => isWml(c, local));
}

function mutableChildren(el: XmlElement): XmlNode[] {
  // The AST's arrays are readonly to stop accidental sharing; property edits
  // here rewrite one element's own child list in place, as builders.ts does.
  return el.children as XmlNode[];
}

/** `w:id`, `w:author`, `w:date` attributes for a new revision marker. */
export function revisionAttrs(id: number, author: string, date: string | undefined): XmlAttr[] {
  const attr = (local: string, value: string): XmlAttr => ({
    name: { uri: WML_NS, local, prefix: "w" },
    value,
    isNamespaceDecl: false,
  });
  return [
    attr("id", String(id)),
    attr("author", author),
    ...(date === undefined ? [] : [attr("date", date)]),
  ];
}

/** An empty `w:` element with the given attributes. */
export function revisionElement(
  local: string,
  attrs: readonly XmlAttr[],
  children: XmlNode[] = [],
): XmlElement {
  return {
    kind: "element",
    name: { uri: WML_NS, local, prefix: "w" },
    attrs,
    children,
    xmlSpace: "default",
    selfClosing: children.length === 0,
  };
}

// --- reading -------------------------------------------------------------------

function runText(run: WmlRun): string {
  let out = "";
  for (const piece of run.pieces) {
    if (piece.kind === "text" || piece.kind === "delText") out += piece.value;
    else if (piece.kind === "tab") out += "\t";
  }
  return out;
}

function rawText(el: XmlElement): string {
  if (named(el, "t") || named(el, "delText")) {
    return el.children.map((c) => (c.kind === "text" ? c.value : "")).join("");
  }
  let out = "";
  for (const c of el.children) if (c.kind === "element") out += rawText(c);
  return out;
}

type RevisionSink = (
  el: { readonly attrs: readonly XmlAttr[] },
  kind: RevisionKind,
  text: string,
) => void;

function visitProperties(props: XmlElement | undefined, sink: RevisionSink): void {
  if (!props) return;
  for (const child of props.children) {
    if (child.kind !== "element" || child.name.uri !== WML_NS) continue;
    if (PROPERTY_CHANGE_SET.has(child.name.local)) sink(child, "format", "");
    // A paragraph mark's rPr and a paragraph's sectPr carry their own changes.
    else if (child.name.local === "rPr" || child.name.local === "sectPr")
      visitProperties(child, sink);
  }
}

function visitParagraph(p: WmlParagraph, sink: RevisionSink): void {
  for (const child of p.children) {
    if (child.kind === "run") {
      if (child.revision) {
        sink(child.revision, child.revision.kind === "ins" ? "insert" : "delete", runText(child));
      }
      visitProperties(child.rPr, sink);
    } else if (isWml(child.node, "ins") || isWml(child.node, "del")) {
      sink(child.node, child.node.name.local === "ins" ? "insert" : "delete", rawText(child.node));
    }
  }
  visitProperties(p.pPr, sink);
  const mark = childEl(p.pPr, "rPr");
  const ins = childEl(mark, "ins");
  const del = childEl(mark, "del");
  if (ins) sink(ins, "paragraphInsert", "");
  if (del) sink(del, "paragraphDelete", "");
}

function visitTable(table: WmlTable, sink: RevisionSink): void {
  visitProperties(table.tblPr, sink);
  for (const row of table.rows) {
    visitProperties(row.trPr, sink);
    const ins = childEl(row.trPr, "ins");
    const del = childEl(row.trPr, "del");
    if (ins) sink(ins, "rowInsert", "");
    if (del) sink(del, "rowDelete", "");
    for (const cell of row.cells) {
      visitProperties(cell.tcPr, sink);
      for (const p of cell.paragraphs) visitParagraph(p, sink);
    }
  }
}

/**
 * Every tracked change in the body, in document order. Runs that share a
 * `w:id` and kind (an insertion split by later formatting) are one entry with
 * their text joined.
 */
export function listRevisions(doc: WmlDocument): RevisionInfo[] {
  const out: RevisionInfo[] = [];
  const byKey = new Map<string, number>();
  const sinkFor =
    (block: number): RevisionSink =>
    (el, kind, text) => {
      const id = wmlAttrValue(el, "id") ?? "";
      const key = `${kind}:${id}`;
      const seen = byKey.get(key);
      if (seen !== undefined) {
        const existing = out[seen];
        if (existing) out[seen] = { ...existing, text: existing.text + text };
        return;
      }
      byKey.set(key, out.length);
      const author = wmlAttrValue(el, "author");
      const date = wmlAttrValue(el, "date");
      out.push({
        id,
        kind,
        ...(author === undefined ? {} : { author }),
        ...(date === undefined ? {} : { date }),
        text,
        block,
      });
    };
  for (const [block, node] of doc.body.blocks.entries()) {
    if (node.kind === "paragraph") visitParagraph(node, sinkFor(block));
    else if (node.kind === "table") visitTable(node, sinkFor(block));
  }
  visitProperties(doc.body.sectPr, sinkFor(doc.body.blocks.length));
  return out;
}

// --- resolving -----------------------------------------------------------------

interface Ctx {
  readonly mode: RevisionMode;
  /** Which revisions to resolve; all when undefined. */
  readonly select: ((id: string) => boolean) | undefined;
  /** `kind:id` of every revision resolved, so split runs count once. */
  readonly resolved: Set<string>;
  /** Stray `<w:delText>` outside any revision (only resolved with "all"). */
  stray: number;
}

function selected(ctx: Ctx, el: { readonly attrs: readonly XmlAttr[] }, kind: string): boolean {
  const id = wmlAttrValue(el, "id") ?? "";
  if (ctx.select && !ctx.select(id)) return false;
  ctx.resolved.add(`${kind}:${id}`);
  return true;
}

/**
 * Accept or reject the property changes inside `props` (and inside the
 * paragraph mark's `rPr` / the paragraph's `sectPr` when `props` is a pPr).
 * Rejecting puts the recorded old properties back, keeping the children a
 * change does not describe (revision markers, the mark's rPr, sectPr).
 */
function resolveProperties(props: XmlElement | undefined, ctx: Ctx): void {
  if (!props) return;
  const list = mutableChildren(props);
  for (const child of list) {
    if (child.kind === "element" && (isWml(child, "rPr") || isWml(child, "sectPr"))) {
      resolveProperties(child, ctx);
    }
  }
  const change = list.find(
    (c): c is XmlElement =>
      c.kind === "element" && c.name.uri === WML_NS && PROPERTY_CHANGE_SET.has(c.name.local),
  );
  if (!change || !selected(ctx, change, "format")) return;
  if (ctx.mode === "accept") {
    list.splice(list.indexOf(change), 1);
    return;
  }
  const old = change.children.find((c): c is XmlElement => c.kind === "element");
  const front = list.filter((c) => c.kind === "element" && PARA_MARK_MARKERS.has(c.name.local));
  const back = list.filter((c) => c.kind === "element" && PPR_TRAILERS.has(c.name.local));
  list.splice(0, list.length, ...front, ...(old ? old.children : []), ...back);
}

function toLiveText(piece: WmlRunPiece): WmlRunPiece {
  if (piece.kind === "delText") return { ...piece, kind: "text" };
  if (piece.kind === "delInstrText") return { ...piece, kind: "instrText" };
  return piece;
}

function resolveRun(run: WmlRun, ctx: Ctx): WmlRun | undefined {
  resolveProperties(run.rPr, ctx);
  const revision: WmlRunRevision | undefined = run.revision;
  if (revision) {
    if (!selected(ctx, revision, revision.kind)) return run;
    const keep = (revision.kind === "ins") === (ctx.mode === "accept");
    if (!keep) return undefined;
    const live: WmlRun = { ...run, pieces: run.pieces.map(toLiveText) };
    delete live.revision;
    return live;
  }
  if (ctx.select) return run;
  // A deleted piece outside any `<w:del>` (hand-unwrapped): deleted text.
  const pieces: WmlRunPiece[] = [];
  for (const piece of run.pieces) {
    if (piece.kind === "delText" || piece.kind === "delInstrText") {
      ctx.stray++;
      if (ctx.mode === "reject") pieces.push(toLiveText(piece));
    } else pieces.push(piece);
  }
  run.pieces = pieces;
  return run;
}

/** Runs of a raw `<w:ins>` / `<w:del>` that held more than runs. */
function promoteRawRevision(raw: WmlRawInline, ctx: Ctx): WmlInline[] {
  const out: WmlInline[] = [];
  for (const child of raw.node.children) {
    if (child.kind !== "element") continue;
    if (isWml(child, "r")) out.push(resolveRawRun(child, ctx));
    else out.push({ kind: "raw", node: child });
  }
  return out;
}

function resolveRawRun(r: XmlElement, ctx: Ctx): WmlInline {
  if (ctx.mode === "accept") return { kind: "raw", node: r };
  // Rejecting a deletion: its `<w:delText>` become `<w:t>` again.
  const rename = (el: XmlElement): XmlElement => ({
    ...el,
    name: named(el, "delText")
      ? { ...el.name, local: "t" }
      : named(el, "delInstrText")
        ? { ...el.name, local: "instrText" }
        : el.name,
    children: el.children.map((c) => (c.kind === "element" ? rename(c) : c)),
  });
  return { kind: "raw", node: rename(r) };
}

type MarkOutcome = "keep" | "merge";

function resolveParagraph(p: WmlParagraph, ctx: Ctx): MarkOutcome {
  const next: WmlInline[] = [];
  for (const child of p.children) {
    if (child.kind === "run") {
      const run = resolveRun(child, ctx);
      if (run) next.push(run);
      continue;
    }
    const isIns = isWml(child.node, "ins");
    if ((isIns || isWml(child.node, "del")) && selected(ctx, child.node, isIns ? "ins" : "del")) {
      if (isIns === (ctx.mode === "accept")) next.push(...promoteRawRevision(child, ctx));
      continue;
    }
    next.push(child);
  }
  p.children = next;
  resolveProperties(p.pPr, ctx);
  const mark = childEl(p.pPr, "rPr");
  if (!mark) return "keep";
  let outcome: MarkOutcome = "keep";
  const list = mutableChildren(mark);
  for (const local of ["ins", "del"] as const) {
    const marker = childEl(mark, local);
    if (!marker || !selected(ctx, marker, `mark-${local}`)) continue;
    list.splice(list.indexOf(marker), 1);
    // Rejecting an inserted mark or accepting a deleted one removes the
    // paragraph break: the paragraph runs on into the next one.
    if ((local === "ins") === (ctx.mode === "reject")) outcome = "merge";
  }
  return outcome;
}

/**
 * Join paragraph `i` with the paragraph after it, as Word does when a
 * paragraph mark goes away: the text runs on, and the joined paragraph keeps
 * the *following* paragraph's properties (they live on the mark that remains).
 */
function joinWithNext(paragraphs: WmlParagraph[], i: number): boolean {
  const here = paragraphs[i];
  const after = paragraphs[i + 1];
  if (!here || !after) return false;
  after.children = [...here.children, ...after.children];
  paragraphs.splice(i, 1);
  return true;
}

function resolveParagraphList(paragraphs: WmlParagraph[], ctx: Ctx): void {
  const merges: number[] = [];
  for (const [i, p] of paragraphs.entries()) {
    if (resolveParagraph(p, ctx) === "merge") merges.push(i);
  }
  for (const i of merges.toReversed()) joinWithNext(paragraphs, i);
}

/** Whether a row survives: rejecting an inserted row or accepting a deleted one removes it. */
function resolveRow(row: WmlTableRow, ctx: Ctx): boolean {
  resolveProperties(row.trPr, ctx);
  let keep = true;
  for (const local of ["ins", "del"] as const) {
    const marker = childEl(row.trPr, local);
    if (!marker || !row.trPr || !selected(ctx, marker, `row-${local}`)) continue;
    const list = mutableChildren(row.trPr);
    list.splice(list.indexOf(marker), 1);
    if ((local === "ins") === (ctx.mode === "reject")) keep = false;
  }
  for (const cell of row.cells) {
    resolveProperties(cell.tcPr, ctx);
    resolveParagraphList(cell.paragraphs, ctx);
  }
  return keep;
}

function resolveBody(doc: WmlDocument, ctx: Ctx): void {
  const blocks = doc.body.blocks;
  const merges: number[] = [];
  for (const [i, block] of blocks.entries()) {
    if (block.kind === "paragraph") {
      if (resolveParagraph(block, ctx) === "merge") merges.push(i);
    } else if (block.kind === "table") {
      resolveProperties(block.tblPr, ctx);
      block.rows = block.rows.filter((row) => resolveRow(row, ctx));
    }
  }
  for (const i of merges.toReversed()) {
    const here = blocks[i];
    const after = blocks[i + 1];
    // A mark before a table (or at the end of the body) cannot run on into
    // anything; the marker alone is resolved.
    if (here?.kind !== "paragraph" || after?.kind !== "paragraph") continue;
    after.children = [...here.children, ...after.children];
    blocks.splice(i, 1);
  }
  // A table whose every row was removed is gone.
  doc.body.blocks = blocks.filter((b) => b.kind !== "table" || b.rows.length > 0);
  resolveProperties(doc.body.sectPr, ctx);
}

/**
 * Accept or reject tracked changes. With `select`, only revisions whose `w:id`
 * it accepts are resolved. Returns how many revisions were resolved (an
 * insertion split over several runs counts once).
 */
export function resolveRevisions(
  doc: WmlDocument,
  mode: RevisionMode,
  select?: (id: string) => boolean,
): number {
  const ctx: Ctx = { mode, select, resolved: new Set(), stray: 0 };
  resolveBody(doc, ctx);
  return ctx.resolved.size + ctx.stray;
}

/**
 * Accept all tracked changes: insertions stay, deletions go, formatting
 * changes keep the new formatting.
 */
export function acceptAllRevisions(doc: WmlDocument): number {
  return resolveRevisions(doc, "accept");
}

/**
 * Reject all tracked changes: insertions go, deletions come back as normal
 * text, formatting changes restore the recorded old formatting.
 */
export function rejectAllRevisions(doc: WmlDocument): number {
  return resolveRevisions(doc, "reject");
}
