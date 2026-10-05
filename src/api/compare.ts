/**
 * Review ▸ Compare: produce a document that shows the differences between an
 * original and a revised document as tracked changes.
 */

import { serializeXml, type XmlNode } from "../internal/xml/index.js";
import {
  nextAnnotationId,
  revisionAttrs,
  revisionDate,
  revisionElement,
  setParagraphMarkRevision,
  toDeletedPiece,
  type WmlBlock,
  type WmlParagraph,
  type WmlRun,
  type WmlTable,
  type WmlTableRow,
} from "../internal/wordprocessingml/index.js";
import {
  acceptAllRevisions,
  clone,
  type Docx,
  isolateParagraphRunRange,
  runTextLength,
} from "./docx.js";
import type { RevisionAuthor } from "./review.js";

export interface CompareOptions extends RevisionAuthor {}

type Op<T> = { kind: "equal"; a: T; b: T } | { kind: "delete"; a: T } | { kind: "insert"; b: T };

/**
 * Myers' O((N+M)·D) shortest edit script between `a` and `b` by key, so long
 * documents with few changes diff in near-linear time.
 */
function diff<T>(a: readonly T[], b: readonly T[], key: (x: T) => string): Op<T>[] {
  const ak = a.map(key);
  const bk = b.map(key);
  const n = ak.length;
  const m = bk.length;
  const max = n + m;
  const offset = max + 1;
  const v = new Int32Array(2 * max + 3);
  const trace: Int32Array[] = [];
  let found = false;
  for (let d = 0; d <= max && !found; d++) {
    trace.push(v.slice());
    for (let k = -d; k <= d; k += 2) {
      let x =
        k === -d || (k !== d && v[offset + k - 1]! < v[offset + k + 1]!)
          ? v[offset + k + 1]!
          : v[offset + k - 1]! + 1;
      let y = x - k;
      while (x < n && y < m && ak[x] === bk[y]) {
        x++;
        y++;
      }
      v[offset + k] = x;
      if (x >= n && y >= m) {
        found = true;
        break;
      }
    }
  }
  // Walk the trace back from (n, m) to (0, 0).
  const ops: Op<T>[] = [];
  let x = n;
  let y = m;
  for (let d = trace.length - 1; d >= 0; d--) {
    const vd = trace[d]!;
    const k = x - y;
    const prevK =
      k === -d || (k !== d && vd[offset + k - 1]! < vd[offset + k + 1]!) ? k + 1 : k - 1;
    const prevX = d === 0 ? 0 : vd[offset + prevK]!;
    const prevY = prevX - prevK;
    while (x > prevX && y > prevY) {
      x--;
      y--;
      ops.push({ kind: "equal", a: a[x]!, b: b[y]! });
    }
    if (d === 0) break;
    if (x === prevX) ops.push({ kind: "insert", b: b[--y]! });
    else ops.push({ kind: "delete", a: a[--x]! });
  }
  return ops.toReversed();
}

/** A paragraph's text in the canvas' character unit (see `runTextLength`). */
function paragraphChars(para: WmlParagraph): string {
  let out = "";
  for (const child of para.children) {
    if (child.kind !== "run") continue;
    for (const piece of child.pieces) {
      if (piece.kind === "text") out += piece.value;
      else if (piece.kind === "tab") out += "\t";
      else if (piece.kind === "break") out += "\n";
      else if (piece.kind === "noBreakHyphen" || piece.kind === "softHyphen") out += "-";
    }
  }
  return out;
}

function xmlText(node: XmlNode | undefined): string {
  return node?.kind === "element" ? serializeXml({ prologue: [], root: node, epilogue: [] }) : "";
}

function blockKey(block: WmlBlock): string {
  if (block.kind === "paragraph") return `p:${paragraphChars(block)}`;
  if (block.kind === "table") {
    return `t:${block.rows.map((r) => r.cells.map((c) => c.paragraphs.map(paragraphChars).join("\n")).join("\t")).join("\v")}`;
  }
  return `r:${xmlText(block.node)}`;
}

// Words, single East Asian characters (Word compares those one by one),
// whitespace runs, and any other single character.
const TOKEN =
  /[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}]|[^\s\p{P}\p{S}\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}]+|\s+|[^]/gu;

interface Token {
  readonly text: string;
  readonly start: number;
}

function tokenize(text: string): Token[] {
  return [...text.matchAll(TOKEN)].map((m) => ({ text: m[0], start: m.index }));
}

class Builder {
  private nextId: number;
  private readonly date: string;

  constructor(
    doc: Docx,
    private readonly who: RevisionAuthor,
  ) {
    this.nextId = nextAnnotationId(doc.document.body.blocks, doc.document.body.sectPr);
    this.date = revisionDate(who.date);
  }

  attrs() {
    return revisionAttrs(this.nextId++, this.who.author, this.date);
  }

  /** Character-aligned runs of `[start, end)` of a copy of `para`. */
  private slice(para: WmlParagraph, start: number, end: number): WmlRun[] {
    const copy = structuredClone(para);
    return isolateParagraphRunRange(copy, start, end);
  }

  private deleted(runs: readonly WmlRun[], attrs = this.attrs()): WmlRun[] {
    return runs.map((run) => ({
      ...run,
      pieces: run.pieces.map(toDeletedPiece),
      revision: { kind: "del", attrs },
    }));
  }

  /** The child index of character `offset`: before the first run reaching past it. */
  private indexAt(para: WmlParagraph, offset: number): number {
    isolateParagraphRunRange(para, 0, offset);
    let cursor = 0;
    for (const [i, child] of para.children.entries()) {
      if (child.kind !== "run") continue;
      const len = runTextLength(child);
      if (cursor + len > offset) return i;
      cursor += len;
    }
    return para.children.length;
  }

  /** Turn `revised` (a paragraph of the result) into the word-level comparison with `original`. */
  compareParagraph(original: WmlParagraph, revised: WmlParagraph): void {
    const a = tokenize(paragraphChars(original));
    const b = tokenize(paragraphChars(revised));
    const ops = diff(a, b, (t) => t.text);
    // Group consecutive deletes / inserts into hunks, then apply them from
    // the end so earlier offsets stay valid.
    interface Hunk {
      at: number;
      delStart: number;
      delEnd: number;
      insEnd: number;
    }
    const hunks: Hunk[] = [];
    let aPos = 0;
    let bPos = 0;
    let open: Hunk | undefined;
    for (const op of ops) {
      if (op.kind === "equal") {
        open = undefined;
        aPos += op.a.text.length;
        bPos += op.b.text.length;
        continue;
      }
      if (!open) {
        open = { at: bPos, delStart: aPos, delEnd: aPos, insEnd: bPos };
        hunks.push(open);
      }
      if (op.kind === "delete") {
        aPos += op.a.text.length;
        open.delEnd = aPos;
      } else {
        bPos += op.b.text.length;
        open.insEnd = bPos;
      }
    }
    for (const hunk of hunks.toReversed()) {
      if (hunk.insEnd > hunk.at) {
        const attrs = this.attrs();
        for (const run of isolateParagraphRunRange(revised, hunk.at, hunk.insEnd)) {
          run.revision = { kind: "ins", attrs };
        }
      }
      if (hunk.delEnd > hunk.delStart) {
        const removed = this.deleted(this.slice(original, hunk.delStart, hunk.delEnd));
        revised.children.splice(this.indexAt(revised, hunk.at), 0, ...removed);
      }
    }
    if (xmlText(original.pPr) !== xmlText(revised.pPr)) {
      const old = original.pPr?.children.filter(
        (c) => !(c.kind === "element" && (c.name.local === "rPr" || c.name.local === "sectPr")),
      );
      revised.pPr ??= revisionElement("pPr", []);
      (revised.pPr.children as XmlNode[]).push(
        revisionElement("pPrChange", this.attrs(), [revisionElement("pPr", [], old ?? [])]),
      );
    }
  }

  /** Paragraphs inserted whole; their marks are placed by {@link markInsertedParagraphs}. */
  private readonly inserted = new Set<WmlParagraph>();

  insertedParagraph(para: WmlParagraph): void {
    const attrs = this.attrs();
    for (const child of para.children)
      if (child.kind === "run") child.revision = { kind: "ins", attrs };
    this.inserted.add(para);
  }

  /**
   * Record the paragraph marks of inserted paragraphs. An inserted paragraph
   * at the very end has no mark of its own to reject — the last mark of a
   * story always stays — so, as Word does, the mark of the paragraph before
   * it is the inserted one.
   */
  private markInsertedParagraphs(blocks: readonly WmlBlock[]): void {
    const last = blocks.at(-1);
    for (const [i, block] of blocks.entries()) {
      if (block.kind !== "paragraph" || !this.inserted.has(block)) continue;
      if (block !== last) {
        setParagraphMarkRevision(block, "ins", this.attrs());
        continue;
      }
      const before = blocks
        .slice(0, i)
        .findLast((b) => b.kind !== "paragraph" || !this.inserted.has(b));
      if (before?.kind === "paragraph") setParagraphMarkRevision(before, "ins", this.attrs());
    }
  }

  deletedParagraph(original: WmlParagraph): WmlParagraph {
    const copy = structuredClone(original);
    const attrs = this.attrs();
    copy.children = copy.children.map((child) =>
      child.kind === "run" ? (this.deleted([child], attrs)[0] ?? child) : child,
    );
    // The copy must not repeat the original's w14:paraId.
    delete copy.attrs;
    setParagraphMarkRevision(copy, "del", this.attrs());
    return copy;
  }

  private markRow(row: WmlTableRow, kind: "ins" | "del"): void {
    row.trPr ??= revisionElement("trPr", []);
    (row.trPr.children as XmlNode[]).push(revisionElement(kind, this.attrs()));
  }

  insertedTable(table: WmlTable): void {
    for (const row of table.rows) {
      this.markRow(row, "ins");
      for (const cell of row.cells) for (const p of cell.paragraphs) this.insertedParagraph(p);
    }
  }

  deletedTable(original: WmlTable): WmlTable {
    const copy = structuredClone(original);
    delete copy.attrs;
    for (const row of copy.rows) {
      delete row.attrs;
      this.markRow(row, "del");
      for (const cell of row.cells) {
        cell.paragraphs = cell.paragraphs.map((p) => this.deletedParagraph(p));
      }
    }
    return copy;
  }

  /** Same-shaped tables compare cell by cell; otherwise the old one is deleted. */
  compareTable(original: WmlTable, revised: WmlTable): WmlBlock[] {
    const sameShape =
      original.rows.length === revised.rows.length &&
      original.rows.every((row, r) => row.cells.length === revised.rows[r]?.cells.length);
    if (!sameShape) {
      this.insertedTable(revised);
      return [this.deletedTable(original), revised];
    }
    for (const [r, row] of revised.rows.entries()) {
      for (const [c, cell] of row.cells.entries()) {
        const old = original.rows[r]?.cells[c];
        if (old)
          cell.paragraphs = this.compareBlocks(old.paragraphs, cell.paragraphs).filter(
            (b): b is WmlParagraph => b.kind === "paragraph",
          );
      }
    }
    return [revised];
  }

  /** Diff two block lists and return the result's blocks with the changes tracked. */
  compareBlocks(original: readonly WmlBlock[], revised: readonly WmlBlock[]): WmlBlock[] {
    const out: WmlBlock[] = [];
    let deletes: WmlBlock[] = [];
    let inserts: WmlBlock[] = [];
    const flush = (): void => {
      // Pair a hunk's deletions with its insertions in order: a pair of
      // paragraphs (or tables) is a modification, the rest whole blocks.
      const pairs = Math.min(deletes.length, inserts.length);
      for (let i = 0; i < Math.max(deletes.length, inserts.length); i++) {
        const a = deletes[i];
        const b = inserts[i];
        if (i < pairs && a?.kind === "paragraph" && b?.kind === "paragraph") {
          this.compareParagraph(a, b);
          out.push(b);
          continue;
        }
        if (i < pairs && a?.kind === "table" && b?.kind === "table") {
          out.push(...this.compareTable(a, b));
          continue;
        }
        if (a?.kind === "paragraph") out.push(this.deletedParagraph(a));
        else if (a?.kind === "table") out.push(this.deletedTable(a));
        if (b?.kind === "paragraph") this.insertedParagraph(b);
        else if (b?.kind === "table") this.insertedTable(b);
        if (b) out.push(b);
      }
      deletes = [];
      inserts = [];
    };
    for (const op of diff(original, revised, blockKey)) {
      if (op.kind === "equal") {
        flush();
        if (op.a.kind === "paragraph" && op.b.kind === "paragraph")
          this.compareParagraph(op.a, op.b);
        out.push(op.b);
      } else if (op.kind === "delete") deletes.push(op.a);
      else inserts.push(op.b);
    }
    flush();
    this.markInsertedParagraphs(out);
    return out;
  }
}

/**
 * Compare two documents (Review ▸ Compare ▸ Compare Documents): returns a new
 * document — the revised one, with its styles and settings — in which every
 * difference from `original` is a tracked change by `options.author`, at word
 * granularity (East Asian text character by character). Paragraphs and
 * same-shaped tables are compared in place; paragraph-property changes become
 * `<w:pPrChange>`. Pending revisions in either input are accepted first.
 * Character formatting differences inside unchanged text are not reported.
 */
export function compareDocuments(original: Docx, revised: Docx, options: CompareOptions): Docx {
  const base = clone(original);
  const result = clone(revised);
  acceptAllRevisions(base);
  acceptAllRevisions(result);
  const builder = new Builder(result, options);
  result.document.body.blocks = builder.compareBlocks(
    base.document.body.blocks,
    result.document.body.blocks,
  );
  result.dirty = true;
  return result;
}
