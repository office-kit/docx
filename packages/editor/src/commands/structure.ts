/**
 * Block-structure commands: insert paragraphs / headings / breaks, and remove
 * blocks. Positional inserts use the append-then-move helper so new content
 * lands at the caret rather than the end of the document.
 */

import {
  appendHeading,
  appendLineBreak,
  appendPageBreak,
  appendParagraph,
  type AppendParagraphOptions,
  ensureHeadingStyles,
  isolateParagraphRunRange,
  mergeParagraphIntoPrevious,
  runTextLength,
  paragraphs,
  removeParagraph,
  removeTable,
  splitParagraphAt,
  applyListToParagraph,
  setParagraphStyle,
  setRunFormat,
  setRunValProp,
  type WmlInline,
  type WmlParagraph,
  type WmlRun,
  type WmlTable,
  type WmlTableCell,
  type XmlElement,
} from "@office-kit/docx";
import { absoluteOffset, positionAtOffset } from "../char-offset.js";
import { blockAt, bodyOf, cellAt, paragraphAt } from "../doc-access.js";
import type { EditorModel } from "../model.js";
import { caretAt, type DocPosition, orderSelection } from "../selection.js";
import { WML_NS } from "../wml-ns.js";
import {
  isEmptyParagraph,
  isEmptyRun,
  isPlainTextRun,
  plainRunText,
  runAtPath,
  setSimpleRunText,
} from "../text-edit.js";
import {
  isTrackingRevisions,
  recordSplit,
  trackedDeleteSelection,
  trackedInsertText,
  trackedMergeBack,
} from "../track-changes.js";
import { caretBlockIndex, moveLastBlockAfter } from "./insert-util.js";
import { newListId } from "./list-util.js";
import { insertTableCommand } from "./table.js";
import type { Command } from "./types.js";

/**
 * Split the caret paragraph in two at the caret (the Enter key), replacing a
 * range selection first. Inside a table cell the split stays in the cell. On
 * a block that is neither (an unmodelled raw block) an empty paragraph is
 * inserted after it.
 */
export const splitParagraphCommand: Command<void> = {
  id: "structure.splitParagraph",
  group: "structure",
  label: "Split paragraph",
  run(model) {
    const tracking = isTrackingRevisions(model);
    if (tracking) trackedDeleteSelection(model);
    splitAtCaret(model);
    if (tracking) recordSplit(model);
  },
};

function splitAtCaret(model: EditorModel): void {
  const pos = caretReplacingSelection(model);
  if (!pos) return;
  if (pos.cell) {
    model.setSelection(caretAt(splitCellParagraph(model, pos)));
    return;
  }
  const newBlock = splitParagraphAt(model.doc, pos.block, pos.inline ?? 0, pos.offset ?? 0);
  if (newBlock >= 0) {
    model.setSelection(caretAt({ block: newBlock, inline: 0, offset: 0 }));
  } else {
    const at = caretBlockIndex(model.doc, pos.block);
    appendParagraph(model.doc, "", {});
    moveLastBlockAfter(model.doc, at);
    model.setSelection(caretAt({ block: at + 1 }));
  }
}

/**
 * Merge the caret paragraph into the previous one (Backspace at paragraph
 * start). No-op when there is no preceding paragraph.
 */
export const mergeBackCommand: Command<void> = {
  id: "structure.mergeBack",
  group: "structure",
  label: "Merge with previous paragraph",
  run(model) {
    // With Track Changes on, the paragraph mark is marked deleted instead.
    if (isTrackingRevisions(model) && trackedMergeBack(model)) return;
    const pos = model.selection?.focus;
    if (!pos) return;
    if (pos.cell) {
      // Within a cell: remove the paragraph mark between this paragraph and
      // the previous one of the same cell (all inline content is kept). The
      // first paragraph of a cell has nothing to merge into, as in Word.
      const cell = cellAt(model.doc, pos);
      const paraIndex = pos.para ?? 0;
      const prev = cell?.paragraphs[paraIndex - 1];
      const cur = cell?.paragraphs[paraIndex];
      if (!cell || !prev || !cur) return;
      const caret = positionAtOffset(prev, Number.POSITIVE_INFINITY, {
        ...pos,
        para: paraIndex - 1,
      });
      prev.children = [...prev.children, ...cur.children];
      cell.paragraphs.splice(paraIndex, 1);
      model.setSelection(caretAt(caret));
      return;
    }
    const blocks = bodyOf(model.doc, pos).blocks;
    const prev = blocks[pos.block - 1];
    if (prev?.kind === "table") {
      backIntoTable(model, pos, prev);
      return;
    }
    const joined = mergeParagraphIntoPrevious(model.doc, pos.block);
    if (joined) model.setSelection(caretAt(joined));
  },
  isEnabled: (model) => {
    const pos = model.selection?.focus;
    if (!pos) return false;
    return pos.cell ? (pos.para ?? 0) > 0 : pos.block > 0;
  },
};

/**
 * Backspace at the start of a paragraph that follows a table, as in Word: the
 * text never joins a cell. An empty paragraph is removed (unless it is the
 * last block, which the body must keep) and the caret moves to the end of the
 * table's last cell; a paragraph with text stays as it is.
 */
function backIntoTable(model: EditorModel, pos: DocPosition, table: WmlTable): void {
  const blocks = bodyOf(model.doc, pos).blocks;
  const cur = blocks[pos.block];
  if (cur?.kind !== "paragraph" || !isEmptyParagraph(cur)) return;
  const row = table.rows.length - 1;
  const cells = table.rows[row]?.cells ?? [];
  const col = cells.length - 1;
  const paras = cells[col]?.paragraphs ?? [];
  const last = paras.at(-1);
  if (!last) return;
  if (pos.block < blocks.length - 1) blocks.splice(pos.block, 1);
  const { story } = pos;
  const at = {
    ...(story ? { story } : {}),
    block: pos.block - 1,
    cell: { row, col },
    para: paras.length - 1,
  };
  model.setSelection(caretAt(positionAtOffset(last, Number.POSITIVE_INFINITY, at)));
}

/** Insert `text` at `pos` (a run boundary or inside a text run); returns the caret after it. */
function insertIntoRun(model: EditorModel, pos: DocPosition, text: string): DocPosition {
  const para = paragraphAt(model.doc, pos);
  if (!para) throw new Error("The caret is not inside a paragraph.");
  let run = runAtPath(model.doc, pos);
  if (!run && (pos.inline ?? 0) === 0) {
    // A paragraph with no runs at all (e.g. an empty paragraph from Word).
    run = emptyRun();
    para.children.unshift(run);
  }
  if (!run) throw new Error("The caret run no longer exists.");
  if (text === "") return pos;
  if (isPlainTextRun(run)) {
    const current = plainRunText(run);
    const offset = Math.min(pos.offset ?? 0, current.length);
    setSimpleRunText(run, current.slice(0, offset) + text + current.slice(offset));
    return { ...pos, offset: offset + text.length };
  }
  // A run holding a break / field cannot take text in its middle. At its
  // edges, add a sibling run with the same formatting instead.
  const at = Math.min(pos.offset ?? 0, runTextLength(run));
  if (at !== 0 && at !== runTextLength(run)) {
    throw new Error("Cannot insert text inside a run containing breaks or fields.");
  }
  const sibling: WmlRun = {
    kind: "run",
    ...(run.rPr ? { rPr: structuredClone(run.rPr) } : {}),
    pieces: [],
    extras: [],
  };
  setSimpleRunText(sibling, text);
  para.children.splice(para.children.indexOf(run) + (at === 0 ? 0 : 1), 0, sibling);
  const inline = para.children.filter((c) => c.kind === "run").indexOf(sibling);
  return { ...pos, inline, offset: text.length };
}

function emptyRun(): WmlRun {
  return { kind: "run", pieces: [], extras: [] };
}

/** Give a paragraph an (empty) run if it has none, so the caret has a home. */
function ensureRun(children: WmlInline[]): WmlInline[] {
  return children.some((c) => c.kind === "run") ? children : [...children, emptyRun()];
}

/**
 * Split the paragraph's runs at the caret and return the index in
 * `para.children` where content after the caret begins. Non-run inlines
 * (hyperlinks, fields) sitting exactly at the caret go to the `boundary`
 * side; callers that delete one side pick the side that keeps them, since
 * those inlines are read-only on the canvas. A run with tabs / fields that the
 * caret falls inside of cannot be split and stays whole on the "after" side.
 */
function childIndexAtCaret(
  para: WmlParagraph,
  pos: DocPosition,
  boundary: "before" | "after" = "after",
): number {
  const before = new Set<WmlRun>(isolateParagraphRunRange(para, 0, absoluteOffset(para, pos)));
  // Just past the last run before the caret (zero-length runs are never in
  // `before`, so they cannot cut the "before" side short).
  let index = 0;
  for (const [i, child] of para.children.entries()) {
    if (child.kind === "run" && before.has(child)) index = i + 1;
  }
  if (boundary === "after") return index;
  while (index < para.children.length && para.children[index]?.kind !== "run") index++;
  return index;
}

/**
 * Enter inside a table cell: split the cell paragraph at the caret into two
 * paragraphs of the same cell (the library's `splitParagraphAt` addresses
 * top-level blocks only). The new paragraph copies `pPr` but no `<w:p>`
 * attributes, so `w14:paraId` stays unique.
 */
function splitCellParagraph(model: EditorModel, pos: DocPosition): DocPosition {
  const cell = cellAt(model.doc, pos);
  const paraIndex = pos.para ?? 0;
  const para = cell?.paragraphs[paraIndex];
  if (!cell || !para) throw new Error("The caret is not inside a table cell paragraph.");
  const after = para.children.splice(childIndexAtCaret(para, pos));
  para.children = ensureRun(para.children);
  const next: WmlParagraph = {
    kind: "paragraph",
    ...(para.pPr ? { pPr: structuredClone(para.pPr) } : {}),
    children: after.some((c) => c.kind === "run") ? after : [emptyRun(), ...after],
    extras: [],
  };
  cell.paragraphs.splice(paraIndex + 1, 0, next);
  return {
    block: pos.block,
    ...(pos.cell ? { cell: pos.cell } : {}),
    para: paraIndex + 1,
    inline: 0,
    offset: 0,
  };
}

function samePara(a: DocPosition, b: DocPosition): boolean {
  return (
    a.block === b.block &&
    a.cell?.row === b.cell?.row &&
    a.cell?.col === b.cell?.col &&
    (a.para ?? 0) === (b.para ?? 0)
  );
}

function sectPrOf(para: WmlParagraph): XmlElement | undefined {
  return para.pPr?.children.find(
    (c): c is XmlElement =>
      c.kind === "element" && c.name.uri === WML_NS && c.name.local === "sectPr",
  );
}

/**
 * Delete the selected range and return the collapsed caret where it was.
 * Within one paragraph (also inside a table cell) the covered runs are
 * removed; across top-level paragraphs the boundary paragraphs are split, the
 * blocks in between dropped, and the two remaining halves joined.
 */
function deleteRange(model: EditorModel, start: DocPosition, end: DocPosition): DocPosition {
  const doc = model.doc;
  if (samePara(start, end)) {
    const para = paragraphAt(doc, start);
    if (!para) throw new Error("The selection is not inside a paragraph.");
    const from = absoluteOffset(para, start);
    const removed = new Set<WmlRun>(
      isolateParagraphRunRange(para, from, absoluteOffset(para, end)),
    );
    para.children = ensureRun(para.children.filter((c) => !(c.kind === "run" && removed.has(c))));
    return positionAtOffset(para, from, { ...start });
  }
  if (start.cell || end.cell) {
    const cell = cellAt(doc, start);
    const sameCell =
      cell !== undefined &&
      start.block === end.block &&
      start.cell?.row === end.cell?.row &&
      start.cell?.col === end.cell?.col;
    if (!sameCell) return deleteAcrossCells(model, start, end);
    const first = cell.paragraphs[start.para ?? 0];
    const last = cell.paragraphs[end.para ?? 0];
    if (!first || !last) throw new Error("The selection is not inside the cell's paragraphs.");
    const from = absoluteOffset(first, start);
    const tail = last.children.slice(childIndexAtCaret(last, end));
    first.children = ensureRun([
      ...first.children.slice(0, childIndexAtCaret(first, start, "before")),
      ...tail,
    ]);
    cell.paragraphs.splice((start.para ?? 0) + 1, (end.para ?? 0) - (start.para ?? 0));
    return positionAtOffset(first, from, { ...start });
  }
  // Split the end first so the start position's indices stay valid.
  const tailBlock = splitParagraphAt(doc, end.block, end.inline ?? 0, end.offset ?? 0);
  const headTail = splitParagraphAt(doc, start.block, start.inline ?? 0, start.offset ?? 0);
  if (tailBlock < 0 || headTail < 0) {
    throw new Error("A selection endpoint is not on a top-level paragraph.");
  }
  // Blocks now: [start head] [start rest] [whole blocks…] [end head] [end tail].
  const blocks = doc.document.body.blocks;
  blocks.splice(start.block + 1, end.block - start.block + 1);
  const head = blocks[start.block];
  const tail = blocks[start.block + 1];
  // The end paragraph's section break still closes a section after the join.
  const tailSectPr = tail?.kind === "paragraph" ? sectPrOf(tail) : undefined;
  const joined = mergeParagraphIntoPrevious(doc, start.block + 1);
  if (!joined) throw new Error("Could not join the paragraphs around the selection.");
  if (tailSectPr && head?.kind === "paragraph" && !sectPrOf(head)) {
    head.pPr = head.pPr
      ? { ...head.pPr, children: [...head.pPr.children, tailSectPr] }
      : {
          kind: "element",
          name: { uri: WML_NS, local: "pPr", prefix: "w" },
          attrs: [],
          children: [tailSectPr],
          xmlSpace: "default",
          selfClosing: false,
        };
  }
  return joined;
}

/**
 * Delete a range with an end inside a table, as Word does once a selection
 * leaves a cell: within one table the covered cells are emptied; across the
 * table's edge the rows the range touches go (a table losing every row goes
 * with them), and the text outside the table is cut up to the table.
 */
function deleteAcrossCells(model: EditorModel, start: DocPosition, end: DocPosition): DocPosition {
  const blocks = model.doc.document.body.blocks;
  const startBlock = blocks[start.block];
  const first = start.cell;
  const last = end.cell;
  if (start.block === end.block && first && last && startBlock?.kind === "table") {
    startBlock.rows.forEach((row, r) =>
      row.cells.forEach((cell, c) => {
        const afterFirst = r > first.row || (r === first.row && c >= first.col);
        const beforeLast = r < last.row || (r === last.row && c <= last.col);
        if (afterFirst && beforeLast) clearCell(cell);
      }),
    );
    return { block: start.block, cell: first, para: 0, inline: 0, offset: 0 };
  }
  // Work from the end backwards so earlier block indices stay valid.
  const endBlock = blocks[end.block];
  if (last && endBlock?.kind === "table") {
    endBlock.rows.splice(0, last.row + 1);
    if (endBlock.rows.length === 0) blocks.splice(end.block, 1);
  } else if (endBlock?.kind === "paragraph") {
    cutParagraph(endBlock, 0, absoluteOffset(endBlock, end));
  }
  blocks.splice(start.block + 1, end.block - start.block - 1);
  if (first && startBlock?.kind === "table") {
    startBlock.rows.splice(first.row);
    if (startBlock.rows.length > 0) return startOfBlock(model, start.block + 1);
    blocks.splice(start.block, 1);
    return startOfBlock(model, start.block);
  }
  if (startBlock?.kind !== "paragraph")
    throw new Error("The selection does not start in a paragraph.");
  const from = absoluteOffset(startBlock, start);
  cutParagraph(startBlock, from, paragraphLength(startBlock));
  return positionAtOffset(startBlock, from, { ...start });
}

function clearCell(cell: WmlTableCell): void {
  const [first] = cell.paragraphs;
  if (!first) return;
  first.children = ensureRun([]);
  cell.paragraphs = [first];
}

function cutParagraph(para: WmlParagraph, from: number, to: number): void {
  const removed = new Set<WmlRun>(isolateParagraphRunRange(para, from, to));
  para.children = ensureRun(para.children.filter((c) => !(c.kind === "run" && removed.has(c))));
}

function paragraphLength(para: WmlParagraph): number {
  return para.children.reduce((n, c) => n + (c.kind === "run" ? runTextLength(c) : 0), 0);
}

/** The first caret position in a top-level block (its first cell for a table). */
function startOfBlock(model: EditorModel, index: number): DocPosition {
  const block = model.doc.document.body.blocks[index];
  return block?.kind === "table"
    ? { block: index, cell: { row: 0, col: 0 }, para: 0, inline: 0, offset: 0 }
    : { block: index, inline: 0, offset: 0 };
}

/**
 * The caret an insertion should use: the focus for a collapsed selection, or —
 * as in Word, where typing / Enter over a range replaces it — the caret left
 * after deleting the range.
 */
function caretReplacingSelection(model: EditorModel): DocPosition | undefined {
  const sel = model.selection;
  if (!sel) return undefined;
  const { start, end, collapsed } = orderSelection(sel);
  return collapsed ? sel.focus : deleteRange(model, start, end);
}

/** Delete the selected text (Backspace / Delete / Cut over a range). */
export const deleteSelectionCommand: Command<void> = {
  id: "structure.deleteSelection",
  group: "structure",
  label: "Delete selection",
  run(model) {
    if (isTrackingRevisions(model)) {
      trackedDeleteSelection(model);
      return;
    }
    const sel = model.selection;
    if (!sel) return;
    const { start, end, collapsed } = orderSelection(sel);
    if (collapsed) return;
    model.setSelection(caretAt(deleteRange(model, start, end)));
  },
  isEnabled: (model) => !!model.selection && !orderSelection(model.selection).collapsed,
};

/**
 * Insert plain text at the caret (paste / programmatic typing), replacing the
 * selection when it is a range. Line breaks in `text` split the paragraph, so
 * a multi-line insert is still one undo step.
 */
export const insertTextCommand: Command<{ text: string }> = {
  id: "structure.insertText",
  group: "structure",
  label: "Insert text",
  run(model, { text }) {
    if (isTrackingRevisions(model)) {
      trackedInsertText(model, text);
      return;
    }
    const start = caretReplacingSelection(model);
    if (!start) throw new Error("Insert text needs a caret position.");
    const [first = "", ...rest] = text.split(/\r\n|\r|\n/);
    let caret = insertIntoRun(model, start, first);
    for (const line of rest) caret = insertIntoRun(model, splitAt(model, caret), line);
    model.setSelection(caretAt(caret));
  },
  isEnabled: (model) => !!model.selection,
};

/** Enter at `caret`: split its paragraph (in a cell, within the cell); the start of the second half. */
function splitAt(model: EditorModel, caret: DocPosition): DocPosition {
  if (caret.cell) return splitCellParagraph(model, caret);
  const block = splitParagraphAt(model.doc, caret.block, caret.inline ?? 0, caret.offset ?? 0);
  if (block < 0) throw new Error("The caret is not on a top-level paragraph.");
  return { block, inline: 0, offset: 0 };
}

/** The character formatting a pasted run keeps from its source. */
export interface PastedRunFormat {
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly underline?: boolean;
  readonly strike?: boolean;
  readonly verticalAlign?: "superscript" | "subscript";
  /** Six hex digits without `#`. */
  readonly color?: string;
}

/** Pasted text in one formatting; `\t` is a tab and `\n` a line break. */
export interface PastedRun {
  readonly text: string;
  readonly format: PastedRunFormat;
}

export interface PastedParagraph {
  readonly kind: "paragraph";
  readonly runs: readonly PastedRun[];
  /** Heading 1–9, as the built-in heading style of that level. */
  readonly heading?: number;
  /** A list item; consecutive items of one kind form one list. */
  readonly list?: { readonly kind: "bullet" | "numbered"; readonly level: number };
}

export interface PastedTable {
  readonly kind: "table";
  /** Rows of cells, each cell its paragraphs. */
  readonly rows: ReadonlyArray<ReadonlyArray<readonly PastedParagraph[]>>;
}

/** Formatted content from the clipboard (see `parseClipboardHtml`). */
export type PastedBlock = PastedParagraph | PastedTable;

/** The list the previous pasted paragraph went into, which the next item of its kind continues. */
type OpenList = { readonly kind: "bullet" | "numbered"; readonly numId: number } | undefined;

/**
 * Paste formatted content at the caret, replacing the selection, as one undo
 * step. As in Word, the first pasted paragraph continues the caret's
 * paragraph and the rest of that paragraph follows the last one; a pasted
 * table goes between the two halves. Inside a table cell, pasted tables
 * become their cells' paragraphs (nested tables are not supported).
 */
export const insertFragmentCommand: Command<{ blocks: readonly PastedBlock[] }> = {
  id: "structure.insertFragment",
  group: "structure",
  label: "Paste",
  run(model, { blocks }) {
    if (isTrackingRevisions(model)) {
      // A tracked paste is recorded as inserted text; its formatting would
      // need tracked property changes of its own.
      trackedInsertText(model, fragmentText(blocks));
      return;
    }
    let caret = caretReplacingSelection(model);
    if (!caret) throw new Error("Paste needs a caret position.");
    let merge = true;
    let list: OpenList;
    const pasted = caret.cell ? flattenTables(blocks) : blocks;
    // Text pasted within one paragraph takes the destination's paragraph
    // formatting; pasted paragraphs bring their own (in Word, their marks).
    const inline = pasted.length === 1 && pasted[0]?.kind === "paragraph";
    for (const block of pasted) {
      if (block.kind === "table") {
        caret = insertPastedTable(model, caret, block);
        merge = true;
        list = undefined;
        continue;
      }
      if (!merge) caret = splitAt(model, caret);
      const para = paragraphAt(model.doc, caret);
      if (!para) throw new Error("The caret is not inside a paragraph.");
      const fresh = !inline || isEmptyParagraph(para);
      caret = insertPastedRuns(para, caret, block.runs);
      if (fresh) list = formatPastedParagraph(model, para, block, list);
      merge = false;
    }
    model.setSelection(caretAt(caret));
  },
  isEnabled: (model) => !!model.selection,
};

function pastedText(p: PastedParagraph): string {
  return p.runs.map((r) => r.text).join("");
}

function fragmentText(blocks: readonly PastedBlock[]): string {
  return blocks
    .map((b) =>
      b.kind === "paragraph"
        ? pastedText(b)
        : b.rows
            .map((row) => row.map((cell) => cell.map(pastedText).join(" ")).join("\t"))
            .join("\n"),
    )
    .join("\n");
}

function flattenTables(blocks: readonly PastedBlock[]): PastedParagraph[] {
  return blocks.flatMap((b) => (b.kind === "paragraph" ? [b] : b.rows.flat(2)));
}

function buildPastedRun({ text, format }: PastedRun): WmlRun {
  const run: WmlRun = { kind: "run", pieces: [], extras: [] };
  for (const [i, line] of text.split("\n").entries()) {
    if (i > 0) run.pieces.push({ kind: "break" });
    const part: WmlRun = { kind: "run", pieces: [], extras: [] };
    setSimpleRunText(part, line);
    run.pieces.push(...part.pieces);
  }
  setRunFormat(run, {
    ...(format.bold ? { bold: true } : {}),
    ...(format.italic ? { italic: true } : {}),
    ...(format.strike ? { strike: true } : {}),
    ...(format.underline ? { underline: "single" } : {}),
    ...(format.color ? { color: format.color } : {}),
  });
  if (format.verticalAlign) setRunValProp(run, "vertAlign", format.verticalAlign);
  return run;
}

/** Put the runs in at the caret; the caret after them. */
function insertPastedRuns(
  para: WmlParagraph,
  caret: DocPosition,
  runs: readonly PastedRun[],
): DocPosition {
  const built = runs.filter((r) => r.text !== "").map(buildPastedRun);
  const last = built.at(-1);
  if (!last) return caret;
  // The empty runs an empty paragraph or a split leaves would linger next to
  // the pasted ones.
  const at = absoluteOffset(para, caret);
  para.children = para.children.filter((c) => !(c.kind === "run" && isEmptyRun(c)));
  para.children.splice(childIndexAtCaret(para, positionAtOffset(para, at, caret)), 0, ...built);
  const inline = para.children.filter((c) => c.kind === "run").indexOf(last);
  return { ...caret, inline, offset: runTextLength(last) };
}

/** Give a pasted paragraph its heading style or list; returns the list it is in. */
function formatPastedParagraph(
  model: EditorModel,
  para: WmlParagraph,
  pasted: PastedParagraph,
  open: OpenList,
): OpenList {
  if (pasted.heading !== undefined) {
    ensureHeadingStyles(model.doc, pasted.heading);
    setParagraphStyle(para, `Heading${pasted.heading}`);
  }
  if (!pasted.list) return undefined;
  const { kind, level } = pasted.list;
  const numId = open?.kind === kind ? open.numId : newListId(model.doc, kind);
  if (numId === undefined) return undefined;
  applyListToParagraph(model.doc, para, numId, level);
  return { kind, numId };
}

/**
 * A pasted table between the halves of the caret's paragraph (a caret at its
 * start leaves no empty half before it); the caret moves to the second half.
 */
function insertPastedTable(
  model: EditorModel,
  caret: DocPosition,
  pasted: PastedTable,
): DocPosition {
  const rows = pasted.rows.length;
  const cols = Math.max(0, ...pasted.rows.map((row) => row.length));
  if (rows === 0 || cols === 0) return caret;
  splitAt(model, caret);
  model.setSelection(caretAt({ ...caret, inline: 0, offset: 0 }));
  insertTableCommand.run(model, { rows, cols });
  const blocks = model.doc.document.body.blocks;
  const table = blocks[caret.block + 1];
  if (table?.kind !== "table") throw new Error("The pasted table was not inserted.");
  let list: OpenList;
  table.rows.forEach((row, r) =>
    row.cells.forEach((cell, c) => {
      const paras = pasted.rows[r]?.[c] ?? [];
      if (paras.length === 0) return;
      cell.paragraphs = paras.map((p) => {
        const para: WmlParagraph = { kind: "paragraph", children: [], extras: [] };
        insertPastedRuns(para, { block: 0, inline: 0, offset: 0 }, p.runs);
        para.children = ensureRun(para.children);
        list = formatPastedParagraph(model, para, p, list);
        return para;
      });
    }),
  );
  const head = blocks[caret.block];
  let tail = caret.block + 2;
  if (head?.kind === "paragraph" && isEmptyParagraph(head) && !sectPrOf(head)) {
    blocks.splice(caret.block, 1);
    tail--;
  }
  return { ...(caret.story ? { story: caret.story } : {}), block: tail, inline: 0, offset: 0 };
}

export const insertParagraphCommand: Command<{ text?: string; options?: AppendParagraphOptions }> =
  {
    id: "structure.insertParagraph",
    group: "structure",
    label: "Insert paragraph",
    run(model, params) {
      const at = caretBlockIndex(model.doc, model.selection?.focus.block);
      appendParagraph(model.doc, params.text ?? "", params.options ?? {});
      moveLastBlockAfter(model.doc, at);
      model.setSelection(caretAt({ block: at + 1 }));
    },
  };

export const insertHeadingCommand: Command<{ text: string; level?: number }> = {
  id: "structure.insertHeading",
  group: "structure",
  label: "Insert heading",
  run(model, { text, level = 1 }) {
    // Define the built-in heading styles so the heading's pStyle reference is
    // not dangling (Word would otherwise fall back to Normal).
    ensureHeadingStyles(model.doc);
    const at = caretBlockIndex(model.doc, model.selection?.focus.block);
    appendHeading(model.doc, text, level);
    moveLastBlockAfter(model.doc, at);
    model.setSelection(caretAt({ block: at + 1 }));
  },
};

export const insertPageBreakCommand: Command<void> = {
  id: "structure.insertPageBreak",
  group: "structure",
  label: "Page break",
  run(model) {
    const at = caretBlockIndex(model.doc, model.selection?.focus.block);
    appendPageBreak(model.doc);
    moveLastBlockAfter(model.doc, at);
  },
};

export const insertLineBreakCommand: Command<{
  kind?: "line" | "page" | "column" | "textWrapping";
}> = {
  id: "structure.insertLineBreak",
  group: "structure",
  label: "Line break",
  run(model, { kind = "line" }) {
    const pos = caretReplacingSelection(model);
    const para = pos ? paragraphAt(model.doc, pos) : undefined;
    if (!pos || !para) return;
    // Move the break run the library appended onto the caret boundary
    // (append-then-move, as for blocks).
    const index = childIndexAtCaret(para, pos);
    const breakRun = appendLineBreak(model.doc, para, kind);
    para.children.pop();
    para.children.splice(index, 0, breakRun);
    // Caret goes after the break: on the following run, or a new empty one.
    const runs = para.children.filter((c): c is WmlRun => c.kind === "run");
    let breakInline = runs.indexOf(breakRun);
    if (breakInline === runs.length - 1) {
      para.children.splice(index + 1, 0, emptyRun());
    }
    breakInline += 1;
    model.setSelection(caretAt({ ...pos, inline: breakInline, offset: 0 }));
  },
  isEnabled: (model) => {
    const pos = model.selection?.focus;
    return !!(pos && paragraphAt(model.doc, pos));
  },
};

export const deleteBlockCommand: Command<void> = {
  id: "structure.deleteBlock",
  group: "structure",
  label: "Delete block",
  run(model) {
    const block = model.selection?.focus.block;
    if (block === undefined) return;
    const node = blockAt(model.doc, block);
    if (!node) return;
    if (node.kind === "table") {
      const tableIndex = countKind(model, "table", block);
      removeTable(model.doc, tableIndex);
    } else if (node.kind === "paragraph") {
      const paraIndex = paragraphs(model.doc).indexOf(node);
      if (paraIndex >= 0) removeParagraph(model.doc, paraIndex);
    }
    model.setSelection(caretAt({ block: Math.max(block - 1, 0) }));
  },
  isEnabled: (model) => {
    const block = model.selection?.focus.block;
    return block !== undefined && !!blockAt(model.doc, block);
  },
};

/** Index of a block among blocks of the same kind up to `blockIndex`. */
function countKind(model: EditorModel, kind: "table" | "paragraph", blockIndex: number): number {
  const list = model.doc.document.body.blocks;
  let n = 0;
  for (let i = 0; i < blockIndex && i < list.length; i++) {
    if (list[i]?.kind === kind) n++;
  }
  return n;
}

export const structureCommands = [
  insertTextCommand,
  insertFragmentCommand,
  insertParagraphCommand,
  splitParagraphCommand,
  mergeBackCommand,
  insertHeadingCommand,
  insertPageBreakCommand,
  insertLineBreakCommand,
  deleteBlockCommand,
];
