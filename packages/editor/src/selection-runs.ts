/**
 * Selection-aware run access for character formatting.
 *
 * The naïve approach — format every run the selection *touches* — bolds the
 * whole line when you select a few characters, because runs are the atomic unit
 * of formatting. Instead, {@link applyToSelectionRuns} isolates exactly the
 * selected characters into their own runs (splitting at the boundaries) before
 * applying the change, and updates the selection so the same text stays
 * highlighted. {@link overlappingSelectionRuns} is the read-only counterpart for
 * `isActive` / `isEnabled` state (it never mutates the document).
 */

import {
  type Docx,
  isolateParagraphRunRange,
  runTextLength,
  type WmlParagraph,
  type WmlRun,
} from "@office-kit/docx";
import { bodyOf, paragraphAt, paragraphsInRange } from "./doc-access.js";
import type { EditorModel } from "./model.js";
import { withRunFormatTracking } from "./track-changes.js";
import {
  type CellCoord,
  type DocPosition,
  caretAt,
  type OrderedSelection,
  orderSelection,
  type Selection,
} from "./selection.js";
import { positionAtOffset } from "./char-offset.js";

function paragraphRuns(para: WmlParagraph): WmlRun[] {
  return para.children.filter((c): c is WmlRun => c.kind === "run");
}

/** Absolute character offset of a position within its paragraph's run text. */
function absoluteChar(runs: WmlRun[], inline: number, offset: number): number {
  let n = 0;
  for (let i = 0; i < inline && i < runs.length; i++) n += runTextLength(runs[i]!);
  return n + offset;
}

interface ParaWindow {
  para: WmlParagraph;
  startChar: number;
  endChar: number;
  block: number;
  cell?: CellCoord;
}

/** Whether two positions denote the same collapsed caret. */
function isCollapsed(a: DocPosition, b: DocPosition): boolean {
  return (
    a.block === b.block &&
    (a.inline ?? 0) === (b.inline ?? 0) &&
    (a.offset ?? 0) === (b.offset ?? 0) &&
    a.cell?.row === b.cell?.row &&
    a.cell?.col === b.cell?.col
  );
}

/** The paragraph windows a range covers, with per-paragraph char ranges. */
function rangeWindows(doc: Docx, { start, end }: OrderedSelection): ParaWindow[] {
  // Single paragraph (including inside a table cell): a precise char window.
  if (
    start.block === end.block &&
    start.cell?.row === end.cell?.row &&
    start.cell?.col === end.cell?.col
  ) {
    const para = paragraphAt(doc, start);
    if (!para) return [];
    const runs = paragraphRuns(para);
    return [
      {
        para,
        startChar: absoluteChar(runs, start.inline ?? 0, start.offset ?? 0),
        endChar: absoluteChar(runs, end.inline ?? 0, end.offset ?? 0),
        block: start.block,
        ...(start.cell ? { cell: start.cell } : {}),
      },
    ];
  }

  // Multi-block selection over top-level paragraphs. Tables fall back to whole
  // cell paragraphs (a rare selection shape not worth per-char precision here).
  const out: ParaWindow[] = [];
  const blocks = bodyOf(doc, start).blocks;
  for (let b = start.block; b <= end.block && b < blocks.length; b++) {
    const node = blocks[b];
    if (node?.kind === "paragraph") {
      const runs = paragraphRuns(node);
      const full = runs.reduce((n, r) => n + runTextLength(r), 0);
      out.push({
        para: node,
        startChar: b === start.block ? absoluteChar(runs, start.inline ?? 0, start.offset ?? 0) : 0,
        endChar: b === end.block ? absoluteChar(runs, end.inline ?? 0, end.offset ?? 0) : full,
        block: b,
      });
    } else if (node?.kind === "table") {
      for (const row of node.rows) {
        for (const cell of row.cells) {
          for (const p of cell.paragraphs) {
            const full = paragraphRuns(p).reduce((n, r) => n + runTextLength(r), 0);
            out.push({ para: p, startChar: 0, endChar: full, block: b });
          }
        }
      }
    }
  }
  return out;
}

function selectionWindows(model: EditorModel): ParaWindow[] {
  const sel = model.selection;
  return sel ? rangeWindows(model.doc, orderSelection(sel)) : [];
}

/** A run together with the paragraph that holds it (for style resolution). */
export interface RunRef {
  readonly run: WmlRun;
  readonly para: WmlParagraph;
}

/** Runs with at least one character inside the windows. */
function runRefsCovered(wins: ParaWindow[]): RunRef[] {
  const out: RunRef[] = [];
  for (const w of wins) {
    let cursor = 0;
    for (const run of paragraphRuns(w.para)) {
      const s = cursor;
      const e = cursor + runTextLength(run);
      cursor = e;
      if (e > s && s < w.endChar && e > w.startChar) out.push({ run, para: w.para });
    }
  }
  return out;
}

function runsCovered(wins: ParaWindow[]): WmlRun[] {
  return runRefsCovered(wins).map((ref) => ref.run);
}

/**
 * Every run touched by an ordered selection: the runs the range covers at
 * least one character of, in document order. When the selection is collapsed
 * (a caret) this is the runs of the caret paragraph, letting toggle commands
 * (bold on an empty selection) still report/flip state.
 */
export function runsInRange(doc: Docx, sel: OrderedSelection): WmlRun[] {
  if (sel.collapsed) {
    return paragraphsInRange(doc, sel).flatMap((p) => paragraphRuns(p));
  }
  return runsCovered(rangeWindows(doc, sel));
}

/** Runs overlapping the selection — read-only, for active/enabled state. */
export function overlappingSelectionRuns(model: EditorModel): WmlRun[] {
  return overlappingSelectionRunRefs(model).map((ref) => ref.run);
}

/** {@link overlappingSelectionRuns} with each run's paragraph. */
export function overlappingSelectionRunRefs(model: EditorModel): RunRef[] {
  const sel = model.selection;
  if (!sel) return [];
  const { start, end } = orderSelection(sel);

  // Collapsed caret: the run at the caret (so toggles still report state).
  if (isCollapsed(start, end)) {
    const para = paragraphAt(model.doc, start);
    if (!para) return [];
    const runs = paragraphRuns(para);
    const at = runs[start.inline ?? 0] ?? runs[0];
    return at ? [{ run: at, para }] : [];
  }
  return runRefsCovered(selectionWindows(model));
}

/**
 * Character formatting at a caret, as in Word: inside a word it formats the
 * whole word; anywhere else it formats what is typed next, held by an empty
 * run at the caret (the caret's own run when that is already empty, as after
 * pressing Bold twice).
 */
function applyAtCaret(
  model: EditorModel,
  caret: DocPosition,
  applyFn: (run: WmlRun, para: WmlParagraph) => void,
): void {
  const para = paragraphAt(model.doc, caret);
  if (!para) return;
  const runs = paragraphRuns(para);
  const own = runs[caret.inline ?? 0];
  if (own && runTextLength(own) === 0) {
    withRunFormatTracking(model, own, () => applyFn(own, para));
    return;
  }
  PENDING.delete(model);
  const at = absoluteChar(runs, caret.inline ?? 0, caret.offset ?? 0);
  const word = wordAround(paragraphCharText(runs), at);
  if (word) {
    for (const run of isolateParagraphRunRange(para, word[0], word[1])) {
      withRunFormatTracking(model, run, () => applyFn(run, para));
    }
    model.setSelection(caretAt(positionAtOffset(para, at, caret)));
    return;
  }
  const pending = insertEmptyRunAt(para, at);
  if (!pending) {
    // The caret sits inside a run that cannot be split (a tab or field in
    // it): format that run, as before there was a pending run.
    if (own) withRunFormatTracking(model, own, () => applyFn(own, para));
    return;
  }
  withRunFormatTracking(model, pending, () => applyFn(pending, para));
  PENDING.set(model, { run: pending, para });
  const inline = paragraphRuns(para).indexOf(pending);
  const pos = { ...caret, inline, offset: 0 };
  model.setSelection({ anchor: pos, focus: pos });
}

// The empty run caret formatting made for the text typed next, per editor.
const PENDING = new WeakMap<EditorModel, { run: WmlRun; para: WmlParagraph }>();

/**
 * The caret is moving to `selection`. If it leaves the run caret formatting
 * made for the text typed next (Bold at a caret, say) and nothing was typed
 * there, the run goes, as Word drops pending formatting once the caret moves
 * on. Returns `selection` with the run indices after the removed run shifted,
 * or `undefined` when the document did not change. Call it outside an edit,
 * with positions read before the change.
 */
export function releasePendingFormat(
  model: EditorModel,
  selection: Selection,
): Selection | undefined {
  const held = PENDING.get(model);
  if (!held) return undefined;
  const runs = paragraphRuns(held.para);
  const index = runs.indexOf(held.run);
  const inHeld = (pos: DocPosition): boolean => paragraphAt(model.doc, pos) === held.para;
  if (index < 0 || runTextLength(held.run) > 0) {
    PENDING.delete(model);
    return undefined;
  }
  const onRun = (pos: DocPosition): boolean => inHeld(pos) && (pos.inline ?? 0) === index;
  if (onRun(selection.anchor) || onRun(selection.focus)) return undefined;
  PENDING.delete(model);
  held.para.children = held.para.children.filter((c) => c !== held.run);
  const shift = (pos: DocPosition): DocPosition =>
    inHeld(pos) && (pos.inline ?? 0) > index ? { ...pos, inline: (pos.inline ?? 0) - 1 } : pos;
  return { anchor: shift(selection.anchor), focus: shift(selection.focus) };
}

/** A paragraph's run text, one character per unit `runTextLength` counts. */
function paragraphCharText(runs: WmlRun[]): string {
  return runs
    .flatMap((run) =>
      run.pieces.map((p) => {
        if (p.kind === "text") return p.value;
        // Tabs, breaks and hyphens count one character and never join a word.
        return runTextLength({ kind: "run", pieces: [p], extras: [] }) ? " " : "";
      }),
    )
    .join("");
}

/** The word the character offset `at` falls strictly inside, if any. */
function wordAround(text: string, at: number): [number, number] | undefined {
  for (const seg of new Intl.Segmenter(undefined, { granularity: "word" }).segment(text)) {
    const start = seg.index;
    const end = start + seg.segment.length;
    if (start >= at) return undefined;
    if (at < end) return seg.isWordLike ? [start, end] : undefined;
  }
  return undefined;
}

/**
 * Split the runs at character `at` and put an empty run there carrying the
 * formatting of the text before it (after it at the paragraph start), the
 * formatting typing at that point would take. `undefined` when `at` falls
 * inside a run that cannot be split.
 */
function insertEmptyRunAt(para: WmlParagraph, at: number): WmlRun | undefined {
  isolateParagraphRunRange(para, 0, at);
  let cursor = 0;
  let index = 0;
  let source: WmlRun | undefined;
  for (const [i, child] of para.children.entries()) {
    if (child.kind !== "run") continue;
    if (cursor === at) {
      index = i;
      source ??= child;
      break;
    }
    cursor += runTextLength(child);
    if (cursor > at) return undefined;
    index = i + 1;
    source = child;
  }
  const run: WmlRun = {
    kind: "run",
    ...(source?.rPr ? { rPr: structuredClone(source.rPr) } : {}),
    pieces: [],
    extras: [],
  };
  para.children.splice(index, 0, run);
  return run;
}

/**
 * Apply `applyFn` to exactly the runs covering the selection, splitting runs at
 * the selection boundaries first so a partial selection formats only the
 * selected characters. Keeps the same text selected afterwards.
 */
export function applyToSelectionRuns(
  model: EditorModel,
  applyFn: (run: WmlRun, para: WmlParagraph) => void,
): void {
  const sel = model.selection;
  if (!sel) return;
  const { start, end } = orderSelection(sel);

  if (isCollapsed(start, end)) {
    applyAtCaret(model, start, applyFn);
    return;
  }

  const wins = selectionWindows(model);
  for (const w of wins) {
    for (const run of isolateParagraphRunRange(w.para, w.startChar, w.endChar))
      withRunFormatTracking(model, run, () => applyFn(run, w.para));
  }

  // Re-anchor the selection to the isolated run boundaries (single paragraph).
  if (wins.length === 1) {
    const w = wins[0]!;
    const runs = paragraphRuns(w.para);
    let cursor = 0;
    let startInline = -1;
    let endInline = -1;
    runs.forEach((run, idx) => {
      const s = cursor;
      const e = cursor + runTextLength(run);
      cursor = e;
      if (e > s && s >= w.startChar && e <= w.endChar) {
        if (startInline < 0) startInline = idx;
        endInline = idx;
      }
    });
    if (startInline >= 0 && endInline >= 0) {
      const cell = w.cell ? { cell: w.cell } : {};
      model.setSelection({
        anchor: { block: w.block, ...cell, inline: startInline, offset: 0 },
        focus: {
          block: w.block,
          ...cell,
          inline: endInline,
          offset: runTextLength(runs[endInline]!),
        },
      });
    }
  }
}
