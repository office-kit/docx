/**
 * Low-level text-sync helpers used by the canvas to push typed text back into
 * the AST. The canvas renders each run as an editable span; on input it reads
 * the span's text and writes it to the matching run here.
 */

import type { Docx, WmlRun, WmlRunPiece } from "@office-kit/docx";
import { paragraphAt } from "./doc-access.js";
import type { DocPosition } from "./selection.js";

/** Resolve the run addressed by a position's `inline` index. */
export function runAtPath(doc: Docx, pos: DocPosition): WmlRun | undefined {
  const para = paragraphAt(doc, pos);
  if (!para) return undefined;
  const runs = para.children.filter((c): c is WmlRun => c.kind === "run");
  return runs[pos.inline ?? 0];
}

/** Whether a run holds only text and tabs, which the canvas edits as plain text (`\t` for a tab). */
export function isPlainTextRun(run: WmlRun): boolean {
  return run.pieces.every((p) => p.kind === "text" || p.kind === "tab");
}

/** A plain-text run's text, tabs as `\t`. */
export function plainRunText(run: WmlRun): string {
  return run.pieces
    .map((p) => (p.kind === "text" ? p.value : p.kind === "tab" ? "\t" : ""))
    .join("");
}

/**
 * Replace a run's text when it holds only text and tabs; each `\t` becomes a
 * `<w:tab/>`. Returns false (and changes nothing) for runs carrying breaks,
 * drawings, fields, etc. — the canvas leaves those to command-level edits so
 * nothing is silently dropped.
 */
export function setSimpleRunText(run: WmlRun, text: string): boolean {
  if (!isPlainTextRun(run)) return false;
  const pieces: WmlRunPiece[] = [];
  text.split("\t").forEach((value, i) => {
    if (i > 0) pieces.push({ kind: "tab" });
    if (value) pieces.push({ kind: "text", value, preserveSpace: /^\s|\s$|\s\s/.test(value) });
  });
  run.pieces = pieces;
  return true;
}
