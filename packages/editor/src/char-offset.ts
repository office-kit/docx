/**
 * Conversions between a {@link DocPosition}'s (run index, offset-in-run) and a
 * paragraph-absolute character offset — the unit `isolateParagraphRunRange`
 * works in. Characters are counted with `runTextLength`, the same unit the
 * canvas renders, so DOM offsets and model offsets agree.
 */

import { runTextLength, type WmlParagraph, type WmlRun } from "@office-kit/docx";
import type { DocPosition } from "./selection.js";

function runsOf(para: WmlParagraph): WmlRun[] {
  return para.children.filter((c): c is WmlRun => c.kind === "run");
}

/** Paragraph-absolute character offset of `pos` (clamped to its run). */
export function absoluteOffset(para: WmlParagraph, pos: DocPosition): number {
  const runs = runsOf(para);
  const inline = pos.inline ?? 0;
  let abs = 0;
  for (const run of runs.slice(0, inline)) abs += runTextLength(run);
  const here = runs[inline];
  return abs + Math.min(pos.offset ?? 0, here ? runTextLength(here) : 0);
}

/**
 * The position of absolute offset `abs`, placed at the *end* of the run that
 * finishes there (so a caret after a deletion stays on the preceding text and
 * inherits its formatting, as in Word). `base` supplies block / cell / para.
 */
export function positionAtOffset(para: WmlParagraph, abs: number, base: DocPosition): DocPosition {
  const runs = runsOf(para);
  let cursor = 0;
  for (const [i, run] of runs.entries()) {
    const len = runTextLength(run);
    if (abs <= cursor + len) return { ...base, inline: i, offset: abs - cursor };
    cursor += len;
  }
  const last = Math.max(runs.length - 1, 0);
  const lastRun = runs[last];
  return { ...base, inline: last, offset: lastRun ? runTextLength(lastRun) : 0 };
}
