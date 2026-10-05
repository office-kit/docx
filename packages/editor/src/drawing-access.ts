/**
 * Map between a selected object's run position (what the canvas reports) and
 * the drawing index the `@office-kit/docx` picture APIs take (the position in
 * `imageDrawings` order).
 */

import { type Docx, imageDrawings, type WmlParagraph } from "@office-kit/docx";
import type { DocPosition } from "./selection.js";
import { runAtPath } from "./text-edit.js";

/** The index of the first drawing in the run at `pos`, or -1. */
export function drawingIndexAt(doc: Docx, pos: DocPosition): number {
  const run = runAtPath(doc, pos);
  const piece = run?.pieces.find((p) => p.kind === "drawing");
  return piece?.kind === "drawing" ? imageDrawings(doc).indexOf(piece.node) : -1;
}

/** The run position of drawing `index`, or undefined. */
export function drawingPositionOf(doc: Docx, index: number): DocPosition | undefined {
  const target = imageDrawings(doc)[index];
  if (!target) return undefined;
  const inline = (p: WmlParagraph): number => {
    let runIndex = 0;
    for (const c of p.children) {
      if (c.kind !== "run") continue;
      if (c.pieces.some((piece) => piece.kind === "drawing" && piece.node === target))
        return runIndex;
      runIndex++;
    }
    return -1;
  };
  for (const [block, b] of doc.document.body.blocks.entries()) {
    if (b.kind === "paragraph") {
      const i = inline(b);
      if (i >= 0) return { block, inline: i };
    } else if (b.kind === "table") {
      for (const [row, r] of b.rows.entries()) {
        for (const [col, cell] of r.cells.entries()) {
          for (const [para, p] of cell.paragraphs.entries()) {
            const i = inline(p);
            if (i >= 0) return { block, cell: { row, col }, para, inline: i };
          }
        }
      }
    }
  }
  return undefined;
}
