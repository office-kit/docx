/**
 * Helpers for inserting content at the caret.
 *
 * Block builders (`appendParagraph`, `addTable`, …) push a node to the end of
 * the body; {@link moveLastBlockAfter} moves it to the caret. Inline inserts
 * (fields, symbols, links, equations) take the caret paragraph and a
 * character offset (see `insertionPoint` in ./insert.ts).
 */

import type { Docx, WmlBlock, WmlParagraph } from "@office-kit/docx";
import { paragraphAt } from "../doc-access.js";
import type { EditorModel } from "../model.js";

/** Move the block the builder just appended to sit right after `afterIndex`. */
export function moveLastBlockAfter(doc: Docx, afterIndex: number): void {
  const list = doc.document.body.blocks;
  if (list.length === 0) return;
  const node = list.pop() as WmlBlock;
  const target = Math.min(Math.max(afterIndex + 1, 0), list.length);
  list.splice(target, 0, node);
}

/** The block index the caret currently sits on, or the last block. */
export function caretBlockIndex(doc: Docx, block: number | undefined): number {
  const count = doc.document.body.blocks.length;
  if (block === undefined) return Math.max(count - 1, 0);
  return Math.min(Math.max(block, 0), Math.max(count - 1, 0));
}

/** The caret paragraph without touching the selection, or undefined. */
export function caretParagraph(model: EditorModel): WmlParagraph | undefined {
  const pos = model.selection?.focus;
  return pos ? paragraphAt(model.doc, pos) : undefined;
}
