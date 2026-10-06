/** List helpers shared by the list and paste commands (not public API). */

import { addBulletList, addNumberedList, type Docx, getParagraphNumbering } from "@office-kit/docx";

/**
 * A new list instance of the given kind, as its numId. The library creates
 * definitions only together with list paragraphs, so one is seeded and the
 * seed paragraph dropped again.
 */
export function newListId(doc: Docx, kind: "bullet" | "numbered"): number | undefined {
  const [seed] = kind === "bullet" ? addBulletList(doc, [""]) : addNumberedList(doc, [""]);
  if (!seed) return undefined;
  const blocks = doc.document.body.blocks;
  const idx = blocks.indexOf(seed);
  if (idx >= 0) blocks.splice(idx, 1);
  return getParagraphNumbering(seed)?.numId;
}
