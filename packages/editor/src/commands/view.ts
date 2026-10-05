/**
 * View tab commands: the view and zoom saved with the document, and the
 * Outline view's Outlining tools (levels, promote / demote, move up / down).
 */

import {
  type DocumentView,
  type DocumentZoom,
  ensureHeadingStyles,
  setDocumentView,
  setDocumentZoom,
  setParagraphStyle,
  setParagraphValProp,
  type WmlParagraph,
} from "@office-kit/docx";
import { paragraphsInRange } from "../doc-access.js";
import { BODY_TEXT_LEVEL, MAX_HEADING_LEVEL, outlineLevelOf } from "../outline.js";
import type { EditorModel } from "../model.js";
import { orderSelection } from "../selection.js";
import type { Command } from "./types.js";

/** Give paragraphs an outline level the way Word's Outlining tab does: through the Heading styles. */
function setLevel(model: EditorModel, para: WmlParagraph, level: number): void {
  setParagraphValProp(para, "outlineLvl", undefined);
  if (level >= BODY_TEXT_LEVEL) {
    setParagraphStyle(para, undefined);
    return;
  }
  ensureHeadingStyles(model.doc, level);
  setParagraphStyle(para, `Heading${level}`);
}

function selectedTopLevelParagraphs(model: EditorModel): WmlParagraph[] {
  const sel = model.selection;
  return sel ? paragraphsInRange(model.doc, orderSelection(sel)) : [];
}

export const saveDocumentViewCommand: Command<{ view: DocumentView }> = {
  id: "view.setView",
  group: "view",
  label: "View",
  run(model, { view }) {
    setDocumentView(model.doc, view);
  },
};

export const saveDocumentZoomCommand: Command<DocumentZoom> = {
  id: "view.setZoom",
  group: "view",
  label: "Zoom",
  run(model, zoom) {
    setDocumentZoom(model.doc, zoom);
  },
};

/** Outlining ▸ Outline Level: a level 1–9, or {@link BODY_TEXT_LEVEL}. */
export const setOutlineLevelCommand: Command<{ level: number }> = {
  id: "outline.setLevel",
  group: "paragraph",
  label: "Outline Level",
  run(model, { level }) {
    if (!Number.isInteger(level) || level < 1 || level > BODY_TEXT_LEVEL) {
      throw new RangeError(`Outline level must be 1–9 or ${BODY_TEXT_LEVEL}, got ${level}.`);
    }
    for (const para of selectedTopLevelParagraphs(model)) setLevel(model, para, level);
  },
  isEnabled: (model) => selectedTopLevelParagraphs(model).length > 0,
};

/** The level of the nearest heading before `index` (1 when there is none). */
function precedingHeadingLevel(model: EditorModel, para: WmlParagraph): number {
  const blocks = model.doc.document.body.blocks;
  for (let i = blocks.indexOf(para) - 1; i >= 0; i--) {
    const block = blocks[i];
    if (block?.kind !== "paragraph") continue;
    const level = outlineLevelOf(block);
    if (level < BODY_TEXT_LEVEL) return level;
  }
  return 1;
}

/** Outlining ▸ Promote: one level up; Body Text becomes a heading at the level around it. */
export const promoteCommand: Command<void> = {
  id: "outline.promote",
  group: "paragraph",
  label: "Promote",
  run(model) {
    for (const para of selectedTopLevelParagraphs(model)) {
      const level = outlineLevelOf(para);
      const next =
        level >= BODY_TEXT_LEVEL ? precedingHeadingLevel(model, para) : Math.max(level - 1, 1);
      setLevel(model, para, next);
    }
  },
  isEnabled: (model) => selectedTopLevelParagraphs(model).length > 0,
};

/** Outlining ▸ Demote: one level down (Heading 9 stays). */
export const demoteCommand: Command<void> = {
  id: "outline.demote",
  group: "paragraph",
  label: "Demote",
  run(model) {
    for (const para of selectedTopLevelParagraphs(model)) {
      const level = outlineLevelOf(para);
      if (level < MAX_HEADING_LEVEL) setLevel(model, para, level + 1);
    }
  },
  isEnabled: (model) => selectedTopLevelParagraphs(model).length > 0,
};

/** Outlining ▸ Move Up (-1) / Move Down (+1): the selected blocks swap with their neighbour. */
export const moveBlocksCommand: Command<{ direction: -1 | 1 }> = {
  id: "outline.move",
  group: "structure",
  label: "Move",
  run(model, { direction }) {
    const sel = model.selection;
    if (!sel) return;
    const { start, end } = orderSelection(sel);
    const blocks = model.doc.document.body.blocks;
    const neighbour = direction < 0 ? start.block - 1 : end.block + 1;
    const moved = blocks[neighbour];
    if (!moved) return;
    blocks.splice(neighbour, 1);
    blocks.splice(direction < 0 ? end.block : start.block, 0, moved);
    const shift = (p: typeof start) => ({ ...p, block: p.block + direction });
    model.setSelection({ anchor: shift(sel.anchor), focus: shift(sel.focus) });
  },
  isEnabled: (model) => {
    const sel = model.selection;
    if (!sel) return false;
    const { start, end } = orderSelection(sel);
    const count = model.doc.document.body.blocks.length;
    return start.block > 0 || end.block < count - 1;
  },
};

export const viewCommands = [
  saveDocumentViewCommand,
  saveDocumentZoomCommand,
  setOutlineLevelCommand,
  promoteCommand,
  demoteCommand,
  moveBlocksCommand,
];
