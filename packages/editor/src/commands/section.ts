/**
 * Section / page-setup commands: page size, margins, orientation, and section
 * breaks. They act on the section holding the caret, the whole document, or
 * — like Word's "Apply to: This point forward" — a new section that starts at
 * the caret.
 */

import {
  insertSectionBreak,
  type PageMargins,
  type PageSize,
  type SectionScope,
  sectionIndexAt,
  type SectionStart,
  setPageMargins,
  setPageOrientation,
  setPageSize,
  splitParagraphAt,
} from "@office-kit/docx";
import type { EditorModel } from "../model.js";
import { caretAt } from "../selection.js";
import type { Command } from "./types.js";

/** Word's "Apply to" choices. */
export type SectionTarget = "document" | "section" | "forward";

/** The section the caret is in (the last one when there is no caret). */
export function caretSection(model: EditorModel): number | undefined {
  const block = model.selection?.focus.block;
  return block === undefined ? undefined : sectionIndexAt(model.doc, block);
}

/**
 * Insert a section break at the caret (Layout ▸ Breaks ▸ Section Breaks):
 * the caret paragraph is split there and the first half ends the section,
 * as in Word. Inside a table the break goes after the table. Returns the
 * index of the section that now starts at the caret.
 */
export function breakSectionAtCaret(model: EditorModel, start: SectionStart): number {
  const pos = model.selection?.focus;
  const doc = model.doc;
  const lastBlock = Math.max(doc.document.body.blocks.length - 1, 0);
  if (!pos) return insertSectionBreak(doc, lastBlock, start);
  if (pos.cell) {
    const section = insertSectionBreak(doc, pos.block, start);
    model.setSelection(caretAt({ block: pos.block + 1 }));
    return section;
  }
  const next = splitParagraphAt(doc, pos.block, pos.inline ?? 0, pos.offset ?? 0);
  if (next < 0) {
    // Not a paragraph (an unmodelled block): the break follows it.
    return insertSectionBreak(doc, pos.block, start);
  }
  const section = insertSectionBreak(doc, pos.block, start);
  model.setSelection(caretAt({ block: next, inline: 0, offset: 0 }));
  return section;
}

/**
 * The sections an "Apply to" choice covers. `forward` first starts a new
 * section at the caret, with `forwardStart` as its start type (Word uses a
 * next-page break for page setup and a continuous one for columns).
 */
export function sectionScopeFor(
  model: EditorModel,
  target: SectionTarget,
  forwardStart: SectionStart,
): SectionScope {
  if (target === "document") return "all";
  if (target === "forward") return breakSectionAtCaret(model, forwardStart);
  return caretSection(model) ?? "all";
}

export const setPageSizeCommand: Command<{ size: PageSize; target?: SectionTarget }> = {
  id: "section.pageSize",
  group: "section",
  label: "Size",
  run(model, { size, target = "section" }) {
    setPageSize(model.doc, size, sectionScopeFor(model, target, "nextPage"));
  },
};

export const setPageMarginsCommand: Command<{ margins: PageMargins; target?: SectionTarget }> = {
  id: "section.pageMargins",
  group: "section",
  label: "Margins",
  run(model, { margins, target = "section" }) {
    setPageMargins(model.doc, margins, sectionScopeFor(model, target, "nextPage"));
  },
};

export const setOrientationCommand: Command<{
  orientation: "portrait" | "landscape";
  target?: SectionTarget;
}> = {
  id: "section.orientation",
  group: "section",
  label: "Orientation",
  run(model, { orientation, target = "section" }) {
    setPageOrientation(model.doc, orientation, sectionScopeFor(model, target, "nextPage"));
  },
};

export const insertSectionBreakCommand: Command<{ type?: SectionStart }> = {
  id: "section.break",
  group: "section",
  label: "Section break",
  run(model, { type = "nextPage" }) {
    breakSectionAtCaret(model, type);
  },
};

export const sectionCommands = [
  setPageSizeCommand,
  setPageMarginsCommand,
  setOrientationCommand,
  insertSectionBreakCommand,
];
