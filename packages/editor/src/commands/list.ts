/**
 * List commands: turn the selected paragraphs into a bullet or numbered list,
 * or drop an entire list block. `applyListToParagraph` binds a paragraph to a
 * numbering definition; `addBulletList` / `addNumberedList` seed the numbering
 * part and append fresh list items.
 */

import {
  addBulletList,
  addListDefinition,
  addNumberedList,
  applyListToParagraph,
  childElementsOf,
  ensureBuiltinStyle,
  getElementAttr,
  getParagraphNumbering,
  getParagraphStyle,
  type ListLevel,
  numberingPart,
  restartList,
  setParagraphValProp,
  updateStyleFormatting,
  type WmlParagraph,
} from "@office-kit/docx";
import { paragraphAt, paragraphsInRange } from "../doc-access.js";
import { bodyParagraphs } from "../find.js";
import { createNumberingResolver } from "../list-numbering.js";
import type { EditorModel } from "../model.js";
import { createStyleResolver } from "../resolve.js";
import { orderSelection } from "../selection.js";
import { caretBlockIndex, moveLastBlockAfter } from "./insert-util.js";
import type { Command } from "./types.js";

/** Insert a new bullet/numbered list built from lines, after the caret. */
function insertList(
  model: EditorModel,
  items: readonly string[],
  build: typeof addBulletList,
): void {
  const at = caretBlockIndex(model.doc, model.selection?.focus.block);
  const paras = build(model.doc, items.length ? items : [""]);
  // Builders append each list paragraph to the end; move them as a group.
  for (let i = 0; i < paras.length; i++) moveLastBlockAfter(model.doc, at + i);
}

export const insertBulletListCommand: Command<{ items: readonly string[] }> = {
  id: "list.insertBullet",
  group: "list",
  label: "Bulleted list",
  run(model, { items }) {
    insertList(model, items, addBulletList);
  },
};

export const insertNumberedListCommand: Command<{ items: readonly string[] }> = {
  id: "list.insertNumbered",
  group: "list",
  label: "Numbered list",
  run(model, { items }) {
    insertList(model, items, addNumberedList);
  },
};

/** Bind every selected paragraph to a list (reusing or seeding a numId). */
export const applyListCommand: Command<{ kind: "bullet" | "numbered" }> = {
  id: "list.apply",
  group: "list",
  label: "Apply list",
  run(model, { kind }) {
    const sel = model.selection;
    if (!sel) return;
    const paras = paragraphsInRange(model.doc, orderSelection(sel));
    if (paras.length === 0) return;
    // Seed a list definition of the requested kind, read the numId off the
    // seeded paragraph, bind the real selection to it, then drop the seed.
    const seed =
      kind === "bullet" ? addBulletList(model.doc, [""]) : addNumberedList(model.doc, [""]);
    const seededBlock = seed[0];
    const numId = seededBlock ? getParagraphNumbering(seededBlock)?.numId : undefined;
    if (seededBlock) {
      const idx = model.doc.document.body.blocks.indexOf(seededBlock);
      if (idx >= 0) model.doc.document.body.blocks.splice(idx, 1);
    }
    if (numId === undefined) return;
    for (const p of paras) applyListToParagraph(model.doc, p, numId);
  },
  isEnabled: (model) => {
    const sel = model.selection;
    return !!sel && paragraphsInRange(model.doc, orderSelection(sel)).length > 0;
  },
};

function selectedParagraphs(model: EditorModel): WmlParagraph[] {
  const sel = model.selection;
  return sel ? paragraphsInRange(model.doc, orderSelection(sel)) : [];
}

/** Every paragraph of the body in order (cells included). */
function allParagraphs(model: EditorModel): WmlParagraph[] {
  return bodyParagraphs(model.doc).map((p) => p.para);
}

/**
 * A list instance whose definition has exactly these levels and no
 * overrides, so applying the same library entry again continues that list —
 * as Word continues the last list of the same format.
 */
function matchingList(model: EditorModel, levels: readonly ListLevel[]): number | undefined {
  const part = numberingPart(model.doc);
  if (!part) return undefined;
  const resolver = createNumberingResolver(model.doc);
  for (const num of part.nums) {
    if (childElementsOf(num).some((c) => c.name.local === "lvlOverride")) continue;
    const numId = Number(getElementAttr(num, "numId"));
    if (!Number.isInteger(numId)) continue;
    const same = levels.every((level, ilvl) => {
      const existing = resolver.level(numId, ilvl);
      return existing?.format === level.format && existing.text === level.text;
    });
    if (same) return numId;
  }
  return undefined;
}

/** The list the selected paragraphs all belong to directly, if one. */
function sharedList(paras: readonly WmlParagraph[]): number | undefined {
  const ids = new Set(paras.map((p) => getParagraphNumbering(p)?.numId));
  const [only] = ids;
  return ids.size === 1 && only !== undefined && only !== 0 ? only : undefined;
}

/**
 * Apply a list library entry (Bullets / Numbering / Multilevel List). When the
 * selection is one whole list, the whole list takes the new format; otherwise
 * the selected paragraphs join a list of that format (continuing an existing
 * one, as Word does) at their current level. Levels linked to heading styles
 * also number those styles, so every heading of that level is numbered.
 */
export const applyListPresetCommand: Command<{ levels: readonly ListLevel[] }> = {
  id: "list.applyPreset",
  group: "list",
  label: "Apply list format",
  run(model, { levels }) {
    const paras = selectedParagraphs(model);
    if (paras.length === 0) return;
    const numId = matchingList(model, levels) ?? addListDefinition(model.doc, levels);
    const linked = levels.flatMap((level, ilvl) =>
      level.style ? [{ style: level.style, ilvl }] : [],
    );
    for (const { style, ilvl } of linked) {
      ensureBuiltinStyle(model.doc, style);
      updateStyleFormatting(model.doc, style, {
        paragraph: (view) => applyListToParagraph(model.doc, view, numId, ilvl),
      });
    }
    const linkedStyles = new Set(linked.map((l) => l.style));
    const previous = sharedList(paras);
    const targets =
      previous === undefined
        ? paras
        : allParagraphs(model).filter((p) => getParagraphNumbering(p)?.numId === previous);
    for (const p of targets) {
      const style = getParagraphStyle(p);
      // A paragraph in a linked heading style is numbered through its style.
      if (style !== undefined && linkedStyles.has(style)) {
        setParagraphValProp(p, "numPr", undefined);
        continue;
      }
      applyListToParagraph(model.doc, p, numId, getParagraphNumbering(p)?.ilvl ?? 0);
    }
  },
  isEnabled: (model) => selectedParagraphs(model).length > 0,
};

/** Bullets / Numbering ▸ None: take the selected paragraphs out of their list. */
export const removeListCommand: Command<void> = {
  id: "list.remove",
  group: "list",
  label: "No list",
  run(model) {
    const styles = createStyleResolver(model.doc);
    for (const p of selectedParagraphs(model)) {
      setParagraphValProp(p, "numPr", undefined);
      // A style that numbers the paragraph is overridden with numId 0 (§17.9.18).
      if (styles.paragraph(p).numbering) applyListToParagraph(model.doc, p, 0);
    }
  },
  isEnabled: (model) => selectedParagraphs(model).length > 0,
};

/** The kind of list the selection is in, for the pressed state of Bullets / Numbering. */
export function selectionListKind(model: EditorModel): "bullet" | "numbered" | undefined {
  const paras = selectedParagraphs(model);
  if (paras.length === 0) return undefined;
  const styles = createStyleResolver(model.doc);
  const resolver = createNumberingResolver(model.doc);
  const kinds = new Set(
    paras.map((p) => {
      const numbering = styles.paragraph(p).numbering;
      const level = numbering && resolver.level(numbering.numId, numbering.ilvl);
      if (!level) return undefined;
      return level.format === "bullet" ? "bullet" : "numbered";
    }),
  );
  const [only] = kinds;
  return kinds.size === 1 ? only : undefined;
}

/** Change List Level: move the selected list items to level `ilvl` (0–8). */
export const setListLevelCommand: Command<{ ilvl: number }> = {
  id: "list.level",
  group: "list",
  label: "Change List Level",
  run(model, { ilvl }) {
    if (!Number.isInteger(ilvl) || ilvl < 0 || ilvl > MAX_LEVEL) {
      throw new RangeError(`List level must be 0–${MAX_LEVEL}, got ${String(ilvl)}.`);
    }
    for (const p of selectedParagraphs(model)) {
      const list = getParagraphNumbering(p);
      if (list) applyListToParagraph(model.doc, p, list.numId, ilvl);
    }
  },
  isEnabled: (model) => selectedParagraphs(model).some((p) => !!getParagraphNumbering(p)),
};

const MAX_LEVEL = 8;

function caretList(
  model: EditorModel,
): { para: WmlParagraph; numId: number; ilvl: number } | undefined {
  const pos = model.selection?.focus;
  const para = pos ? paragraphAt(model.doc, pos) : undefined;
  const list = para ? getParagraphNumbering(para) : undefined;
  return para && list && list.numId !== 0 ? { para, ...list } : undefined;
}

/** Move the caret paragraph and the rest of its list (from it onward) to list `numId`. */
function moveListTail(
  model: EditorModel,
  from: WmlParagraph,
  oldNumId: number,
  numId: number,
): void {
  const paras = allParagraphs(model);
  for (const p of paras.slice(paras.indexOf(from))) {
    const list = getParagraphNumbering(p);
    if (list?.numId === oldNumId) applyListToParagraph(model.doc, p, numId, list.ilvl);
  }
}

/**
 * Restart Numbering / Set Numbering Value ▸ Start new list: the caret's list
 * item and the items after it become a new list starting at `start` (the
 * level's own start value when omitted).
 */
export const restartNumberingCommand: Command<{ start?: number }> = {
  id: "list.restart",
  group: "list",
  label: "Restart Numbering",
  run(model, { start }) {
    const list = caretList(model);
    if (!list) return;
    const level = createNumberingResolver(model.doc).level(list.numId, list.ilvl);
    const numId = restartList(model.doc, list.numId, list.ilvl, start ?? level?.start ?? 1);
    if (numId !== undefined) moveListTail(model, list.para, list.numId, numId);
  },
  isEnabled: (model) => !!caretList(model),
};

/**
 * Continue Numbering: the caret's list (from the caret on) joins the nearest
 * earlier list of the same definition, so its numbers carry on.
 */
export const continueNumberingCommand: Command<void> = {
  id: "list.continue",
  group: "list",
  label: "Continue Numbering",
  run(model) {
    const list = caretList(model);
    if (!list) return;
    const resolver = createNumberingResolver(model.doc);
    const abstract = resolver.level(list.numId, list.ilvl)?.abstractNumId;
    const paras = allParagraphs(model);
    const before = paras.slice(0, paras.indexOf(list.para)).toReversed();
    const target = before
      .map((p) => getParagraphNumbering(p))
      .find(
        (n) =>
          n !== undefined &&
          n.numId !== list.numId &&
          n.numId !== 0 &&
          resolver.level(n.numId, list.ilvl)?.abstractNumId === abstract,
      );
    if (target) moveListTail(model, list.para, list.numId, target.numId);
  },
  isEnabled: (model) => !!caretList(model),
};

export const listCommands = [
  insertBulletListCommand,
  insertNumberedListCommand,
  applyListCommand,
  applyListPresetCommand,
  removeListCommand,
  setListLevelCommand,
  restartNumberingCommand,
  continueNumberingCommand,
];
