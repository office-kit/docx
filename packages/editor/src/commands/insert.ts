/**
 * Insert-tab commands (pages, links, fields, symbols, equations, drop caps,
 * text from file, signature lines). Each one calls the `@office-kit/docx`
 * insert APIs at the caret; a range selection is replaced, as in Word.
 */

import {
  appendPageBreak,
  editHyperlink,
  type DropCapOptions,
  ensureReferenceBookmark,
  type FieldContext,
  type FormFieldOptions,
  type HyperlinkTarget,
  insertBookmark,
  insertCoverPage,
  insertDocumentContent,
  insertEquation,
  insertField,
  insertFormField,
  insertHyperlink,
  insertSignatureLine,
  insertSymbol,
  openDocx,
  paragraphHyperlinks,
  removeBookmark,
  removeCoverPage,
  removeHyperlink,
  runTextLength,
  setDropCap,
  setEquation,
  type SignatureLineOptions,
  splitParagraphAt,
  type SymbolSpec,
  updateFields,
  type WmlInline,
  type WmlParagraph,
  type WmlRun,
} from "@office-kit/docx";
import { absoluteOffset } from "../char-offset.js";
import { paragraphAt } from "../doc-access.js";
import type { EditorModel } from "../model.js";
import { caretAt, type DocPosition, orderSelection } from "../selection.js";
import { caretBlockIndex, caretParagraph, moveLastBlockAfter } from "./insert-util.js";
import { type CoverPagePresetId, coverPagePreset } from "./insert-presets.js";
import { deleteSelectionCommand } from "./structure.js";
import type { Command } from "./types.js";

export interface InsertionPoint {
  readonly paragraph: WmlParagraph;
  /** Paragraph-absolute character offset (the unit `runTextLength` counts). */
  readonly offset: number;
  /** The caret position (block / cell / paragraph of `paragraph`). */
  readonly pos: DocPosition;
}

/**
 * The caret paragraph and offset an inline insert goes to, after deleting a
 * range selection (Word replaces the selection with what you insert).
 */
export function insertionPoint(model: EditorModel): InsertionPoint {
  const sel = model.selection;
  if (!sel) throw new Error("Place the caret in the document first.");
  if (!orderSelection(sel).collapsed) deleteSelectionCommand.run(model);
  const pos = model.selection?.focus;
  const paragraph = pos ? paragraphAt(model.doc, pos) : undefined;
  if (!pos || !paragraph) throw new Error("The caret is not inside a paragraph.");
  return { paragraph, offset: absoluteOffset(paragraph, pos), pos };
}

/**
 * Put the caret right after `inline` (just inserted into `point.paragraph`):
 * at the start of the next run, adding an empty run when the inline is last,
 * so typing continues after — not inside — a field or link.
 */
function caretAfter(model: EditorModel, point: InsertionPoint, inline: WmlInline | undefined): void {
  const children = point.paragraph.children;
  const at = inline ? children.indexOf(inline) : -1;
  let next = children.slice(at + 1).find((c): c is WmlRun => c.kind === "run");
  if (!next) {
    next = { kind: "run", pieces: [], extras: [] };
    children.splice(at + 1, 0, next);
  }
  const inlineIndex = children.filter((c): c is WmlRun => c.kind === "run").indexOf(next);
  model.setSelection(caretAt({ ...point.pos, inline: inlineIndex, offset: 0 }));
}

const hasCaretParagraph = (model: EditorModel): boolean => !!caretParagraph(model);

// --- Pages -----------------------------------------------------------------------

export const insertCoverPageCommand: Command<{ preset: CoverPagePresetId }> = {
  id: "insert.coverPage",
  group: "insert",
  label: "Cover Page",
  run(model, { preset }) {
    insertCoverPage(model.doc, coverPagePreset(model.doc, preset));
    model.setSelection(caretAt({ block: 0, inline: 0, offset: 0 }));
  },
};

export const removeCoverPageCommand: Command<void> = {
  id: "insert.removeCoverPage",
  group: "insert",
  label: "Remove Current Cover Page",
  run(model) {
    if (!removeCoverPage(model.doc)) throw new Error("The document has no cover page.");
    model.setSelection(caretAt({ block: 0, inline: 0, offset: 0 }));
  },
};

/**
 * Blank Page: a page break, an empty paragraph, and another page break after
 * the caret paragraph — the caret lands on the new page, as in Word.
 */
export const insertBlankPageCommand: Command<void> = {
  id: "insert.blankPage",
  group: "insert",
  label: "Blank Page",
  run(model) {
    const at = caretBlockIndex(model.doc, model.selection?.focus.block);
    appendPageBreak(model.doc);
    moveLastBlockAfter(model.doc, at);
    model.doc.document.body.blocks.splice(at + 2, 0, { kind: "paragraph", children: [], extras: [] });
    appendPageBreak(model.doc);
    moveLastBlockAfter(model.doc, at + 2);
    model.setSelection(caretAt({ block: at + 2, inline: 0, offset: 0 }));
  },
};

// --- Links -------------------------------------------------------------------------

/** Insert a link showing `text` at the caret (replacing the selection). */
export const insertLinkCommand: Command<{ text: string; target: HyperlinkTarget }> = {
  id: "insert.link",
  group: "insert",
  label: "Link",
  run(model, { text, target }) {
    const point = insertionPoint(model);
    const link = insertHyperlink(model.doc, point.paragraph, point.offset, text, target);
    const inline = point.paragraph.children.find((c) => c.kind === "raw" && c.node === link);
    caretAfter(model, point, inline);
  },
  isEnabled: hasCaretParagraph,
};

/** The `index`-th hyperlink of the caret paragraph (Edit Link / Remove Link). */
export const editLinkCommand: Command<{ index: number; text?: string; target: HyperlinkTarget }> = {
  id: "insert.editLink",
  group: "insert",
  label: "Edit Link",
  run(model, { index, text, target }) {
    const para = caretParagraph(model);
    const link = para ? paragraphHyperlinks(model.doc, para)[index] : undefined;
    if (!para || !link) throw new Error("There is no link here.");
    editHyperlink(model.doc, para, link.element, target, text);
  },
  isEnabled: hasCaretParagraph,
};

export const removeLinkCommand: Command<{ index: number }> = {
  id: "insert.removeLink",
  group: "insert",
  label: "Remove Link",
  run(model, { index }) {
    const para = caretParagraph(model);
    const link = para ? paragraphHyperlinks(model.doc, para)[index] : undefined;
    if (!para || !link) throw new Error("There is no link here.");
    removeHyperlink(model.doc, para, link.element);
  },
  isEnabled: (model) => {
    const para = caretParagraph(model);
    return !!para && paragraphHyperlinks(model.doc, para).length > 0;
  },
};

/** Bookmark the selection (or the caret point) under `name`. */
export const addBookmarkAtSelectionCommand: Command<{ name: string }> = {
  id: "insert.bookmark",
  group: "insert",
  label: "Bookmark",
  run(model, { name }) {
    const sel = model.selection;
    if (!sel) throw new Error("Place the caret in the document first.");
    const { start, end } = orderSelection(sel);
    const startPara = paragraphAt(model.doc, start);
    const endPara = paragraphAt(model.doc, end);
    if (!startPara || !endPara) throw new Error("A bookmark must start and end in paragraphs.");
    insertBookmark(
      model.doc,
      name,
      { paragraph: startPara, offset: absoluteOffset(startPara, start) },
      { paragraph: endPara, offset: absoluteOffset(endPara, end) },
    );
  },
  isEnabled: hasCaretParagraph,
};

export const deleteBookmarkCommand: Command<{ name: string }> = {
  id: "insert.deleteBookmark",
  group: "insert",
  label: "Delete Bookmark",
  run(model, { name }) {
    if (!removeBookmark(model.doc, name)) throw new Error(`No bookmark named ${name}.`);
  },
};

/** What a cross-reference points at. */
export type CrossReferenceTarget =
  | { readonly kind: "bookmark"; readonly name: string }
  /** A heading / numbered item / caption paragraph: a hidden `_Ref` bookmark is added. */
  | { readonly kind: "paragraph"; readonly paragraph: WmlParagraph }
  /** A footnote / endnote reference mark (the run holding it). */
  | { readonly kind: "note"; readonly paragraph: WmlParagraph; readonly run: WmlRun };

export interface CrossReferenceOptions {
  readonly target: CrossReferenceTarget;
  /** REF (text, paragraph number), PAGEREF (page) or NOTEREF (note number). */
  readonly field: "REF" | "PAGEREF" | "NOTEREF";
  /** Field switches: `\n`, `\r`, `\w`, `\p`, `\f` … */
  readonly switches?: readonly string[];
  /** Insert as hyperlink (`\h`). */
  readonly hyperlink?: boolean;
  /** Include above/below: a second field with `\p` after the first. */
  readonly aboveBelow?: boolean;
  readonly context?: FieldContext;
}

/**
 * Insert ▸ Cross-reference: a REF / PAGEREF / NOTEREF field to a bookmark,
 * creating Word's hidden `_Ref` bookmark around a heading, caption or note
 * reference the first time it is referenced.
 */
export const insertCrossReferenceCommand: Command<CrossReferenceOptions> = {
  id: "insert.crossReference",
  group: "insert",
  label: "Cross-reference",
  run(model, options) {
    const { target } = options;
    const name =
      target.kind === "bookmark"
        ? target.name
        : target.kind === "note"
          ? ensureReferenceBookmark(model.doc, target.paragraph, target.run)
          : ensureReferenceBookmark(model.doc, target.paragraph);
    const switches = [...(options.switches ?? [])];
    if (options.hyperlink) switches.push("\\h");
    const point = insertionPoint(model);
    const instruction = [options.field, name, ...switches].join(" ");
    const context = options.context ?? {};
    let runs = insertField(model.doc, point.paragraph, point.offset, instruction, { context });
    const end = runs[runs.length - 1];
    if (options.aboveBelow && end) {
      // Word writes "<ref> <ref \p>": a space, then the same reference with \p.
      const space: WmlRun = { kind: "run", pieces: [{ kind: "text", value: " ", preserveSpace: true }], extras: [] };
      point.paragraph.children.splice(point.paragraph.children.indexOf(end) + 1, 0, space);
      const at = offsetAfter(point.paragraph, space);
      const flags = options.hyperlink ? " \\p \\h" : " \\p";
      runs = insertField(model.doc, point.paragraph, at, `${options.field} ${name}${flags}`, { context });
    }
    caretAfter(model, point, runs[runs.length - 1]);
  },
  isEnabled: hasCaretParagraph,
};

/** Character offset just after `run` in its paragraph. */
function offsetAfter(para: WmlParagraph, run: WmlRun): number {
  const runs = para.children.filter((c): c is WmlRun => c.kind === "run");
  const index = runs.indexOf(run);
  return absoluteOffset(para, { block: 0, inline: index + 1, offset: 0 });
}

// --- Fields, date & time, symbols ------------------------------------------------------

export const insertFieldAtCaretCommand: Command<{ instruction: string; result?: string; context?: FieldContext }> = {
  id: "insert.field",
  group: "insert",
  label: "Field",
  run(model, { instruction, result, context }) {
    const point = insertionPoint(model);
    const runs = insertField(model.doc, point.paragraph, point.offset, instruction, {
      ...(result !== undefined ? { result } : {}),
      ...(context ? { context } : {}),
    });
    caretAfter(model, point, runs[runs.length - 1]);
  },
  isEnabled: hasCaretParagraph,
};

export const insertFormFieldCommand: Command<FormFieldOptions> = {
  id: "insert.formField",
  group: "insert",
  label: "Form Field",
  run(model, options) {
    const point = insertionPoint(model);
    const runs = insertFormField(model.doc, point.paragraph, point.offset, options);
    caretAfter(model, point, runs[runs.length - 1]);
  },
  isEnabled: hasCaretParagraph,
};

/** Update Field: recompute every computable field result. Returns how many changed. */
export const updateFieldsCommand: Command<FieldContext & { types?: readonly string[] }, number> = {
  id: "insert.updateFields",
  group: "insert",
  label: "Update Field",
  run(model, options) {
    return updateFields(model.doc, options);
  },
};

export const insertSymbolCommand: Command<{ symbol: SymbolSpec }> = {
  id: "insert.symbol",
  group: "insert",
  label: "Symbol",
  run(model, { symbol }) {
    const point = insertionPoint(model);
    const run = insertSymbol(model.doc, point.paragraph, point.offset, symbol);
    caretAfter(model, point, run);
  },
  isEnabled: hasCaretParagraph,
};

// --- Equations, drop caps, signature line ------------------------------------------------

export const insertEquationCommand: Command<{ linear: string; display?: boolean }> = {
  id: "insert.equation",
  group: "insert",
  label: "Equation",
  run(model, { linear, display }) {
    const point = insertionPoint(model);
    // Word inserts a display equation into an empty paragraph, inline otherwise.
    const empty = point.paragraph.children.every((c) => c.kind === "run" && runTextLength(c) === 0);
    const element = insertEquation(model.doc, point.paragraph, point.offset, linear, {
      display: display ?? empty,
    });
    const inline = point.paragraph.children.find((c) => c.kind === "raw" && c.node === element);
    caretAfter(model, point, inline);
  },
  isEnabled: hasCaretParagraph,
};

/** Replace an existing equation (addressed by its paragraph and index there). */
export const editEquationCommand: Command<{ at: DocPosition; index: number; linear: string }> = {
  id: "insert.editEquation",
  group: "insert",
  label: "Edit Equation",
  run(model, { at, index, linear }) {
    const para = paragraphAt(model.doc, at);
    if (!para) throw new Error("The equation's paragraph no longer exists.");
    setEquation(model.doc, para, index, linear);
  },
};

export const dropCapCommand: Command<DropCapOptions> = {
  id: "insert.dropCap",
  group: "insert",
  label: "Drop Cap",
  run(model, options) {
    const pos = model.selection?.focus;
    const para = caretParagraph(model);
    if (!pos || !para) throw new Error("Place the caret in a paragraph.");
    const before = model.doc.document.body.blocks.length;
    setDropCap(model.doc, para, options);
    // The frame paragraph is added / removed before the caret paragraph.
    const shift = model.doc.document.body.blocks.length - before;
    if (!pos.cell) model.setSelection(caretAt({ block: pos.block + shift, inline: 0, offset: 0 }));
  },
  isEnabled: hasCaretParagraph,
};

export const insertSignatureLineCommand: Command<SignatureLineOptions> = {
  id: "insert.signatureLine",
  group: "insert",
  label: "Signature Line",
  run(model, options) {
    const point = insertionPoint(model);
    const run = insertSignatureLine(model.doc, point.paragraph, point.offset, options);
    caretAfter(model, point, run);
  },
  isEnabled: hasCaretParagraph,
};

// --- Object ▸ Text from File --------------------------------------------------------------

/**
 * Insert another .docx's body at the caret: the caret paragraph is split and
 * the content goes between the halves.
 */
export const insertTextFromFileCommand: Command<{ bytes: Uint8Array }, number> = {
  id: "insert.textFromFile",
  group: "insert",
  label: "Text from File",
  run(model, { bytes }) {
    const source = openDocx(bytes);
    const pos = model.selection?.focus;
    if (pos && !pos.cell) {
      const tail = splitParagraphAt(model.doc, pos.block, pos.inline ?? 0, pos.offset ?? 0);
      const at = tail >= 0 ? tail : pos.block + 1;
      const count = insertDocumentContent(model.doc, source, at);
      model.setSelection(caretAt({ block: at + count, inline: 0, offset: 0 }));
      return count;
    }
    const at = caretBlockIndex(model.doc, pos?.block) + 1;
    return insertDocumentContent(model.doc, source, at);
  },
};

export const insertCommands: ReadonlyArray<Command<never, unknown>> = [
  insertCoverPageCommand,
  removeCoverPageCommand,
  insertBlankPageCommand,
  insertLinkCommand,
  editLinkCommand,
  removeLinkCommand,
  addBookmarkAtSelectionCommand,
  deleteBookmarkCommand,
  insertCrossReferenceCommand,
  insertFieldAtCaretCommand,
  insertFormFieldCommand,
  updateFieldsCommand,
  insertSymbolCommand,
  insertEquationCommand,
  editEquationCommand,
  dropCapCommand,
  insertSignatureLineCommand,
  insertTextFromFileCommand,
];
