/**
 * Review tab commands: comments, Track Changes, accepting and rejecting
 * revisions, proofing language, and document protection.
 */

import {
  acceptAllRevisions,
  acceptRevisions,
  addComment,
  addEditableRange,
  documentProtection,
  listStyles,
  protectDocument,
  type ProtectOptions,
  rejectAllRevisions,
  rejectRevisions,
  removeAllComments,
  removeComment,
  removeEditableRange,
  type RunLanguage,
  setCommentText,
  setDocumentSettingOnOff,
  setRunLanguage,
  setRunOnOff,
  setStyleOnOff,
  setWriteProtection,
  unprotectDocument,
  type WriteProtectionOptions,
} from "@office-kit/docx";
import { absoluteOffset } from "../char-offset.js";
import { paragraphAt } from "../doc-access.js";
import type { EditorModel } from "../model.js";
import { revisionIdsAtSelection } from "../review-nav.js";
import { orderSelection } from "../selection.js";
import { applyToSelectionRuns, overlappingSelectionRuns } from "../selection-runs.js";
import {
  isTrackingRevisions,
  reviewerOf,
  trackedDeleteChar,
  trackRevisionsSetting,
} from "../track-changes.js";
import type { Command } from "./types.js";

function caretParagraph(model: EditorModel) {
  const pos = model.selection?.focus;
  return pos ? paragraphAt(model.doc, pos) : undefined;
}

/** Characters that end a word for "comment on the word at the caret". */
const WORD_BREAK = /[\s\p{P}]/u;

/** The `[start, end)` character range of the word around `offset` in `text`. */
function wordAround(text: string, offset: number): [number, number] {
  let start = offset;
  let end = offset;
  while (start > 0 && !WORD_BREAK.test(text[start - 1] ?? " ")) start--;
  while (end < text.length && !WORD_BREAK.test(text[end] ?? " ")) end++;
  return [start, end];
}

export interface NewCommentParams {
  readonly text: string;
  /** Defaults to the editor's reviewer (see `setReviewer`). */
  readonly author?: string;
  readonly initials?: string;
  readonly date?: string;
}

/**
 * Review ▸ New Comment: comment on the selection, or — with a caret, as
 * Word does — on the word the caret is in. Returns the comment id.
 */
export const addCommentCommand: Command<NewCommentParams, number | undefined> = {
  id: "review.addComment",
  group: "review",
  label: "New Comment",
  run(model, params) {
    const sel = model.selection;
    const para = caretParagraph(model);
    if (!sel || !para) return undefined;
    const reviewer = reviewerOf(model);
    const options = {
      author: params.author ?? reviewer.author,
      initials: params.initials ?? reviewer.initials,
      text: params.text,
      date: params.date ?? new Date().toISOString().replace(/\.\d{3}Z$/, "Z"),
    };
    const { start, end, collapsed } = orderSelection(sel);
    const startPara = paragraphAt(model.doc, start);
    const endPara = paragraphAt(model.doc, end);
    if (!startPara || !endPara) return undefined;
    if (collapsed) {
      const text = startPara.children
        .map((c) =>
          c.kind === "run"
            ? c.pieces
                .map((p) => (p.kind === "text" ? p.value : p.kind === "tab" ? "\t" : ""))
                .join("")
            : "",
        )
        .join("");
      const [from, to] = wordAround(text, absoluteOffset(startPara, start));
      return addComment(model.doc, startPara, { ...options, range: { start: from, end: to } });
    }
    return addComment(model.doc, startPara, {
      ...options,
      range: {
        start: absoluteOffset(startPara, start),
        end: absoluteOffset(endPara, end),
        ...(endPara === startPara ? {} : { endParagraph: endPara }),
      },
    });
  },
  isEnabled: (model) => !!caretParagraph(model),
};

export const editCommentCommand: Command<{ id: number; text: string }> = {
  id: "review.editComment",
  group: "review",
  label: "Edit Comment",
  run(model, { id, text }) {
    if (!setCommentText(model.doc, id, text)) throw new Error(`No comment ${id}.`);
  },
};

export const deleteCommentCommand: Command<{ id: number }> = {
  id: "review.deleteComment",
  group: "review",
  label: "Delete Comment",
  run(model, { id }) {
    if (!removeComment(model.doc, id)) throw new Error(`No comment ${id}.`);
  },
};

export const deleteAllCommentsCommand: Command<void> = {
  id: "review.deleteAllComments",
  group: "review",
  label: "Delete All Comments in Document",
  run(model) {
    removeAllComments(model.doc);
  },
};

function forcedByProtection(model: EditorModel): boolean {
  const protection = documentProtection(model.doc);
  return !!protection?.enforced && protection.edit === "trackedChanges";
}

/** Review ▸ Track Changes (`w:trackRevisions`). */
export const toggleTrackChangesCommand: Command<void> = {
  id: "review.trackChanges",
  group: "review",
  label: "Track Changes",
  run(model) {
    setDocumentSettingOnOff(model.doc, "trackRevisions", !trackRevisionsSetting(model));
  },
  // Protection "tracked changes only" keeps tracking on; the toggle is locked.
  isEnabled: (model) => !forcedByProtection(model),
  isActive: (model) => isTrackingRevisions(model),
};

export const acceptAllRevisionsCommand: Command<void> = {
  id: "review.acceptAll",
  group: "review",
  label: "Accept All Changes",
  run(model) {
    acceptAllRevisions(model.doc);
  },
};

export const rejectAllRevisionsCommand: Command<void> = {
  id: "review.rejectAll",
  group: "review",
  label: "Reject All Changes",
  run(model) {
    rejectAllRevisions(model.doc);
  },
};

/** Accept All Changes and Stop Tracking. */
export const acceptAllAndStopCommand: Command<void> = {
  id: "review.acceptAllAndStop",
  group: "review",
  label: "Accept All Changes and Stop Tracking",
  run(model) {
    acceptAllRevisions(model.doc);
    setDocumentSettingOnOff(model.doc, "trackRevisions", false);
  },
};

/** Reject All Changes and Stop Tracking. */
export const rejectAllAndStopCommand: Command<void> = {
  id: "review.rejectAllAndStop",
  group: "review",
  label: "Reject All Changes and Stop Tracking",
  run(model) {
    rejectAllRevisions(model.doc);
    setDocumentSettingOnOff(model.doc, "trackRevisions", false);
  },
};

/**
 * Accept the given revisions (All Changes Shown), or without ids the change
 * at the selection (Accept This Change). Returns how many were resolved.
 */
export const acceptRevisionsCommand: Command<{ ids?: readonly string[] }, number> = {
  id: "review.accept",
  group: "review",
  label: "Accept This Change",
  run(model, { ids }) {
    return acceptRevisions(model.doc, ids ?? revisionIdsAtSelection(model));
  },
  isEnabled: (model) => revisionIdsAtSelection(model).length > 0,
};

/** Reject the given revisions, or the change at the selection (Reject This Change). */
export const rejectRevisionsCommand: Command<{ ids?: readonly string[] }, number> = {
  id: "review.reject",
  group: "review",
  label: "Reject This Change",
  run(model, { ids }) {
    return rejectRevisions(model.doc, ids ?? revisionIdsAtSelection(model));
  },
  isEnabled: (model) => revisionIdsAtSelection(model).length > 0,
};

/**
 * Backspace (-1) / Delete (+1) at a collapsed caret while Track Changes is
 * on: marks the character (or paragraph mark) deleted instead of removing it.
 */
export const trackedDeleteCommand: Command<{ direction: -1 | 1 }> = {
  id: "review.trackedDelete",
  group: "review",
  label: "Delete",
  run(model, { direction }) {
    trackedDeleteChar(model, direction);
  },
  isEnabled: (model) => isTrackingRevisions(model),
};

export interface LanguageParams {
  readonly language: RunLanguage;
  /** Do not check spelling or grammar (`w:noProof`). */
  readonly noProof: boolean;
}

/** Review ▸ Language ▸ Set Proofing Language, for the selected text. */
export const setProofingLanguageCommand: Command<LanguageParams> = {
  id: "review.language",
  group: "review",
  label: "Language",
  run(model, { language, noProof }) {
    applyToSelectionRuns(model, (run) => {
      setRunLanguage(run, language);
      setRunOnOff(run, "noProof", noProof);
    });
  },
  isEnabled: (model) => overlappingSelectionRuns(model).length > 0,
};

/** Restrict Editing ▸ Start Enforcing Protection. */
export const protectDocumentCommand: Command<ProtectOptions> = {
  id: "review.protect",
  group: "review",
  label: "Protect Document",
  run(model, options) {
    protectDocument(model.doc, options);
  },
};

/** Restrict Editing ▸ Stop Protection. A wrong password throws, leaving the document as it was. */
export const unprotectDocumentCommand: Command<{ password?: string }> = {
  id: "review.unprotect",
  group: "review",
  label: "Stop Protection",
  run(model, { password }) {
    if (!unprotectDocument(model.doc, password ?? "")) {
      throw new Error("The password is incorrect.");
    }
  },
};

/** Always Open Read-Only, and the password to modify. */
export const setWriteProtectionCommand: Command<WriteProtectionOptions | undefined> = {
  id: "review.writeProtection",
  group: "review",
  label: "Always Open Read-Only",
  run(model, options) {
    setWriteProtection(model.doc, options);
  },
};

/**
 * Restrict Editing ▸ Limit formatting to a selection of styles: every style
 * not in `allowed` is locked (`w:locked`).
 */
export const lockStylesCommand: Command<{ allowed: readonly string[] }> = {
  id: "review.lockStyles",
  group: "review",
  label: "Limit formatting to a selection of styles",
  run(model, { allowed }) {
    const keep = new Set(allowed);
    for (const { styleId } of listStyles(model.doc)) {
      setStyleOnOff(model.doc, styleId, "locked", !keep.has(styleId));
    }
  },
};

/** Restrict Editing ▸ Exceptions: let a group edit the selection while protected. */
export const addEditableRangeCommand: Command<{ group?: "everyone"; editor?: string }, string> = {
  id: "review.addEditableRange",
  group: "review",
  label: "Exceptions",
  run(model, who) {
    const sel = model.selection;
    if (!sel) throw new Error("Select the text that stays editable.");
    const { start, end } = orderSelection(sel);
    const startPara = paragraphAt(model.doc, start);
    const endPara = paragraphAt(model.doc, end);
    if (!startPara || !endPara) throw new Error("The selection is not inside paragraphs.");
    return addEditableRange(
      model.doc,
      startPara,
      absoluteOffset(startPara, start),
      absoluteOffset(endPara, end),
      {
        ...(who.editor ? { editor: who.editor } : { group: who.group ?? "everyone" }),
        ...(endPara === startPara ? {} : { endParagraph: endPara }),
      },
    );
  },
  isEnabled: (model) => !!model.selection && !orderSelection(model.selection).collapsed,
};

export const removeEditableRangeCommand: Command<{ id: string }> = {
  id: "review.removeEditableRange",
  group: "review",
  label: "Remove exception",
  run(model, { id }) {
    removeEditableRange(model.doc, id);
  },
};

export const reviewCommands = [
  addCommentCommand,
  editCommentCommand,
  deleteCommentCommand,
  deleteAllCommentsCommand,
  toggleTrackChangesCommand,
  acceptAllRevisionsCommand,
  rejectAllRevisionsCommand,
  acceptAllAndStopCommand,
  rejectAllAndStopCommand,
  acceptRevisionsCommand,
  rejectRevisionsCommand,
  trackedDeleteCommand,
  setProofingLanguageCommand,
  protectDocumentCommand,
  unprotectDocumentCommand,
  setWriteProtectionCommand,
  lockStylesCommand,
  addEditableRangeCommand,
  removeEditableRangeCommand,
];
