/**
 * Honouring document protection in the editor (Restrict Editing): which
 * commands an enforced restriction still allows, and whether the canvas
 * accepts typing at all.
 */

import { documentProtection } from "@office-kit/docx";
import type { Command } from "./commands/types.js";
import type { EditorModel } from "./model.js";

/** Word's message when an edit is refused by protection. */
export const PROTECTED_MESSAGE =
  "This modification is not allowed because the selection is locked.";

// Allowed whatever the restriction: stopping it, and changing only the view.
const ALWAYS_ALLOWED: ReadonlySet<string> = new Set([
  "review.unprotect",
  "view.setView",
  "view.setZoom",
]);
const COMMENT_COMMANDS: ReadonlySet<string> = new Set([
  "review.addComment",
  "review.editComment",
  "review.deleteComment",
]);
// "Tracked changes" lets the reader edit (tracked) but not decide on changes.
const REVIEW_DECISIONS: ReadonlySet<string> = new Set([
  "review.accept",
  "review.reject",
  "review.acceptAll",
  "review.rejectAll",
  "review.acceptAllAndStop",
  "review.rejectAllAndStop",
  "review.trackChanges",
]);
// Direct formatting, refused when formatting is limited to unlocked styles.
const FORMATTING_GROUPS: ReadonlySet<string> = new Set(["text", "paragraph", "list"]);

/**
 * Why the document's enforced protection refuses `command`, or `undefined`
 * when it may run.
 */
export function protectionRefusal(
  model: EditorModel,
  command: Pick<Command<never, unknown>, "id" | "group">,
): string | undefined {
  const protection = documentProtection(model.doc);
  if (!protection?.enforced || ALWAYS_ALLOWED.has(command.id)) return undefined;
  if (protection.formatting && FORMATTING_GROUPS.has(command.group)) return PROTECTED_MESSAGE;
  switch (protection.edit) {
    case "none":
      return undefined;
    case "trackedChanges":
      return REVIEW_DECISIONS.has(command.id) ? PROTECTED_MESSAGE : undefined;
    case "comments":
      return COMMENT_COMMANDS.has(command.id) ? undefined : PROTECTED_MESSAGE;
    default:
      // Read only, and filling in forms (form fields are not editable on the canvas).
      return PROTECTED_MESSAGE;
  }
}

/**
 * Whether protection keeps the canvas from taking typed text: read only,
 * comments only, or forms. Ranges marked as exceptions stay editable.
 */
export function isEditingLocked(model: EditorModel): boolean {
  const protection = documentProtection(model.doc);
  return (
    !!protection?.enforced && protection.edit !== "none" && protection.edit !== "trackedChanges"
  );
}
