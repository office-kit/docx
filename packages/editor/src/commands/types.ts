/**
 * Command layer types.
 *
 * A command is the *only* way the editor mutates a document. Each command's
 * `run` calls the `@office-kit/docx` public API through {@link EditorModel},
 * wrapping the mutation in the model's undo snapshot. The set of command ids is
 * the single source of truth the coverage ledger validates against: a
 * WordprocessingML element may be classified `edit` only if a real command id
 * here handles it.
 */

import type { EditorModel } from "../model.js";

/** Ribbon groupings, mirroring Word's tabs. */
export type FeatureGroup =
  | "text"
  | "paragraph"
  | "structure"
  | "list"
  | "table"
  | "image"
  | "style"
  | "section"
  | "headerFooter"
  | "references"
  | "review"
  | "shape"
  | "advanced";

/**
 * A command definition. `run` receives the model plus command-specific params;
 * it must call {@link EditorModel.beginEdit} before mutating and
 * {@link EditorModel.commit} after (the {@link runCommand} helper does this).
 */
export interface Command<P = void, R = void> {
  readonly id: string;
  readonly group: FeatureGroup;
  /** Human label for the UI (English; the UI layer localizes). */
  readonly label: string;
  /** Longer description for tooltips / command palette. */
  readonly description?: string;
  /**
   * Perform the mutation and optionally report a result (e.g. a replacement
   * count). Report failures by throwing: {@link runCommand} then rolls the
   * whole edit back.
   */
  run(model: EditorModel, params: P): R;
  /** Whether the command applies to the current selection/state. */
  isEnabled?(model: EditorModel): boolean;
  /** For toggle commands (bold, italic…): whether it is currently on. */
  isActive?(model: EditorModel): boolean;
}

/**
 * Run a command atomically: snapshot for undo, mutate, commit. Returns the
 * command's result, or `undefined` when the command is disabled (nothing ran).
 */
export function runCommand<P, R>(
  model: EditorModel,
  command: Command<P, R>,
  params: P,
): R | undefined {
  if (command.isEnabled && !command.isEnabled(model)) return undefined;
  model.beginEdit();
  let result: R;
  try {
    result = command.run(model, params);
  } catch (err) {
    // Roll back without touching redo history (undo() would push the
    // half-applied document onto it), then rethrow so the failure surfaces.
    model.abortEdit();
    throw err;
  }
  model.commit();
  return result;
}
