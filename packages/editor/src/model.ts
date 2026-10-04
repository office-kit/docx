/**
 * The editor's stateful controller.
 *
 * An editor is inherently stateful (current document, selection, undo history,
 * subscribers), so unlike `@office-kit/docx`'s function API this is modelled as
 * a small class. It never serializes OOXML itself: every document mutation is
 * performed by a {@link Command} that calls the `@office-kit/docx` public API.
 * The library remains the single source of truth for the document ("one way to
 * do one thing").
 */

import { clone, type Docx } from "@office-kit/docx";
import type { Selection } from "./selection.js";

export interface EditorSnapshot {
  readonly doc: Docx;
  readonly selection: Selection | null;
}

export type ChangeListener = (model: EditorModel) => void;

/** How deep the undo/redo history is allowed to grow. */
const DEFAULT_HISTORY_LIMIT = 200;

export class EditorModel {
  private docState: Docx;
  private selectionState: Selection | null = null;
  private readonly undoStack: EditorSnapshot[] = [];
  private redoStack: EditorSnapshot[] = [];
  // The pre-edit snapshot and the redo history `beginEdit` cleared, kept until
  // the edit commits so an aborted edit can put both back (a failed command
  // must not fork the timeline). Held here rather than read back off the undo
  // stack because `historyLimit` may already have evicted it.
  private pending: { snapshot: EditorSnapshot; redo: EditorSnapshot[] } | null = null;
  private readonly listeners = new Set<ChangeListener>();
  private readonly historyLimit: number;

  constructor(doc: Docx, options: { historyLimit?: number } = {}) {
    this.docState = doc;
    this.historyLimit = options.historyLimit ?? DEFAULT_HISTORY_LIMIT;
  }

  get doc(): Docx {
    return this.docState;
  }

  get selection(): Selection | null {
    return this.selectionState;
  }

  setSelection(selection: Selection | null): void {
    this.selectionState = selection;
    this.emit();
  }

  /**
   * Snapshot the current document onto the undo stack before a mutating command
   * runs. Redo history is cleared because a new edit forks the timeline. The
   * snapshot clones the document so later mutations don't alias it.
   */
  beginEdit(): void {
    // Edits do not nest: a second beginEdit would overwrite the pending
    // snapshot, so a later abortEdit would restore the wrong state.
    if (this.pending) throw new Error("beginEdit() called while another edit is open.");
    const snapshot = { doc: clone(this.docState), selection: this.selectionState };
    this.undoStack.push(snapshot);
    if (this.undoStack.length > this.historyLimit) this.undoStack.shift();
    this.pending = { snapshot, redo: this.redoStack };
    this.redoStack = [];
  }

  /**
   * Roll back the edit opened by {@link beginEdit}: restore the pre-edit
   * document and selection and the redo history it cleared. Unlike
   * {@link undo}, nothing is pushed onto the redo stack, so the failed change
   * cannot be "redone" later.
   */
  abortEdit(): void {
    const pending = this.pending;
    if (!pending) throw new Error("abortEdit() called without a matching beginEdit().");
    if (this.undoStack.at(-1) === pending.snapshot) this.undoStack.pop();
    this.docState = pending.snapshot.doc;
    this.selectionState = pending.snapshot.selection;
    this.redoStack = pending.redo;
    this.pending = null;
    this.emit();
  }

  /** Notify subscribers that the document (or selection) changed. */
  commit(nextSelection?: Selection | null): void {
    this.pending = null;
    if (nextSelection !== undefined) this.selectionState = nextSelection;
    this.docState.dirty = true;
    this.emit();
  }

  canUndo(): boolean {
    return this.undoStack.length > 0;
  }

  canRedo(): boolean {
    return this.redoStack.length > 0;
  }

  undo(): void {
    this.pending = null;
    const prev = this.undoStack.pop();
    if (!prev) return;
    this.redoStack.push({ doc: clone(this.docState), selection: this.selectionState });
    this.docState = prev.doc;
    this.selectionState = prev.selection;
    this.emit();
  }

  redo(): void {
    this.pending = null;
    const next = this.redoStack.pop();
    if (!next) return;
    this.undoStack.push({ doc: clone(this.docState), selection: this.selectionState });
    this.docState = next.doc;
    this.selectionState = next.selection;
    this.emit();
  }

  /** Replace the whole document (e.g. after opening a new file). Clears history. */
  replaceDocument(doc: Docx): void {
    this.docState = doc;
    this.selectionState = null;
    this.undoStack.length = 0;
    this.redoStack = [];
    this.pending = null;
    this.emit();
  }

  subscribe(listener: ChangeListener): () => void {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }

  private emit(): void {
    for (const listener of this.listeners) listener(this);
  }
}
