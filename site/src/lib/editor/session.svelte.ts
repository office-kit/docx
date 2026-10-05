/**
 * The editor page's shared state: the open document, the chrome around it, and
 * the one way to run a command. Every ribbon tab, pane, and dialog reads the
 * session from Svelte context instead of receiving a dozen props, so a tab is
 * a self-contained component.
 */

import { getContext, setContext } from "svelte";
import {
  caretAt,
  commands,
  runCommand,
  type Command,
  type DocPosition,
  type EditorModel,
  type PageKind,
  type Selection,
  type StoryRef,
} from "@office-kit/docx-editor";
import { text as documentText, type WmlParagraph } from "@office-kit/docx";

// Word's zoom range: 10 % – 500 %.
export const MIN_ZOOM = 0.1;
export const MAX_ZOOM = 5;

/** Word's document views (View ▸ Views). */
export type ViewMode = "print" | "read" | "web" | "outline" | "draft";

/** Review ▸ Display for Review. */
export type MarkupMode = "simple" | "all" | "none" | "original";

/**
 * A floating or inline object the user clicked (picture, shape, chart …). The
 * object-specific contextual tab (Picture Format, Shape Format …) shows while
 * one is selected.
 */
export interface SelectedObject {
  readonly kind: "picture" | "shape" | "textBox" | "chart" | "smartArt" | "equation" | "ink";
  /** The run (or inline) that holds the object. */
  readonly at: DocPosition;
}

/** The header or footer open for editing: which section's, and for which page kind. */
export interface HeaderFooterTarget {
  readonly section: number;
  readonly kind: "header" | "footer";
  readonly type: PageKind;
}

/** What the canvas reports about a laid-out page. */
export interface PageSummary {
  readonly section: number;
  readonly kind: PageKind;
}

/** Panes docked beside the page, in Word's positions. */
export type PaneSide = "left" | "right";

export class EditorSession {
  model = $state<EditorModel | null>(null);
  /**
   * Bumped when the document's structure changes, so the canvas re-renders.
   * Typing inside the canvas does not bump it (the DOM is already current).
   */
  version = $state(0);
  /** Bumped on any edit or selection change; reactive reads depend on it. */
  tick = $state(0);
  fileName = $state("Document1.docx");
  /** One-line message in the status bar (errors from rejected commands). */
  status = $state("");
  /** The open drop-down menu (one at a time, like Word's ribbon). */
  openMenu = $state<string | null>(null);
  /** The open modal dialog, by id (see `Dialog.svelte`). */
  dialog = $state<string | null>(null);
  /** The pane docked on each side, by id (see `panes/registry.ts`). */
  pane = $state<Record<PaneSide, string | null>>({ left: null, right: null });
  /** The search box text, shared by the title bar and the Navigation pane. */
  search = $state("");
  zoom = $state(1);
  showMarks = $state(false);
  viewMode = $state<ViewMode>("print");
  pageMovement = $state<"vertical" | "sideToSide">("vertical");
  showRuler = $state(false);
  showGridlines = $state(false);
  markup = $state<MarkupMode>("all");
  /** Laid-out page count and the page holding the caret, reported by the canvas. */
  pageCount = $state(1);
  currentPage = $state(1);
  selectedObject = $state<SelectedObject | null>(null);
  /** The laid-out pages, reported by the canvas after each layout (Print Layout only). */
  pages = $state<readonly PageSummary[]>([]);
  /** The header/footer being edited (Header & Footer tab), or null. */
  headerFooter = $state<HeaderFooterTarget | null>(null);
  /** Header & Footer ▸ Show Document Text. */
  showDocumentText = $state(true);
  /**
   * The page a body paragraph is laid out on (1-based), set by the canvas;
   * undefined outside Print Layout. Used by field updates (TOC, PAGEREF).
   */
  pageOfParagraph: ((paragraph: WmlParagraph) => number) | undefined = undefined;
  // Where the caret was in the body before a header/footer was opened.
  private bodySelection: Selection | null = null;
  wordCount = $state(0);
  charCount = $state(0);

  /**
   * Run a command and re-render. Commands are atomic: a failure (e.g. rejected
   * input) leaves the document untouched, so it is reported in the status bar
   * rather than crashing the page.
   */
  apply<P, R>(cmd: Command<P, R>, params: P): R | undefined {
    this.openMenu = null;
    const model = this.model;
    if (!model) return undefined;
    try {
      const result = runCommand(model, cmd, params);
      this.status = "";
      return result;
    } catch (err) {
      this.status = `${cmd.label} failed: ${(err as Error).message}`;
      return undefined;
    } finally {
      this.changed();
    }
  }

  /** Whether a toggle command is "on" for the selection (pressed button). */
  active(cmd: Command<never, unknown>): boolean {
    // `tick` is read so the button state recomputes on selection/edit changes.
    return this.tick >= 0 && !!this.model && (cmd.isActive?.(this.model) ?? false);
  }

  /** Whether a selection-based command can run now. */
  enabled(cmd: Command<never, unknown>): boolean {
    return this.tick >= 0 && !!this.model?.selection && (cmd.isEnabled?.(this.model) ?? true);
  }

  /** After a structural change made outside a command (undo, raw XML, …). */
  changed(): void {
    this.version++;
    this.edited();
  }

  /** After an edit the canvas already shows (typing): refresh chrome only. */
  edited(): void {
    this.tick++;
    const model = this.model;
    if (!model) return;
    const text = documentText(model.doc);
    this.charCount = text.length;
    this.wordCount = text.trim().match(/\S+/g)?.length ?? 0;
  }

  /** Swap in another document and reset everything derived from the last one. */
  load(next: EditorModel, fileName: string): void {
    // Word opens a document with the caret at its start.
    if (!next.selection && next.doc.document.body.blocks[0]?.kind === "paragraph") {
      next.setSelection(caretAt({ block: 0, inline: 0, offset: 0 }));
    }
    this.model = next;
    this.headerFooter = null;
    this.bodySelection = null;
    this.fileName = fileName;
    this.status = "";
    this.changed();
  }

  /** The story the selection is in (header, footer, note …), or undefined for the body. */
  get story(): StoryRef | undefined {
    return this.tick >= 0 ? this.model?.story : undefined;
  }

  /**
   * Open the current page's header or footer for editing (Insert ▸ Header ▸
   * Edit Header, or a double-click in the header area), creating it if the
   * section has none.
   */
  editHeaderFooter(kind: "header" | "footer", page = this.currentPage - 1): void {
    const model = this.model;
    if (!model) return;
    const summary = this.pages[page] ?? { section: 0, kind: "default" as const };
    const target: HeaderFooterTarget = { section: summary.section, kind, type: summary.kind };
    if (!model.story) this.bodySelection = model.selection;
    this.apply(commands.editHeaderFooterCommand, target);
    if (model.story) this.headerFooter = target;
  }

  /** Header & Footer ▸ Close Header and Footer (or Esc): back to the body. */
  closeHeaderFooter(): void {
    const model = this.model;
    this.headerFooter = null;
    this.showDocumentText = true;
    if (!model) return;
    model.setSelection(this.bodySelection ?? caretAt({ block: 0, inline: 0, offset: 0 }));
    this.bodySelection = null;
    this.changed();
  }

  undo(): void {
    this.model?.undo();
    this.changed();
  }

  redo(): void {
    this.model?.redo();
    this.changed();
  }

  setZoom(z: number): void {
    this.zoom = Math.max(MIN_ZOOM, Math.min(MAX_ZOOM, Math.round(z * 100) / 100));
  }

  toggleMenu(id: string): void {
    this.openMenu = this.openMenu === id ? null : id;
  }

  /** Show a pane, or hide it if it is already showing (Word's toggle buttons). */
  togglePane(side: PaneSide, id: string): void {
    this.pane[side] = this.pane[side] === id ? null : id;
  }

  openDialog(id: string): void {
    this.openMenu = null;
    this.dialog = id;
  }
}

const KEY = Symbol("editor-session");

export function setSession(session: EditorSession): void {
  setContext(KEY, session);
}

export function getSession(): EditorSession {
  return getContext<EditorSession>(KEY);
}
