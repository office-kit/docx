/**
 * Review and View actions shared by the ribbon, the status bar and the panes:
 * moving the caret to a change or comment, which comment or changes a
 * command acts on, and Word's fitted zoom levels.
 */

import { revisions } from "@office-kit/docx";
import {
  adjacentReviewMark,
  caretAt,
  commands,
  type DocPosition,
  pageGeometry,
  reviewMarks,
  type ReviewMark,
} from "@office-kit/docx-editor";
import type { EditorSession } from "./session.svelte";

/** Put the caret at `position` and scroll it into view, as Previous / Next do. */
function goTo(session: EditorSession, position: DocPosition | undefined): boolean {
  const model = session.model;
  if (!model || !position) return false;
  model.setSelection(caretAt(position));
  session.changed();
  // The canvas restores the DOM caret on its next render; scroll once it has.
  requestAnimationFrame(() => {
    window.getSelection()?.focusNode?.parentElement?.scrollIntoView({ block: "center" });
  });
  return true;
}

/** Review ▸ Previous / Next (change or comment), wrapping around like Word. */
export function goToAdjacent(
  session: EditorSession,
  kind: ReviewMark["kind"],
  direction: -1 | 1,
): boolean {
  return (
    !!session.model && goTo(session, adjacentReviewMark(session.model, kind, direction)?.position)
  );
}

/** Jump to a comment or revision by id (a click in the Comments / Reviewing pane). */
export function goToMark(session: EditorSession, kind: ReviewMark["kind"], id: string): boolean {
  const model = session.model;
  if (!model) return false;
  return goTo(
    session,
    reviewMarks(model.doc).find((m) => m.kind === kind && m.id === id)?.position,
  );
}

/** Jump to the start of a top-level block (an Accessibility pane finding). */
export function goToBlock(session: EditorSession, block: number): boolean {
  return goTo(session, { block, inline: 0, offset: 0 });
}

/** View ▸ Zoom's fitted choices. */
export type ZoomFit = "pageWidth" | "textWidth" | "wholePage" | "multiplePages";

const TWIPS_PER_CSS_PX = 15;
// Room around the page inside the surface (its padding and the page shadow).
const SURFACE_GUTTER_PX = 56;
// Gap Word leaves between pages shown side by side.
const PAGE_GAP_PX = 16;

/** The zoom factor that makes the page (or its text) fit the visible surface. */
export function fitZoom(session: EditorSession, fit: ZoomFit): number | undefined {
  const model = session.model;
  const surface = document.querySelector<HTMLElement>(".wk-app .surface");
  if (!model || !surface) return undefined;
  const page = pageGeometry(model.doc);
  const width = surface.clientWidth - SURFACE_GUTTER_PX;
  const height = surface.clientHeight - SURFACE_GUTTER_PX;
  const pageWidth = page.width / TWIPS_PER_CSS_PX;
  const pageHeight = page.height / TWIPS_PER_CSS_PX;
  switch (fit) {
    case "pageWidth":
      return width / pageWidth;
    case "textWidth":
      return width / ((page.width - page.left - page.right) / TWIPS_PER_CSS_PX);
    case "wholePage":
      return Math.min(width / pageWidth, height / pageHeight);
    case "multiplePages":
      return Math.min((width - PAGE_GAP_PX) / (2 * pageWidth), height / pageHeight);
  }
}

/**
 * The comment the caret is in (its highlighted range), else the one selected
 * in the Comments pane: what Review ▸ Delete removes.
 */
export function commentAtCaret(session: EditorSession): number | undefined {
  const node = window.getSelection()?.focusNode;
  const host = node instanceof Element ? node : node?.parentElement;
  const ids = host?.closest<HTMLElement>("[data-wk-comment]")?.dataset["wkComment"];
  const first = ids?.split(" ")[0];
  if (first !== undefined) return Number(first);
  return session.prefs.activeComment ?? undefined;
}

/**
 * Review ▸ New Comment: Word anchors an empty comment and puts the cursor in
 * its card, so the text is typed in the Comments pane.
 */
export function newComment(session: EditorSession): void {
  const id = session.apply(commands.addCommentCommand, { text: "" });
  if (id === undefined) return;
  session.prefs.activeComment = id;
  session.pane.right = "comments";
}

/** Accept / Reject All Changes Shown: the kinds Markup Options currently shows. */
export function revisionIdsShown(session: EditorSession): string[] {
  const model = session.model;
  if (!model) return [];
  const { showFormatting, showInsertionsDeletions } = session.prefs;
  return revisions(model.doc)
    .filter((r) => (r.kind === "format" ? showFormatting : showInsertionsDeletions))
    .map((r) => r.id);
}
