/**
 * Find and Replace in the page: running a search with the session's options,
 * moving the selection to a result, and painting every result with the CSS
 * Custom Highlight API (no DOM changes, so the canvas' caret mapping is
 * untouched).
 */

import {
  commands,
  findMatches,
  selectionOf,
  type DocPosition,
  type FindMatch,
  type FindOptions,
} from "@office-kit/docx-editor";
import type { EditorSession } from "../../../session.svelte";
import type { HomeState } from "./state.svelte";

const HIGHLIGHT = "wk-find";
const ZWSP = "​";

export function findOptions(session: EditorSession, home: HomeState): FindOptions {
  const f = home.find;
  return {
    query: session.search,
    matchCase: f.matchCase,
    wholeWords: f.wholeWords,
    wildcards: f.wildcards,
    format: f.format,
  };
}

export interface SearchResult {
  readonly matches: FindMatch[];
  /** Why an invalid wildcard pattern found nothing. */
  readonly error?: string;
}

/** Every result; an invalid pattern finds nothing and says why. */
export function search(session: EditorSession, options: FindOptions): SearchResult {
  const model = session.model;
  const hasFormat = Object.values(options.format ?? {}).some((v) => v !== undefined && v !== "");
  if (!model || (options.query === "" && !hasFormat)) return { matches: [] };
  try {
    return { matches: findMatches(model.doc, options) };
  } catch (err) {
    // An unbalanced wildcard pattern is the user's input, not a bug.
    return { matches: [], error: (err as Error).message };
  }
}

/** Select a result in the page, as Find Next does. */
export function goTo(session: EditorSession, match: FindMatch | undefined): void {
  const model = session.model;
  if (!model || !match) return;
  const sel = selectionOf(model.doc, match);
  if (!sel) return;
  model.setSelection(sel);
  session.changed();
}

export function next(
  session: EditorSession,
  matches: readonly FindMatch[],
  direction: 1 | -1,
): void {
  if (session.model) goTo(session, commands.nextMatch(session.model, matches, direction));
}

function paragraphSelector(pos: DocPosition): string {
  return pos.cell
    ? `[data-wk-block="${pos.block}"][data-wk-cell="${pos.cell.row},${pos.cell.col}"][data-wk-para="${pos.para ?? 0}"]`
    : `[data-wk-block="${pos.block}"]:not([data-wk-cell])`;
}

/** The text node and offset the canvas renders a position at. */
function domPoint(root: ParentNode, pos: DocPosition): { node: Node; offset: number } | undefined {
  const span = root.querySelector(
    `.wk-run${paragraphSelector(pos)}[data-wk-inline="${pos.inline ?? 0}"]`,
  );
  const text = span?.firstChild;
  if (!text || text.nodeType !== Node.TEXT_NODE) return undefined;
  const length = (text.textContent ?? "").replaceAll(ZWSP, "").length;
  return { node: text, offset: Math.min(pos.offset ?? 0, length) };
}

/** Paint every result in the page (Navigation pane's highlight-all). */
export function paintMatches(session: EditorSession, matches: readonly FindMatch[]): void {
  // Browsers without the Highlight API simply show no highlight.
  if (typeof CSS === "undefined" || !("highlights" in CSS)) return;
  const model = session.model;
  const root = document.querySelector(".wk-app .surface");
  if (!model || !root || matches.length === 0) {
    CSS.highlights.delete(HIGHLIGHT);
    return;
  }
  const ranges: Range[] = [];
  for (const match of matches) {
    const sel = selectionOf(model.doc, match);
    const start = sel && domPoint(root, sel.anchor);
    const end = sel && domPoint(root, sel.focus);
    if (!start || !end) continue;
    const range = document.createRange();
    range.setStart(start.node, start.offset);
    range.setEnd(end.node, end.offset);
    ranges.push(range);
  }
  CSS.highlights.set(HIGHLIGHT, new Highlight(...ranges));
}

export function clearPaint(): void {
  if (typeof CSS !== "undefined" && "highlights" in CSS) CSS.highlights.delete(HIGHLIGHT);
}
