/**
 * Document position ↔ DOM on the paged canvas. A position can have several
 * DOM copies: a header repeats on every page, and a paragraph split across
 * pages is shown in two clipped windows. The caret goes to the copy where
 * the position is actually visible (on the preferred page for headers).
 */

import { type DocPosition, runPoint, STORY_ATTR, storyKeyOf } from "@office-kit/docx-editor";

/** Attribute selector for the paragraph a position lives in. */
export function paragraphSelector(pos: DocPosition): string {
  return pos.cell
    ? `[data-wk-block="${pos.block}"][data-wk-cell="${pos.cell.row},${pos.cell.col}"][data-wk-para="${pos.para ?? 0}"]`
    : `[data-wk-block="${pos.block}"]:not([data-wk-cell])`;
}

/** The containers a position's story is rendered in (the root for the body). */
function scopes(root: HTMLElement, pos: DocPosition): HTMLElement[] {
  const key = storyKeyOf(pos);
  if (!key) return [root];
  return [...root.querySelectorAll<HTMLElement>(`[${STORY_ATTR}]`)].filter(
    (el) => el.getAttribute(STORY_ATTR) === key,
  );
}

/** Whether an element belongs to the position's story (and not a nested one). */
function inStory(el: Element, pos: DocPosition): boolean {
  const owner = el.closest(`[${STORY_ATTR}]`)?.getAttribute(STORY_ATTR) ?? "";
  return owner === storyKeyOf(pos);
}

/** Every DOM copy of a selector within the position's story. */
export function copiesOf(root: HTMLElement, pos: DocPosition, selector: string): HTMLElement[] {
  return scopes(root, pos).flatMap((scope) =>
    [...scope.querySelectorAll<HTMLElement>(selector)].filter((el) => inStory(el, pos)),
  );
}

/** Run span copies for a position, or the paragraph copies when it has no run. */
function hosts(root: HTMLElement, pos: DocPosition): HTMLElement[] {
  const spans = copiesOf(
    root,
    pos,
    `.wk-run${paragraphSelector(pos)}[data-wk-inline="${pos.inline ?? 0}"]`,
  );
  return spans.length ? spans : copiesOf(root, pos, `.wk-p${paragraphSelector(pos)}`);
}

export interface DomPoint {
  readonly node: Node;
  readonly offset: number;
}

function pointIn(host: HTMLElement, pos: DocPosition): DomPoint {
  return host.classList.contains("wk-run")
    ? runPoint(host, pos.offset ?? 0)
    : { node: host, offset: 0 };
}

function visible(point: DomPoint, host: HTMLElement): boolean {
  const clip = host.closest(".wk-slice");
  if (!clip) return true;
  const range = document.createRange();
  range.setStart(point.node, point.offset);
  range.collapse(true);
  const rect = range.getClientRects()[0] ?? host.getBoundingClientRect();
  const box = clip.getBoundingClientRect();
  const mid = (rect.top + rect.bottom) / 2;
  return mid >= box.top && mid <= box.bottom;
}

/**
 * The DOM point for a position: on `preferPage` when it has a copy there
 * (headers, footers), else the first copy where it is visible.
 */
export function domPoint(
  root: HTMLElement,
  pos: DocPosition,
  preferPage?: number,
): DomPoint | null {
  const candidates = hosts(root, pos);
  if (preferPage !== undefined) {
    const onPage = candidates.find(
      (h) => h.closest<HTMLElement>(".wk-pagebox")?.dataset.page === String(preferPage),
    );
    if (onPage) return pointIn(onPage, pos);
  }
  let fallback: DomPoint | null = null;
  for (const host of candidates) {
    const point = pointIn(host, pos);
    if (visible(point, host)) return point;
    fallback ??= point;
  }
  return fallback;
}

/** The rendered paragraph holding the DOM selection's focus. */
export function focusParagraph(): HTMLElement | null {
  const node = window.getSelection()?.focusNode ?? null;
  const el = node instanceof HTMLElement ? node : (node?.parentElement ?? null);
  return el?.closest<HTMLElement>(".wk-p") ?? null;
}

/**
 * Copy an edited paragraph's content to its other copies (the same header on
 * other pages, the other window of a split paragraph) so they all show the
 * typing at once.
 */
export function mirrorParagraph(root: HTMLElement, edited: HTMLElement, pos: DocPosition): void {
  for (const copy of copiesOf(root, pos, `.wk-p${paragraphSelector(pos)}`)) {
    if (copy !== edited && copy.innerHTML !== edited.innerHTML) copy.innerHTML = edited.innerHTML;
  }
}

/** The 0-based page the DOM selection's focus is on, if any. */
export function focusPage(): number | undefined {
  const node = window.getSelection()?.focusNode ?? null;
  const el = node instanceof HTMLElement ? node : (node?.parentElement ?? null);
  const page = el?.closest<HTMLElement>(".wk-pagebox")?.dataset.page;
  return page === undefined ? undefined : Number(page);
}

/**
 * The paragraphs a caret can move through from `from`, in reading order: the
 * header / footer / note it is in (as shown on that page), or else the body
 * without the second window a split paragraph shows.
 */
function storyParagraphs(root: HTMLElement, from: HTMLElement): HTMLElement[] {
  const story = from.closest<HTMLElement>(STORY_SELECTOR);
  if (story) return [...story.querySelectorAll<HTMLElement>(".wk-p")];
  return [...root.querySelectorAll<HTMLElement>(".wk-p")].filter(
    (p) => !p.closest("[data-wk-clone]") && !p.closest(STORY_SELECTOR),
  );
}

/** The first or last caret point inside a paragraph's own text. */
export function edgePoint(p: HTMLElement, end: boolean): DomPoint {
  const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT, {
    // List labels and other generated text are not editable.
    acceptNode: (n) =>
      n.parentElement?.closest('[contenteditable="false"]')
        ? NodeFilter.FILTER_REJECT
        : NodeFilter.FILTER_ACCEPT,
  });
  let first: Text | null = null;
  let last: Text | null = null;
  for (let n = walker.nextNode(); n; n = walker.nextNode()) {
    if (!(n instanceof Text)) continue;
    first ??= n;
    last = n;
  }
  const node = end ? last : first;
  if (!node) return { node: p, offset: end ? p.childNodes.length : 0 };
  return { node, offset: end ? node.length : 0 };
}

const STORY_SELECTOR = ".wk-hf, .wk-note, .wk-endnotes";

export type ParagraphJump = "previous" | "next" | "first" | "last";

/** Select All (⌘A): the whole story `from` belongs to, first paragraph to last. */
export function storyRange(
  root: HTMLElement,
  from: HTMLElement,
): { start: DomPoint; end: DomPoint } | null {
  const paragraphs = storyParagraphs(root, from);
  const first = paragraphs[0];
  const last = paragraphs.at(-1);
  return first && last ? { start: edgePoint(first, false), end: edgePoint(last, true) } : null;
}

/**
 * Word's paragraph and document jumps (⌘↑ / ⌘↓, ⌘Home / ⌘End): the caret
 * point to move to from the paragraph `from` holding `at`. ⌘↑ inside a
 * paragraph goes to its own start first, as in Word.
 */
export function paragraphJumpPoint(
  root: HTMLElement,
  from: HTMLElement,
  at: DomPoint,
  jump: ParagraphJump,
): DomPoint | null {
  const paragraphs = storyParagraphs(root, from);
  const index = paragraphs.indexOf(from);
  if (jump === "first") return paragraphs[0] ? edgePoint(paragraphs[0], false) : null;
  if (jump === "last") {
    const last = paragraphs.at(-1);
    return last ? edgePoint(last, true) : null;
  }
  if (jump === "previous") {
    const start = edgePoint(from, false);
    const atStart = start.node === at.node && start.offset === at.offset;
    const target = atStart ? paragraphs[index - 1] : from;
    return target ? edgePoint(target, false) : null;
  }
  const next = paragraphs[index + 1];
  return next ? edgePoint(next, false) : edgePoint(from, true);
}
