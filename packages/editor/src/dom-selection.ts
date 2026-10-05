/**
 * Map a browser DOM selection to the editor's document positions, using the
 * `data-wk-*` anchors the renderer emits. Browser-only; guarded so importing
 * the module in Node (tests) doesn't touch `window`.
 */

import { parseStoryKey } from "@office-kit/docx";
import type { CellCoord, DocPosition, Selection } from "./selection.js";

/**
 * The attribute a story's container carries (`header:/word/header1.xml`,
 * `footnote:2` …); positions inside it address that story.
 */
export const STORY_ATTR = "data-wk-story";

export function parseCell(value: string | null): CellCoord | undefined {
  if (!value) return undefined;
  const [r, c] = value.split(",").map((n) => Number.parseInt(n, 10));
  if (r === undefined || c === undefined || Number.isNaN(r) || Number.isNaN(c)) return undefined;
  return { row: r, col: c };
}

/** Walk up from a DOM node to the nearest element carrying `data-wk-block`. */
function anchorElement(node: Node | null): HTMLElement | null {
  let el: Node | null = node;
  while (el && !(el instanceof HTMLElement && el.hasAttribute("data-wk-block"))) {
    el = el.parentNode;
  }
  return el as HTMLElement | null;
}

/** Visible character count of a rendered run span (the empty-run placeholder is not content). */
function runSpanLength(span: Element): number {
  return (span.textContent ?? "").replace(/\u200b/g, "").length;
}

function textBefore(span: Element, node: Node, offset: number): number {
  const range = span.ownerDocument.createRange();
  range.setStart(span, 0);
  range.setEnd(node, offset);
  return range.toString().replace(/\u200b/g, "").length;
}

/**
 * The DOM point `offset` characters into a run span. The span holds text
 * nodes and tab elements; a point next to a tab is placed in the text beside
 * it, or between the span's children, never inside the (read-only) tab.
 */
export function runPoint(span: Element, offset: number): { node: Node; offset: number } {
  let remaining = offset;
  const nodes = span.childNodes;
  for (let i = 0; i < nodes.length; i++) {
    const node = nodes[i];
    if (!node) continue;
    const length = (node.textContent ?? "").replace(/\u200b/g, "").length;
    if (node.nodeType === Node.TEXT_NODE) {
      if (remaining <= length) return { node, offset: remaining };
    } else if (remaining === 0) {
      return { node: span, offset: i };
    }
    remaining -= length;
  }
  return { node: span, offset: nodes.length };
}

/**
 * Resolve a DOM point that is *not* inside a run span — on the paragraph
 * element itself, or inside a read-only inline (hyperlink / field text) — to
 * the end of the nearest run span before it. Without this the point's DOM
 * offset (a child index, or an offset in link text) would be misread as a
 * character offset into run 0.
 */
function inlineAndOffsetInParagraph(
  para: HTMLElement,
  node: Node,
  offset: number,
): { inline?: number; offset: number } {
  // The paragraph child the point sits before.
  let childIndex: number;
  if (node === para) {
    childIndex = offset;
  } else {
    let child: Node | null = node;
    while (child && child.parentNode !== para) child = child.parentNode;
    const index = child ? Array.prototype.indexOf.call(para.childNodes, child) : 0;
    // Past the start of that child means "after it".
    childIndex = offset > 0 ? index + 1 : index;
  }
  for (let i = Math.min(childIndex, para.childNodes.length) - 1; i >= 0; i--) {
    const prev = para.childNodes[i];
    if (prev instanceof HTMLElement && prev.classList.contains("wk-run")) {
      const inline = Number.parseInt(prev.getAttribute("data-wk-inline") ?? "", 10);
      if (!Number.isNaN(inline)) return { inline, offset: runSpanLength(prev) };
    }
  }
  return { inline: 0, offset: 0 };
}

/** Build a {@link DocPosition} from a DOM container node and character offset. */
export function positionFromDom(node: Node | null, offset: number): DocPosition | null {
  const anchor = anchorElement(node);
  if (!anchor || !node) return null;
  const block = Number.parseInt(anchor.getAttribute("data-wk-block") ?? "", 10);
  if (Number.isNaN(block)) return null;
  const cell = parseCell(anchor.getAttribute("data-wk-cell"));
  const paraAttr = anchor.getAttribute("data-wk-para");
  const para = paraAttr ? Number.parseInt(paraAttr, 10) : undefined;
  const inlineAttr = anchor.getAttribute("data-wk-inline");
  let inline = inlineAttr ? Number.parseInt(inlineAttr, 10) : undefined;
  if (inline !== undefined) {
    // A run's text spans several nodes once it holds a tab; count characters
    // from the run's start (a point on an element counts child nodes).
    offset = textBefore(anchor, node, offset);
  } else if (inline === undefined && anchor.classList.contains("wk-p")) {
    ({ inline, offset } = inlineAndOffsetInParagraph(anchor, node, offset));
  }
  const storyKey = anchor.closest(`[${STORY_ATTR}]`)?.getAttribute(STORY_ATTR);
  const story = storyKey ? parseStoryKey(storyKey) : undefined;
  return {
    ...(story ? { story } : {}),
    block,
    ...(cell ? { cell } : {}),
    ...(cell && para !== undefined && !Number.isNaN(para) ? { para } : {}),
    ...(inline !== undefined && !Number.isNaN(inline) ? { inline } : {}),
    offset,
  };
}

/** Read the current window selection as an editor {@link Selection}. */
export function readDomSelection(root: Document | ShadowRoot = document): Selection | null {
  const domSel =
    "getSelection" in root ? (root as Document).getSelection?.() : window.getSelection();
  if (!domSel || domSel.rangeCount === 0) return null;
  const range = domSel.getRangeAt(0);
  const anchor = positionFromDom(range.startContainer, range.startOffset);
  const focus = positionFromDom(range.endContainer, range.endOffset);
  if (!anchor || !focus) return null;
  return { anchor, focus };
}
