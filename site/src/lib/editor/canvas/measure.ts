/**
 * DOM measurement for the paginator. Everything is returned in unzoomed CSS
 * px: the canvas may sit under CSS `zoom`, which scales client rects but not
 * layout sizes, so rect-based numbers are divided by the zoom actually seen.
 *
 * "Height" is along the block axis, the way lines follow each other: down the
 * page, or right to left in vertical text (縦書き, `writing-mode: vertical-rl`).
 */

import type { WmlParagraph } from "@office-kit/docx";
import { runPoint } from "@office-kit/docx-editor";

/** Rect units per layout px (the effective zoom) for an element. */
export function zoomOf(el: HTMLElement): number {
  const layout = el.offsetWidth;
  const rect = el.getBoundingClientRect().width;
  return layout > 0 && rect > 0 ? rect / layout : 1;
}

function isVertical(el: Element): boolean {
  return getComputedStyle(el).writingMode.startsWith("vertical");
}

/** Layout size along the block axis. */
export function blockSize(el: HTMLElement): number {
  return isVertical(el) ? el.offsetWidth : el.offsetHeight;
}

/** Block size including the block-axis margins. */
export function outerHeight(el: HTMLElement): number {
  const style = getComputedStyle(el);
  return (
    blockSize(el) +
    (parseFloat(style.marginBlockStart) || 0) +
    (parseFloat(style.marginBlockEnd) || 0)
  );
}

/**
 * Where a client rect starts and ends on the block axis, from the start of
 * the element's margin box, in unzoomed px. Vertical text is `vertical-rl`
 * (Word's tbRl): its block axis runs right to left.
 */
function blockAxis(content: HTMLElement): {
  start: (r: DOMRect) => number;
  end: (r: DOMRect) => number;
} {
  const zoom = zoomOf(content);
  const margin = (parseFloat(getComputedStyle(content).marginBlockStart) || 0) * zoom;
  const box = content.getBoundingClientRect();
  if (isVertical(content)) {
    const origin = box.right + margin;
    return { start: (r) => (origin - r.right) / zoom, end: (r) => (origin - r.left) / zoom };
  }
  const origin = box.top - margin;
  return { start: (r) => (r.top - origin) / zoom, end: (r) => (r.bottom - origin) / zoom };
}

/**
 * Client rects of the text inside an element, one per line fragment. Element
 * boxes are left out: a block's own rect spans all of its lines.
 */
export function textRects(el: HTMLElement): DOMRect[] {
  const out: DOMRect[] = [];
  const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
  const range = document.createRange();
  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    range.selectNodeContents(node);
    for (const rect of range.getClientRects()) if (rect.height > 0) out.push(rect);
  }
  return out;
}

// Rects whose block-axis ranges overlap by more than this are one line.
const SAME_LINE_SLACK = 1;

/**
 * Bottoms of the element's line boxes, measured from the top of its margin
 * box, in order. The last entry is the element's full outer height so the
 * paragraph's space after travels with its last line.
 */
export function lineBottoms(content: HTMLElement): number[] {
  const axis = blockAxis(content);
  const spans = textRects(content).map((r) => ({ start: axis.start(r), end: axis.end(r) }));
  if (spans.length === 0) return [outerHeight(content)];
  spans.sort((a, b) => a.start - b.start);
  const bottoms: number[] = [];
  let lineBottom = Number.NEGATIVE_INFINITY;
  for (const r of spans) {
    if (r.start >= lineBottom - SAME_LINE_SLACK) {
      if (Number.isFinite(lineBottom)) bottoms.push(lineBottom);
      lineBottom = r.end;
    } else {
      lineBottom = Math.max(lineBottom, r.end);
    }
  }
  bottoms.push(outerHeight(content));
  return bottoms;
}

/** A `w:br w:type="page|column"` inside a paragraph, by run and character. */
export interface ForcedBreak {
  readonly inline: number;
  readonly char: number;
  readonly kind: "page" | "column";
}

/** Run-relative positions of the page / column breaks in a paragraph. */
export function forcedBreaks(para: WmlParagraph): ForcedBreak[] {
  const out: ForcedBreak[] = [];
  let inline = 0;
  for (const child of para.children) {
    if (child.kind !== "run") continue;
    let char = 0;
    for (const piece of child.pieces) {
      if (piece.kind === "break" && (piece.breakType === "page" || piece.breakType === "column")) {
        out.push({ inline, char, kind: piece.breakType });
      }
      // The renderer writes text, tabs, breaks and hyphens as one char each.
      if (piece.kind === "text") char += piece.value.length;
      else if (
        piece.kind === "tab" ||
        piece.kind === "break" ||
        piece.kind === "noBreakHyphen" ||
        piece.kind === "softHyphen"
      )
        char += 1;
    }
    inline++;
  }
  return out;
}

/** Offsets (from the margin-box top) where content after each forced break starts. */
export function breakOffsets(
  content: HTMLElement,
  breaks: readonly ForcedBreak[],
): Array<{ offset: number; kind: "page" | "column" }> {
  if (breaks.length === 0) return [];
  const axis = blockAxis(content);
  const out: Array<{ offset: number; kind: "page" | "column" }> = [];
  for (const br of breaks) {
    const span = content.querySelector(`.wk-run[data-wk-inline="${br.inline}"]`);
    if (!span) continue;
    const { node, offset } = runPoint(span, br.char);
    if (node.nodeType !== Node.TEXT_NODE || offset >= (node.textContent ?? "").length) continue;
    const range = document.createRange();
    range.setStart(node, offset);
    range.setEnd(node, offset + 1);
    const rect = range.getBoundingClientRect();
    out.push({ offset: axis.end(rect), kind: br.kind });
  }
  return out;
}
