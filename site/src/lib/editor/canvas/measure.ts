/**
 * DOM measurement for the paginator. Everything is returned in unzoomed CSS
 * px: the canvas may sit under CSS `zoom`, which scales client rects but not
 * layout sizes, so rect-based numbers are divided by the zoom actually seen.
 */

import type { WmlParagraph } from "@office-kit/docx";

/** Rect units per layout px (the effective zoom) for an element. */
export function zoomOf(el: HTMLElement): number {
  const layout = el.offsetWidth;
  const rect = el.getBoundingClientRect().width;
  return layout > 0 && rect > 0 ? rect / layout : 1;
}

/** Height including vertical margins. */
export function outerHeight(el: HTMLElement): number {
  const style = getComputedStyle(el);
  return (
    el.offsetHeight + (parseFloat(style.marginTop) || 0) + (parseFloat(style.marginBottom) || 0)
  );
}

function marginTop(el: HTMLElement): number {
  return parseFloat(getComputedStyle(el).marginTop) || 0;
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

// Rects whose vertical ranges overlap by more than this are one line.
const SAME_LINE_SLACK = 1;

/**
 * Bottoms of the element's line boxes, measured from the top of its margin
 * box, in order. The last entry is the element's full outer height so the
 * paragraph's space after travels with its last line.
 */
export function lineBottoms(content: HTMLElement): number[] {
  const zoom = zoomOf(content);
  const rects = textRects(content);
  if (rects.length === 0) return [outerHeight(content)];
  rects.sort((a, b) => a.top - b.top);
  const bottoms: number[] = [];
  let lineBottom = Number.NEGATIVE_INFINITY;
  for (const r of rects) {
    if (r.top >= lineBottom - SAME_LINE_SLACK) {
      if (Number.isFinite(lineBottom)) bottoms.push(lineBottom);
      lineBottom = r.bottom;
    } else {
      lineBottom = Math.max(lineBottom, r.bottom);
    }
  }
  bottoms.push(lineBottom);
  const origin = content.getBoundingClientRect().top - marginTop(content) * zoom;
  const out = bottoms.map((b) => (b - origin) / zoom);
  out[out.length - 1] = outerHeight(content);
  return out;
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
  const zoom = zoomOf(content);
  const origin = content.getBoundingClientRect().top - marginTop(content) * zoom;
  const out: Array<{ offset: number; kind: "page" | "column" }> = [];
  for (const br of breaks) {
    const text = content.querySelector(`.wk-run[data-wk-inline="${br.inline}"]`)?.firstChild;
    if (!text || text.nodeType !== Node.TEXT_NODE) continue;
    const range = document.createRange();
    const at = Math.min(br.char, (text.textContent ?? "").length - 1);
    if (at < 0) continue;
    range.setStart(text, at);
    range.setEnd(text, at + 1);
    const rect = range.getBoundingClientRect();
    out.push({ offset: (rect.bottom - origin) / zoom, kind: br.kind });
  }
  return out;
}
