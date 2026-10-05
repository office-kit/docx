/**
 * Lay tabs out against their tab stops (§17.3.1.37 custom stops, §17.15.1.25
 * the default interval), the way Word sets a line: a tab reaches the first
 * custom stop past it, or past every custom stop the next multiple of the
 * default interval. A right / center / decimal stop aligns the text after the
 * tab (up to the next tab) against the stop instead of starting it there.
 *
 * The renderer gives each tab its own `.wk-tab` element and the paragraph a
 * `data-wk-tabs` spec; this sizes the elements once they are in the DOM at
 * their final width. Browser-only. Right-to-left paragraphs are measured from
 * the left edge too, where Word measures them from the right.
 */

import type { TabStopSpec } from "./render-format.js";

const TWIPS_PER_PX = 15;
// Measurement noise below this, in twips, does not move a tab to the next stop.
const EPSILON_TWIPS = 1;

/** Size every tab inside `root` (paragraphs included). */
export function layoutTabStops(root: ParentNode): void {
  if (root instanceof HTMLElement && root.hasAttribute("data-wk-tabs")) layoutParagraph(root);
  for (const para of Array.from(root.querySelectorAll<HTMLElement>("[data-wk-tabs]"))) {
    layoutParagraph(para);
  }
}

function layoutParagraph(para: HTMLElement): void {
  const spec = JSON.parse(para.dataset.wkTabs ?? "") as TabStopSpec;
  const tabs = Array.from(para.querySelectorAll<HTMLElement>(".wk-tab")).filter(
    (tab) => tab.closest("[data-wk-tabs]") === para,
  );
  for (const tab of tabs) {
    tab.style.width = "0px";
    delete tab.dataset.leader;
  }
  const rect = para.getBoundingClientRect();
  // Client pixels per CSS pixel, so a zoomed canvas measures in document units.
  const scale = para.offsetWidth > 0 ? rect.width / para.offsetWidth : 1;
  // Stops count from the text column's edge; the left indent is the margin.
  const origin = rect.left - Number.parseFloat(getComputedStyle(para).marginLeft) * scale;
  const toTwips = (clientX: number) => ((clientX - origin) / scale) * TWIPS_PER_PX;

  tabs.forEach((tab, i) => {
    const at = toTwips(tab.getBoundingClientRect().left);
    const stop = nextStop(spec, at);
    const after = textAfter(tab, tabs[i + 1], para, stop.align === "decimal");
    const segment = after ? (lineWidth(after) / scale) * TWIPS_PER_PX : 0;
    const shift =
      stop.align === "right" || stop.align === "decimal"
        ? segment
        : stop.align === "center"
          ? segment / 2
          : 0;
    // Text too wide to align at the stop starts at the tab, as in Word.
    const width = Math.max(stop.position - at - shift, 0);
    tab.style.width = `${width / TWIPS_PER_PX}px`;
    if (stop.leader !== "none") tab.dataset.leader = stop.leader;
  });
}

function nextStop(spec: TabStopSpec, at: number): TabStopSpec["stops"][number] {
  const past = at + EPSILON_TWIPS;
  const custom = spec.stops.find((s) => s.position > past);
  const hanging =
    spec.hangingIndent !== undefined && spec.hangingIndent > past ? spec.hangingIndent : undefined;
  if (custom && (hanging === undefined || custom.position <= hanging)) return custom;
  if (hanging !== undefined) return { position: hanging, align: "left", leader: "none" };
  const interval = spec.interval > 0 ? spec.interval : 1;
  return { position: (Math.floor(past / interval) + 1) * interval, align: "left", leader: "none" };
}

/**
 * The content a stop aligns: from the tab to the next tab of the paragraph,
 * or for a decimal stop up to the first `.` in it.
 */
function textAfter(
  tab: HTMLElement,
  next: HTMLElement | undefined,
  para: HTMLElement,
  decimal: boolean,
): Range | undefined {
  const range = document.createRange();
  range.setStartAfter(tab);
  if (next) range.setEndBefore(next);
  else range.setEnd(para, para.childNodes.length);
  if (decimal) {
    const walker = document.createTreeWalker(range.commonAncestorContainer, NodeFilter.SHOW_TEXT);
    for (let node = walker.nextNode(); node; node = walker.nextNode()) {
      if (!range.intersectsNode(node)) continue;
      const point = (node.textContent ?? "").indexOf(".");
      if (point >= 0 && range.comparePoint(node, point) === 0) {
        range.setEnd(node, point);
        break;
      }
    }
  }
  return range.collapsed ? undefined : range;
}

/** Width of a range's first line: its rects until one wraps back to the left. */
function lineWidth(range: Range): number {
  let left = Number.POSITIVE_INFINITY;
  let right = Number.NEGATIVE_INFINITY;
  for (const rect of Array.from(range.getClientRects())) {
    if (rect.width === 0) continue;
    if (rect.left < right - 1) break;
    left = Math.min(left, rect.left);
    right = Math.max(right, rect.right);
  }
  return right > left ? right - left : 0;
}
