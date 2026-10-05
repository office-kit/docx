/**
 * Place floating objects (`wp:anchor`) on the rendered canvas. Browser-only:
 * it measures the laid-out paragraphs, so it runs after the canvas HTML is in
 * the document.
 *
 * The canvas is one continuous flow, so "the page" of an object is the
 * page-height band its anchor paragraph falls in. Text wrapping maps onto CSS:
 * Square / Tight / Through become floats (Tight and Through follow the wrap
 * polygon through `shape-outside`), Top and Bottom a block, and Behind / In
 * Front of Text absolutely positioned layers.
 */

import type { FloatData } from "./render-drawing.js";

/** Page size and margins in CSS px (the section's `w:pgSz` / `w:pgMar`). */
export interface PageBox {
  readonly width: number;
  readonly height: number;
  readonly top: number;
  readonly right: number;
  readonly bottom: number;
  readonly left: number;
}

// wrapPolygon points live in a 21600-unit square over the object.
const POLYGON_UNITS = 21600;
const EMU_PER_PX = 9525;

interface Box {
  readonly x: number;
  readonly y: number;
  readonly w: number;
  readonly h: number;
}

/** `el`'s border box in the canvas' own (unzoomed) coordinates. */
function boxIn(canvas: HTMLElement, el: Element, scale: number): Box {
  const c = canvas.getBoundingClientRect();
  const r = el.getBoundingClientRect();
  return {
    x: (r.left - c.left) / scale - canvas.clientLeft,
    y: (r.top - c.top) / scale - canvas.clientTop,
    w: r.width / scale,
    h: r.height / scale,
  };
}

function horizontalFrame(rel: string, page: PageBox, para: Box): readonly [number, number] {
  switch (rel) {
    case "page":
      return [0, page.width];
    case "leftMargin":
    case "insideMargin":
      return [0, page.left];
    case "rightMargin":
    case "outsideMargin":
      return [page.width - page.right, page.width];
    case "character":
      return [para.x, para.x + para.w];
    default:
      return [page.left, page.width - page.right];
  }
}

function verticalFrame(
  rel: string,
  page: PageBox,
  pageTop: number,
  para: Box,
): readonly [number, number] {
  switch (rel) {
    case "page":
      return [pageTop, pageTop + page.height];
    case "topMargin":
    case "insideMargin":
      return [pageTop, pageTop + page.top];
    case "bottomMargin":
    case "outsideMargin":
      return [pageTop + page.height - page.bottom, pageTop + page.height];
    case "margin":
      return [pageTop + page.top, pageTop + page.height - page.bottom];
    default:
      return [para.y, para.y + para.h];
  }
}

function place(
  frame: readonly [number, number],
  size: number,
  pos: FloatData["h"] | FloatData["v"],
): number {
  const [start, end] = frame;
  if ("offsetEmu" in pos) return start + pos.offsetEmu / EMU_PER_PX;
  switch (pos.align) {
    case "center":
      return start + (end - start - size) / 2;
    case "right":
    case "bottom":
    case "outside":
      return end - size;
    default:
      return start;
  }
}

function frames(canvas: HTMLElement, el: HTMLElement, page: PageBox, data: FloatData) {
  const paraEl = el.closest(".wk-p") ?? el.parentElement ?? canvas;
  const scale = canvas.getBoundingClientRect().width / canvas.offsetWidth || 1;
  const para = boxIn(canvas, paraEl, scale);
  const pageTop = Math.floor(Math.max(0, para.y) / page.height) * page.height;
  return {
    h: horizontalFrame(data.h.relativeTo, page, para),
    v: verticalFrame(data.v.relativeTo, page, pageTop, para),
  };
}

/** Where a floating object's top-left corner goes, in canvas coordinates. */
export function floatOrigin(
  canvas: HTMLElement,
  el: HTMLElement,
  page: PageBox,
): { x: number; y: number } {
  const data = readFloat(el);
  if (!data) return { x: 0, y: 0 };
  const f = frames(canvas, el, page, data);
  return { x: place(f.h, data.w, data.h), y: place(f.v, data.hgt, data.v) };
}

/**
 * The start of the frames an object is positioned against (what a zero
 * `posOffset` means). Dragging converts a canvas point to an offset from here.
 */
export function floatFrameStart(
  canvas: HTMLElement,
  el: HTMLElement,
  page: PageBox,
): { x: number; y: number } {
  const data = readFloat(el);
  if (!data) return { x: 0, y: 0 };
  const f = frames(canvas, el, page, data);
  return { x: f.h[0], y: f.v[0] };
}

const LAYOUT_PROPS = [
  "position",
  "left",
  "top",
  "float",
  "clear",
  "margin",
  "display",
  "shape-outside",
  "shape-margin",
];

export function readFloat(el: HTMLElement): FloatData | undefined {
  const raw = el.dataset.wkFloat;
  if (!raw) return undefined;
  // Written by our own renderer; a parse failure would be a renderer bug.
  return JSON.parse(raw) as FloatData;
}

/** Lay out every floating object under `canvas` (call after each render). */
export function layoutFloatingObjects(canvas: HTMLElement, page: PageBox): void {
  // Idempotent: undo the previous pass before measuring (text may have reflowed).
  for (const spacer of Array.from(canvas.querySelectorAll(".wk-float-spacer"))) spacer.remove();
  const floats = Array.from(canvas.querySelectorAll<HTMLElement>(".wk-obj-float"));
  for (const el of floats) for (const prop of LAYOUT_PROPS) el.style.removeProperty(prop);
  for (const el of floats) {
    const data = readFloat(el);
    if (!data) continue;
    const paraEl = el.closest<HTMLElement>(".wk-p");
    const { x, y } = floatOrigin(canvas, el, page);
    if (data.wrap === "behindText" || data.wrap === "inFrontOfText" || !paraEl) {
      el.style.position = "absolute";
      el.style.left = `${x}px`;
      el.style.top = `${y}px`;
      continue;
    }
    // Wrapped objects flow with the text from the top of their paragraph.
    paraEl.prepend(el);
    const scale = canvas.getBoundingClientRect().width / canvas.offsetWidth || 1;
    const para = boxIn(canvas, paraEl, scale);
    const [distT, distR, distB, distL] = data.dist;
    const dy = y - para.y;
    if (data.wrap === "topAndBottom") {
      el.style.display = "block";
      el.style.margin = `${Math.max(0, dy) + distT}px 0 ${distB}px ${x - para.x}px`;
      continue;
    }
    const left = x + data.w / 2 <= para.x + para.w / 2;
    el.style.float = left ? "left" : "right";
    el.style.clear = left ? "left" : "right";
    if (left) el.style.margin = `${distT}px ${distR}px ${distB}px ${x - para.x}px`;
    else el.style.margin = `${distT}px ${para.x + para.w - (x + data.w)}px ${distB}px ${distL}px`;
    if (dy > 0) {
      // A zero-width float holds the object down to its offset without
      // keeping text out of the space above it.
      const spacer = document.createElement("span");
      spacer.className = "wk-float-spacer";
      spacer.style.cssText = `float:${left ? "left" : "right"};width:0;height:${dy}px`;
      el.before(spacer);
    } else if (dy < 0) {
      el.style.position = "relative";
      el.style.top = `${dy}px`;
    }
    if ((data.wrap === "tight" || data.wrap === "through") && data.polygon) {
      const ml = left ? x - para.x : distL;
      const points = data.polygon
        .map(
          ([px, py]) =>
            `${ml + (px / POLYGON_UNITS) * data.w}px ${distT + (py / POLYGON_UNITS) * data.hgt}px`,
        )
        .join(",");
      el.style.shapeOutside = `polygon(${points})`;
      el.style.shapeMargin = `${Math.max(distL, distR)}px`;
    }
  }
}
