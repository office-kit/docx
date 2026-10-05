/**
 * The selected picture or chart, as the Picture Format / Chart Design tabs and
 * the picture panes see it: its drawing index (what the commands take), the
 * drawing read into plain data, and the page box floating objects are placed
 * against.
 */

import {
  drawingIndexAt,
  EMU_PER_PX,
  type Command,
  type PageBox,
  pageGeometry,
} from "@office-kit/docx-editor";
import {
  type ChartSpec,
  type Docx,
  type DrawingInfo,
  imageDrawings,
  readChart,
  readDrawing,
} from "@office-kit/docx";
import type { EditorSession } from "../session.svelte";

export { EMU_PER_PX };

// 1 twip = 1/1440 in; 96 CSS px per inch.
const TWIPS_PER_PX = 15;
export const EMU_PER_CM = 360000;
export const EMU_PER_PT = 12700;

/** The selected drawing's index, or -1 when no picture / chart is selected. */
export function selectedIndex(session: EditorSession): number {
  const sel = session.selectedObject;
  const model = session.model;
  if (!sel || !model || (sel.kind !== "picture" && sel.kind !== "chart")) return -1;
  return drawingIndexAt(model.doc, sel.at);
}

export interface Selected {
  readonly index: number;
  readonly info: DrawingInfo;
}

/** The selected drawing read into plain data (re-read whenever `session.tick` moves). */
export function selected(session: EditorSession): Selected | undefined {
  if (session.tick < 0) return undefined;
  const index = selectedIndex(session);
  const model = session.model;
  const drawing = model ? imageDrawings(model.doc)[index] : undefined;
  if (!model || !drawing) return undefined;
  return { index, info: readDrawing(model.doc, drawing) };
}

export function selectedChart(
  session: EditorSession,
): { index: number; spec: ChartSpec } | undefined {
  const s = selected(session);
  const model = session.model;
  const drawing = model && s ? imageDrawings(model.doc)[s.index] : undefined;
  const spec = model && drawing ? readChart(model.doc, drawing) : undefined;
  return s && spec ? { index: s.index, spec } : undefined;
}

/** Run a drawing command on the selection; `params` without the index. */
export function applyToSelected<P extends { index: number }>(
  session: EditorSession,
  cmd: Command<P, unknown>,
  params: Omit<P, "index">,
): void {
  const index = selectedIndex(session);
  if (index < 0) return;
  // `index` completes P: the caller supplied every other member.
  session.apply(cmd, { ...params, index } as P);
}

export function pageBoxOf(doc: Docx): PageBox {
  const g = pageGeometry(doc);
  return {
    width: g.width / TWIPS_PER_PX,
    height: g.height / TWIPS_PER_PX,
    top: g.top / TWIPS_PER_PX,
    right: g.right / TWIPS_PER_PX,
    bottom: g.bottom / TWIPS_PER_PX,
    left: g.left / TWIPS_PER_PX,
  };
}

/** The text column width in EMU: Word scales a large inserted picture down to it. */
export function textWidthEmu(doc: Docx): number {
  const page = pageBoxOf(doc);
  return (page.width - page.left - page.right) * EMU_PER_PX;
}

/** The rendered element of drawing `index` on the canvas. */
export function drawingElement(index: number): HTMLElement | null {
  return document.querySelector<HTMLElement>(`.wk-canvas [data-wk-drawing="${index}"]`);
}

/** Scale `size` down (never up) to fit `maxWidth`, keeping its aspect. */
export function fitWidth(
  size: { widthEmu: number; heightEmu: number },
  maxWidth: number,
): { widthEmu: number; heightEmu: number } {
  if (size.widthEmu <= maxWidth) return size;
  const k = maxWidth / size.widthEmu;
  return { widthEmu: Math.round(maxWidth), heightEmu: Math.round(size.heightEmu * k) };
}
