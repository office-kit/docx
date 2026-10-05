/** Arrange commands for DrawingML pictures and charts (`<w:drawing>`). */

import { commands, drawingPositionOf, EMU_PER_PX } from "@office-kit/docx-editor";
import {
  type HorizontalAlign,
  imageDrawings,
  readDrawing,
  type VerticalAlign,
} from "@office-kit/docx";
import type { AlignOp, ArrangeOps, PaneObject, PositionPreset } from "./arrange";
import type { EditorSession } from "../session.svelte";
import { applyToSelected, drawingElement, pageBoxOf, selected } from "../picture/state";
import { pictureModes } from "../picture/ui.svelte";

const QUARTER_TURN = 90;
const FULL_TURN = 360;

const PRESET_ALIGN: Readonly<
  Record<Exclude<PositionPreset, "inline">, readonly [HorizontalAlign, VerticalAlign]>
> = {
  topLeft: ["left", "top"],
  topCenter: ["center", "top"],
  topRight: ["right", "top"],
  middleLeft: ["left", "center"],
  middleCenter: ["center", "center"],
  middleRight: ["right", "center"],
  bottomLeft: ["left", "bottom"],
  bottomCenter: ["center", "bottom"],
  bottomRight: ["right", "bottom"],
};

/** Where the selected object is drawn, in the canvas' unzoomed coordinates, and the page band it is on. */
function shownAt(s: EditorSession) {
  const sel = selected(s);
  const el = sel ? drawingElement(sel.index) : null;
  const canvas = el?.closest<HTMLElement>(".wk-canvas");
  const model = s.model;
  if (!el || !canvas || !model) return undefined;
  const scale = canvas.getBoundingClientRect().width / canvas.offsetWidth || 1;
  const c = canvas.getBoundingClientRect();
  const local = (r: DOMRect) => ({
    x: (r.left - c.left) / scale - canvas.clientLeft,
    y: (r.top - c.top) / scale - canvas.clientTop,
  });
  const at = local(el.getBoundingClientRect());
  const para = local((el.closest(".wk-p") ?? el).getBoundingClientRect());
  const page = pageBoxOf(model.doc);
  const pageTop = Math.floor(Math.max(0, para.y) / page.height) * page.height;
  return { at, para, page, pageTop };
}

const emu = (px: number): number => Math.round(px * EMU_PER_PX);

export const pictureArrange: ArrangeOps = {
  wrap: (s) => selected(s)?.info.wrap,

  setWrap(s, wrap) {
    const sel = selected(s);
    if (!sel) return;
    if (sel.info.wrap !== "inline" || wrap === "inline") {
      applyToSelected(s, commands.drawingWrapCommand, { wrap });
      return;
    }
    // Leaving the line: keep the object where it is shown, as Word does.
    const shown = shownAt(s);
    applyToSelected(s, commands.drawingWrapCommand, {
      wrap,
      ...(shown && {
        position: {
          horizontal: { relativeTo: "column", offsetEmu: emu(shown.at.x - shown.page.left) },
          vertical: { relativeTo: "paragraph", offsetEmu: emu(shown.at.y - shown.para.y) },
        },
      }),
    });
  },

  moveWithText(s) {
    const anchor = selected(s)?.info.anchor;
    return anchor ? anchor.vertical.relativeTo === "paragraph" : undefined;
  },

  setMoveWithText(s, on) {
    const shown = shownAt(s);
    if (!shown) return;
    // Same spot on the page, measured from the paragraph (moves with text) or the page (fixed).
    const vertical = on
      ? { relativeTo: "paragraph" as const, offsetEmu: emu(shown.at.y - shown.para.y) }
      : { relativeTo: "page" as const, offsetEmu: emu(shown.at.y - shown.pageTop) };
    applyToSelected(s, commands.drawingPositionCommand, { position: { vertical } });
  },

  position(s, preset) {
    if (preset === "inline") {
      applyToSelected(s, commands.drawingWrapCommand, { wrap: "inline" });
      return;
    }
    const [h, v] = PRESET_ALIGN[preset];
    applyToSelected(s, commands.drawingPositionCommand, {
      position: {
        horizontal: { relativeTo: "margin", align: h },
        vertical: { relativeTo: "margin", align: v },
      },
      wrap: "square",
    });
  },

  order(s, op) {
    applyToSelected(s, commands.drawingOrderCommand, { op });
  },

  align(s, op: AlignOp, relativeTo) {
    const sel = selected(s);
    if (!sel) return;
    // A single object distributes to the middle of the page / margins, as in Word.
    const horizontal: Partial<Record<AlignOp, HorizontalAlign>> = {
      left: "left",
      center: "center",
      right: "right",
      distributeH: "center",
    };
    const vertical: Partial<Record<AlignOp, VerticalAlign>> = {
      top: "top",
      middle: "center",
      bottom: "bottom",
      distributeV: "center",
    };
    const h = horizontal[op];
    const v = vertical[op];
    const position = h
      ? { horizontal: { relativeTo, align: h } }
      : v
        ? { vertical: { relativeTo, align: v } }
        : undefined;
    if (!position) return;
    applyToSelected(s, commands.drawingPositionCommand, {
      position,
      ...(sel.info.wrap === "inline" && { wrap: "square" as const }),
    });
  },

  rotate(s, op) {
    const info = selected(s)?.info;
    if (!info) return;
    const turn = (d: number) => (((info.rotation + d) % FULL_TURN) + FULL_TURN) % FULL_TURN;
    switch (op) {
      case "right90":
        applyToSelected(s, commands.pictureTransformCommand, { rotation: turn(QUARTER_TURN) });
        break;
      case "left90":
        applyToSelected(s, commands.pictureTransformCommand, { rotation: turn(-QUARTER_TURN) });
        break;
      case "flipH":
        applyToSelected(s, commands.pictureTransformCommand, { flipH: !info.flipH });
        break;
      case "flipV":
        applyToSelected(s, commands.pictureTransformCommand, { flipV: !info.flipV });
        break;
    }
  },

  editWrapPoints(s) {
    const info = selected(s)?.info;
    if (!info) return;
    // Wrap points only shape Tight and Through wrapping.
    if (info.wrap !== "tight" && info.wrap !== "through") pictureArrange.setWrap(s, "tight");
    pictureModes.editingWrapPoints = true;
  },

  layoutOptions(s) {
    s.openDialog("picture.layout");
  },
};

/** Selection Pane rows for the document's pictures and charts, top of the stack first. */
export function pictureObjects(s: EditorSession): PaneObject[] {
  const model = s.model;
  if (!model || s.tick < 0) return [];
  const doc = model.doc;
  const rows = imageDrawings(doc).flatMap((drawing, index) => {
    const info = readDrawing(doc, drawing);
    if (info.kind === "other") return [];
    const kind = info.kind;
    const row: PaneObject & { z: number } = {
      key: `drawing-${index}`,
      name: info.name || `${kind === "chart" ? "Chart" : "Picture"} ${info.id}`,
      kind,
      hidden: info.hidden,
      z: info.anchor?.relativeHeight ?? 0,
      select() {
        const at = drawingPositionOf(doc, index);
        if (at) s.selectedObject = { kind, at };
      },
      setHidden: (hidden) => s.apply(commands.drawingHiddenCommand, { index, hidden }),
      rename: (name) => s.apply(commands.drawingNameCommand, { index, name }),
    };
    return [row];
  });
  return rows.toSorted((a, b) => b.z - a.z);
}
