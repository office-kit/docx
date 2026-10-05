/** Arrange commands for VML shapes, text boxes, WordArt and ink (`<w:pict>`). */

import {
  getShapeLayout,
  getShapeName,
  getShapeWrap,
  runShape,
  shapeKind,
  type HorizontalAlign,
  type ShapeLayout,
  type VerticalAlign,
} from "@office-kit/docx";
import { commands, type DocPosition } from "@office-kit/docx-editor";
import type { EditorSession, SelectedObject } from "../session.svelte";
import { selectedShape } from "../shapes/selection";
import { shapeTools } from "../shapes/tools.svelte";
import type { AlignOp, ArrangeOps, PaneObject, PositionPreset } from "./arrange";

const QUARTER_TURN = 90;
const FULL_TURN = 360;

/** Every selected shape: the Shift+click / lasso set, or the one selected shape. */
function selection(s: EditorSession): DocPosition[] {
  const one = s.selectedObject?.at;
  return shapeTools.multi.length ? [...shapeTools.multi] : one ? [one] : [];
}

const POSITIONS: Readonly<
  Record<Exclude<PositionPreset, "inline">, [HorizontalAlign, VerticalAlign]>
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

function layoutsOf(
  s: EditorSession,
  ats: readonly DocPosition[],
): Array<{ at: DocPosition; layout: ShapeLayout }> {
  const model = s.model;
  if (!model) return [];
  return ats.map((at) => ({ at, layout: getShapeLayout(commands.shapeAt(model.doc, at)) }));
}

/**
 * Align with one shape selected aligns it on the page or margin (VML's
 * `mso-position-*` alignment); with several, Word aligns them to each other's
 * bounding box and can distribute them evenly.
 */
function align(s: EditorSession, op: AlignOp, relativeTo: "page" | "margin"): void {
  const items = layoutsOf(s, selection(s));
  if (items.length === 1) {
    const [only] = items;
    if (!only) return;
    const horizontal: Partial<Record<AlignOp, HorizontalAlign>> = {
      left: "left",
      center: "center",
      right: "right",
    };
    const vertical: Partial<Record<AlignOp, VerticalAlign>> = {
      top: "top",
      middle: "center",
      bottom: "bottom",
    };
    const h = horizontal[op];
    const v = vertical[op];
    if (h)
      s.apply(commands.shapeLayoutCommand, {
        at: only.at,
        layout: { horizontalAlign: h, horizontalRelativeTo: relativeTo },
      });
    if (v)
      s.apply(commands.shapeLayoutCommand, {
        at: only.at,
        layout: { verticalAlign: v, verticalRelativeTo: relativeTo },
      });
    return;
  }
  const left = Math.min(...items.map((i) => i.layout.left));
  const top = Math.min(...items.map((i) => i.layout.top));
  const right = Math.max(...items.map((i) => i.layout.left + i.layout.width));
  const bottom = Math.max(...items.map((i) => i.layout.top + i.layout.height));
  if (op === "distributeH" || op === "distributeV") {
    const horizontal = op === "distributeH";
    const sorted = [...items].sort((a, b) =>
      horizontal ? a.layout.left - b.layout.left : a.layout.top - b.layout.top,
    );
    const total = sorted.reduce(
      (sum, i) => sum + (horizontal ? i.layout.width : i.layout.height),
      0,
    );
    const gap =
      ((horizontal ? right - left : bottom - top) - total) / Math.max(1, sorted.length - 1);
    let cursor = horizontal ? left : top;
    for (const { at, layout } of sorted) {
      s.apply(commands.shapeLayoutCommand, {
        at,
        layout: horizontal ? { left: cursor } : { top: cursor },
      });
      cursor += (horizontal ? layout.width : layout.height) + gap;
    }
    return;
  }
  for (const { at, layout } of items) {
    const next: Record<Exclude<AlignOp, "distributeH" | "distributeV">, Partial<ShapeLayout>> = {
      left: { left },
      center: { left: (left + right - layout.width) / 2 },
      right: { left: right - layout.width },
      top: { top },
      middle: { top: (top + bottom - layout.height) / 2 },
      bottom: { top: bottom - layout.height },
    };
    s.apply(commands.shapeLayoutCommand, { at, layout: next[op] });
  }
}

export const shapeArrange: ArrangeOps = {
  wrap: (s) => {
    const sel = selectedShape(s);
    return sel && getShapeWrap(sel.shape);
  },
  setWrap: (s, wrap) => {
    const sel = selectedShape(s);
    if (sel) s.apply(commands.shapeWrapCommand, { at: sel.at, wrap });
  },
  moveWithText: (s) => {
    const sel = selectedShape(s);
    return sel && getShapeLayout(sel.shape).verticalRelativeTo === "paragraph";
  },
  setMoveWithText: (s, on) => {
    const sel = selectedShape(s);
    if (sel) s.apply(commands.shapeMoveWithTextCommand, { at: sel.at, on });
  },
  position: (s, preset) => {
    const sel = selectedShape(s);
    if (!sel) return;
    if (preset === "inline") {
      s.apply(commands.shapeWrapCommand, { at: sel.at, wrap: "inline" });
      return;
    }
    const [h, v] = POSITIONS[preset];
    if (getShapeWrap(sel.shape) === "inline")
      s.apply(commands.shapeWrapCommand, { at: sel.at, wrap: "square" });
    s.apply(commands.shapeLayoutCommand, {
      at: sel.at,
      layout: {
        horizontalAlign: h,
        verticalAlign: v,
        horizontalRelativeTo: "margin",
        verticalRelativeTo: "margin",
      },
    });
  },
  order: (s, order) => {
    const sel = selectedShape(s);
    if (sel) s.apply(commands.shapeOrderCommand, { at: sel.at, order });
  },
  align,
  rotate: (s, op) => {
    for (const { at, layout } of layoutsOf(s, selection(s))) {
      const turn = op === "right90" ? QUARTER_TURN : op === "left90" ? -QUARTER_TURN : 0;
      const patch: Partial<ShapeLayout> =
        op === "flipH"
          ? { flipH: !layout.flipH }
          : op === "flipV"
            ? { flipV: !layout.flipV }
            : { rotation: (((layout.rotation + turn) % FULL_TURN) + FULL_TURN) % FULL_TURN };
      s.apply(commands.shapeLayoutCommand, { at, layout: patch });
    }
  },
  group: (s) => {
    const ats = selection(s);
    if (ats.length < 2) return;
    const at = s.apply(commands.groupShapesCommand, { ats });
    shapeTools.multi = [];
    s.selectedObject = at ? { kind: "shape", at } : null;
  },
  ungroup: (s) => {
    const sel = selectedShape(s);
    if (!sel || shapeKind(sel.shape) !== "group") return;
    s.apply(commands.ungroupShapesCommand, { at: sel.at });
    s.selectedObject = null;
  },
  layoutOptions: (s) => {
    s.pane.right = "formatShape";
  },
};

const KIND_OF: Readonly<Record<string, SelectedObject["kind"]>> = {
  textBox: "textBox",
  ink: "ink",
};

/** Selection Pane rows for the VML shapes in the body, in document order. */
export function shapeObjects(s: EditorSession): PaneObject[] {
  const model = s.model;
  if (!model || s.tick < 0) return [];
  const rows: PaneObject[] = [];
  model.doc.document.body.blocks.forEach((block, b) => {
    if (block.kind !== "paragraph") return;
    let inline = 0;
    for (const child of block.children) {
      if (child.kind !== "run") continue;
      const shape = runShape(child);
      const at: DocPosition = { block: b, inline, offset: 0 };
      inline++;
      if (!shape) continue;
      const kind = KIND_OF[shapeKind(shape)] ?? "shape";
      rows.push({
        key: `shape:${b}:${at.inline}`,
        name: getShapeName(shape),
        kind,
        hidden: getShapeLayout(shape).hidden,
        select: () => {
          s.selectedObject = { kind, at };
          shapeTools.multi = [];
          s.tick++;
        },
        setHidden: (hidden) => s.apply(commands.shapeLayoutCommand, { at, layout: { hidden } }),
        rename: (name) => s.apply(commands.shapeNameCommand, { at, name }),
      });
    }
  });
  return rows;
}
