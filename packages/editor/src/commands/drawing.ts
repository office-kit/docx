/**
 * Arrange commands for pictures and charts (Wrap Text, Position, Bring
 * Forward / Send Backward, Align, Selection Pane) and Insert ▸ Chart / Edit
 * Data. Drawings are addressed by `index` in `imageDrawings` order.
 */

import {
  addChartRun,
  type AnchorOptions,
  arrangeDrawing,
  type ChartSpec,
  type DrawingOrder,
  type HorizontalPosition,
  imageDrawings,
  removeDrawing,
  setChart,
  setDrawingAnchorOptions,
  setDrawingAspectLock,
  setDrawingHidden,
  setDrawingHyperlink,
  setDrawingName,
  setDrawingPosition,
  setDrawingWrap,
  type VerticalPosition,
  type WrapStyle,
} from "@office-kit/docx";
import type { EditorModel } from "../model.js";
import { insertRunAtCaret } from "./image.js";
import type { Command } from "./types.js";

const hasDrawings = (model: EditorModel): boolean => imageDrawings(model.doc).length > 0;

interface Placement {
  readonly horizontal?: HorizontalPosition;
  readonly vertical?: VerticalPosition;
}

/** Wrap Text. `position` places an object leaving the text line where it was shown. */
export const drawingWrapCommand: Command<{ index: number; wrap: WrapStyle; position?: Placement }> =
  {
    id: "drawing.wrap",
    group: "image",
    label: "Wrap Text",
    run(model, { index, wrap, position }) {
      setDrawingWrap(model.doc, index, wrap, position ?? {});
    },
    isEnabled: hasDrawings,
  };

/**
 * Position: a preset (In Line with Text or one of the nine square-wrapped
 * positions), Align, or a drag on the page. Floating objects only, except
 * that `wrap` lets a preset float an inline picture in the same step.
 */
export const drawingPositionCommand: Command<{
  index: number;
  position: Placement;
  wrap?: WrapStyle;
}> = {
  id: "drawing.position",
  group: "image",
  label: "Position",
  run(model, { index, position, wrap }) {
    if (wrap) setDrawingWrap(model.doc, index, wrap, position);
    setDrawingPosition(model.doc, index, position);
  },
  isEnabled: hasDrawings,
};

/** Layout Options: distance from text, wrap side and points, overlap, lock anchor, layout in cell. */
export const drawingAnchorOptionsCommand: Command<{ index: number; options: AnchorOptions }> = {
  id: "drawing.anchorOptions",
  group: "image",
  label: "Layout Options",
  run(model, { index, options }) {
    setDrawingAnchorOptions(model.doc, index, options);
  },
  isEnabled: hasDrawings,
};

export const drawingOrderCommand: Command<{ index: number; op: DrawingOrder }> = {
  id: "drawing.order",
  group: "image",
  label: "Arrange",
  run(model, { index, op }) {
    arrangeDrawing(model.doc, index, op);
  },
  isEnabled: hasDrawings,
};

/** Selection Pane rename. */
export const drawingNameCommand: Command<{ index: number; name: string }> = {
  id: "drawing.name",
  group: "image",
  label: "Rename",
  run(model, { index, name }) {
    setDrawingName(model.doc, index, name);
  },
  isEnabled: hasDrawings,
};

/** Selection Pane show / hide. */
export const drawingHiddenCommand: Command<{ index: number; hidden: boolean }> = {
  id: "drawing.hidden",
  group: "image",
  label: "Show / Hide",
  run(model, { index, hidden }) {
    setDrawingHidden(model.doc, index, hidden);
  },
  isEnabled: hasDrawings,
};

export const drawingHyperlinkCommand: Command<{ index: number; url: string | undefined }> = {
  id: "drawing.hyperlink",
  group: "image",
  label: "Link",
  run(model, { index, url }) {
    setDrawingHyperlink(model.doc, index, url);
  },
  isEnabled: hasDrawings,
};

export const drawingAspectLockCommand: Command<{ index: number; locked: boolean }> = {
  id: "drawing.aspectLock",
  group: "image",
  label: "Lock aspect ratio",
  run(model, { index, locked }) {
    setDrawingAspectLock(model.doc, index, locked);
  },
  isEnabled: hasDrawings,
};

export const drawingDeleteCommand: Command<{ index: number }> = {
  id: "drawing.delete",
  group: "image",
  label: "Delete",
  run(model, { index }) {
    removeDrawing(model.doc, index);
  },
  isEnabled: hasDrawings,
};

// Word inserts a new chart at 6" × 3.5".
const NEW_CHART_SIZE = { widthEmu: 5486400, heightEmu: 3200400 } as const;

export const insertChartCommand: Command<{ spec: ChartSpec }> = {
  id: "chart.insert",
  group: "chart",
  label: "Chart",
  run(model, { spec }) {
    insertRunAtCaret(model, addChartRun(model.doc, spec, NEW_CHART_SIZE));
  },
};

/** Edit Data, Chart Elements, Change Colors, Change Chart Type, Switch Row/Column. */
export const editChartCommand: Command<{ index: number; spec: ChartSpec }> = {
  id: "chart.edit",
  group: "chart",
  label: "Edit Data",
  run(model, { index, spec }) {
    setChart(model.doc, index, spec);
  },
  isEnabled: hasDrawings,
};

export const drawingCommands: ReadonlyArray<Command<never, unknown>> = [
  drawingWrapCommand,
  drawingPositionCommand,
  drawingAnchorOptionsCommand,
  drawingOrderCommand,
  drawingNameCommand,
  drawingHiddenCommand,
  drawingHyperlinkCommand,
  drawingAspectLockCommand,
  drawingDeleteCommand,
  insertChartCommand,
  editChartCommand,
] as ReadonlyArray<Command<never, unknown>>;
