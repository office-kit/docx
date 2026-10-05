/**
 * The Arrange group (Position, Wrap Text, Bring Forward, Send Backward,
 * Selection Pane, Align, Group, Rotate) is shared by Picture Format, Shape
 * Format and Layout. It acts on `session.selectedObject` through the adapter
 * registered for that object's kind, so each object family implements the
 * operations against its own markup (DrawingML pictures, VML shapes …).
 */

import type { EditorSession, SelectedObject } from "../session.svelte";
import { pictureArrange, pictureObjects } from "./arrange-picture";

/** Wrap Text menu entries, in Word's order. */
export type ArrangeWrap =
  | "inline"
  | "square"
  | "tight"
  | "through"
  | "topAndBottom"
  | "behindText"
  | "inFrontOfText";

/** Position gallery: In Line with Text, then the nine "With Text Wrapping" squares. */
export type PositionPreset =
  | "inline"
  | "topLeft"
  | "topCenter"
  | "topRight"
  | "middleLeft"
  | "middleCenter"
  | "middleRight"
  | "bottomLeft"
  | "bottomCenter"
  | "bottomRight";

export type ArrangeOrder =
  | "bringForward"
  | "bringToFront"
  | "bringInFrontOfText"
  | "sendBackward"
  | "sendToBack"
  | "sendBehindText";

export type AlignOp =
  | "left"
  | "center"
  | "right"
  | "top"
  | "middle"
  | "bottom"
  | "distributeH"
  | "distributeV";

export type RotateOp = "right90" | "left90" | "flipV" | "flipH";

/** One object family's implementation of the Arrange commands. Omitted optional members show disabled. */
export interface ArrangeOps {
  /** The current wrap style, for the checked Wrap Text item. */
  wrap(s: EditorSession): ArrangeWrap | undefined;
  setWrap(s: EditorSession, wrap: ArrangeWrap): void;
  moveWithText?(s: EditorSession): boolean | undefined;
  setMoveWithText?(s: EditorSession, on: boolean): void;
  position(s: EditorSession, preset: PositionPreset): void;
  order(s: EditorSession, op: ArrangeOrder): void;
  align(s: EditorSession, op: AlignOp, relativeTo: "page" | "margin"): void;
  rotate(s: EditorSession, op: RotateOp): void;
  group?(s: EditorSession): void;
  ungroup?(s: EditorSession): void;
  editWrapPoints?(s: EditorSession): void;
  /** Wrap Text ▸ More Layout Options… and Rotate ▸ More Rotation Options…. */
  layoutOptions?(s: EditorSession): void;
}

export const ARRANGE_OPS: Partial<Record<SelectedObject["kind"], ArrangeOps>> = {
  picture: pictureArrange,
  chart: pictureArrange,
};

/** A row of the Selection Pane. */
export interface PaneObject {
  readonly key: string;
  readonly name: string;
  readonly kind: SelectedObject["kind"];
  readonly hidden: boolean;
  select(): void;
  setHidden?(hidden: boolean): void;
  rename?(name: string): void;
}

/** Sources of Selection Pane rows, one per object family. */
export const OBJECT_LISTERS: Array<(s: EditorSession) => PaneObject[]> = [pictureObjects];

/** The adapter for the selected object, if any. */
export function arrangeOpsFor(s: EditorSession): ArrangeOps | undefined {
  const kind = s.selectedObject?.kind;
  return kind ? ARRANGE_OPS[kind] : undefined;
}
