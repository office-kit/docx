/** Arrange commands for DrawingML pictures and charts (`<w:drawing>`). */

import type { ArrangeOps, PaneObject } from "./arrange";
import type { EditorSession } from "../session.svelte";

export const pictureArrange: ArrangeOps = {
  wrap: () => undefined,
  setWrap: () => {},
  position: () => {},
  order: () => {},
  align: () => {},
  rotate: () => {},
};

export function pictureObjects(_s: EditorSession): PaneObject[] {
  return [];
}
