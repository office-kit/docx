/** The selected VML shape or SmartArt graphic, for the ribbon and panes. */

import type { SmartArtRef, XmlElement } from "@office-kit/docx";
import { commands, type DocPosition } from "@office-kit/docx-editor";
import type { EditorSession } from "../session.svelte";

const SHAPE_KINDS: ReadonlySet<string> = new Set(["shape", "textBox", "ink"]);

/**
 * The selected shape, re-read on every edit (reads `tick`). Undefined when
 * nothing VML is selected, or when the selection went stale after an undo.
 */
export function selectedShape(
  s: EditorSession,
): { at: DocPosition; shape: XmlElement } | undefined {
  const sel = s.selectedObject;
  const model = s.model;
  if (s.tick < 0 || !sel || !model || !SHAPE_KINDS.has(sel.kind)) return undefined;
  try {
    return { at: sel.at, shape: commands.shapeAt(model.doc, sel.at) };
  } catch {
    // Stale selection: the run no longer holds a shape.
    return undefined;
  }
}

export function selectedSmartArt(
  s: EditorSession,
): { at: DocPosition; ref: SmartArtRef } | undefined {
  const sel = s.selectedObject;
  const model = s.model;
  if (s.tick < 0 || sel?.kind !== "smartArt" || !model) return undefined;
  try {
    return { at: sel.at, ref: commands.smartArtAt(model.doc, sel.at) };
  } catch {
    return undefined;
  }
}
