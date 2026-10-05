/**
 * Pointer tools for drawing on the page, shared by the Insert ▸ Shapes /
 * Text Box galleries, the Draw tab and the canvas overlay: which tool the next
 * drag uses, the Draw tab's pens, and the shapes picked for grouping.
 */

import type { DocPosition } from "@office-kit/docx-editor";
import type { ShapePreset, TextBoxLayout } from "@office-kit/docx";

export type PenKind = "pen" | "pencil" | "highlighter";

export interface Pen {
  readonly id: number;
  readonly kind: PenKind;
  /** RRGGBB. */
  readonly color: string;
  /** Points. */
  readonly weight: number;
}

/** What a drag on the page does. */
export type DrawTool =
  | { readonly kind: "none" }
  | {
      readonly kind: "shape";
      readonly preset: ShapePreset;
      readonly textLayout?: Partial<TextBoxLayout>;
    }
  | { readonly kind: "pen"; readonly pen: number }
  | { readonly kind: "eraser" }
  | { readonly kind: "lasso" }
  /** Create Link: the next text box clicked becomes the link target. */
  | { readonly kind: "link"; readonly from: DocPosition };

// Word's default pens on the Draw tab: black pen, red pen, pencil, yellow highlighter.
const DEFAULT_PENS: readonly Pen[] = [
  { id: 1, kind: "pen", color: "000000", weight: 1 },
  { id: 2, kind: "pen", color: "E71224", weight: 1 },
  { id: 3, kind: "pencil", color: "595959", weight: 1 },
  { id: 4, kind: "highlighter", color: "FFFF00", weight: 6 },
];

/** Highlighter ink is translucent so the text shows through. */
export const HIGHLIGHTER_OPACITY = 0.5;
// The pencil's grain is approximated by a slightly translucent stroke.
export const PENCIL_OPACITY = 0.8;

class ShapeTools {
  tool = $state<DrawTool>({ kind: "none" });
  pens = $state<Pen[]>([...DEFAULT_PENS]);
  /** Draw ▸ Ink to Shape: convert each new stroke to the closest shape. */
  inkToShape = $state(false);
  /** Shapes picked with Shift+click or the lasso, for Group. */
  multi = $state<DocPosition[]>([]);
  /** Recently used presets, newest first (Word's "Recently Used Shapes"). */
  recent = $state<ShapePreset[]>([]);
  /** Edit Shape ▸ Edit Points: the selected freeform shows its vertices as handles. */
  editPoints = $state(false);
  /** The Text Pane row holding the caret, for SmartArt Design ▸ Promote / Demote. */
  smartArtRow = $state(0);

  setTool(tool: DrawTool): void {
    this.tool = tool;
  }

  /** Toggle a pen or the eraser: choosing the active tool again turns it off, as Word does. */
  toggle(tool: DrawTool): void {
    const same =
      this.tool.kind === tool.kind &&
      (tool.kind !== "pen" || (this.tool.kind === "pen" && this.tool.pen === tool.pen));
    this.tool = same ? { kind: "none" } : tool;
  }

  pickPreset(preset: ShapePreset, textLayout?: Partial<TextBoxLayout>): void {
    this.tool = textLayout ? { kind: "shape", preset, textLayout } : { kind: "shape", preset };
    this.recent = [preset, ...this.recent.filter((p) => p !== preset)].slice(0, 12);
  }

  pen(id: number): Pen | undefined {
    return this.pens.find((p) => p.id === id);
  }

  updatePen(id: number, change: Partial<Omit<Pen, "id">>): void {
    this.pens = this.pens.map((p) => (p.id === id ? { ...p, ...change } : p));
  }

  addPen(kind: PenKind): void {
    const id = Math.max(0, ...this.pens.map((p) => p.id)) + 1;
    const base = DEFAULT_PENS.find((p) => p.kind === kind) ?? DEFAULT_PENS[0];
    if (!base) return;
    this.pens = [...this.pens, { ...base, id }];
    this.tool = { kind: "pen", pen: id };
  }

  removePen(id: number): void {
    this.pens = this.pens.filter((p) => p.id !== id);
    if (this.tool.kind === "pen" && this.tool.pen === id) this.tool = { kind: "none" };
  }
}

export const shapeTools = new ShapeTools();

/** Two positions name the same object run. */
export function samePosition(a: DocPosition, b: DocPosition): boolean {
  return (
    a.block === b.block &&
    (a.inline ?? 0) === (b.inline ?? 0) &&
    (a.para ?? 0) === (b.para ?? 0) &&
    a.cell?.row === b.cell?.row &&
    a.cell?.col === b.cell?.col
  );
}
