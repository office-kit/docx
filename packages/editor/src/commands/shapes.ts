/**
 * Shape, text box, WordArt, ink and SmartArt commands (Insert ▸ Shapes /
 * Text Box / WordArt / SmartArt, Shape Format, SmartArt Design, Draw).
 *
 * Objects are addressed by the {@link DocPosition} of the run that holds them
 * (`block` / `inline`, plus `cell` / `para` inside a table), the same
 * position the canvas reports for a clicked object.
 */

import {
  addDrawingCanvas,
  addShape,
  type AddShapeOptions,
  addShapeToGroup,
  addSmartArt,
  type AddSmartArtOptions,
  addWordArt,
  type AddWordArtOptions,
  changeShapePreset,
  type Docx,
  groupShapes,
  linkTextBoxes,
  removeShape,
  removeSmartArt,
  runShape,
  runSmartArt,
  setShapeAltText,
  setShapeFill,
  setShapeLayout,
  setShapeMoveWithText,
  setShapeName,
  setShapeOrder,
  setShapePoints,
  setShapeShadow,
  setShapeStroke,
  setShapeText,
  setShapeWrap,
  setSmartArtColors,
  setSmartArtLayout,
  setSmartArtNodes,
  setSmartArtSize,
  setSmartArtStyle,
  setTextBoxLayout,
  setWordArt,
  shapeText,
  type ShapeFill,
  type ShapeLayout,
  type ShapeOrder,
  type ShapePreset,
  type ShapeShadow,
  type ShapeStroke,
  type ShapeWrap,
  type SmartArtColors,
  type SmartArtLayout,
  type SmartArtNode,
  type SmartArtRef,
  type SmartArtStyle,
  type TextBoxLayout,
  ungroupShapes,
  type WmlParagraph,
  type WmlRun,
  type WordArtText,
  type XmlElement,
} from "@office-kit/docx";
import { paragraphAt } from "../doc-access.js";
import type { EditorModel } from "../model.js";
import type { DocPosition } from "../selection.js";
import { runAtPath, setSimpleRunText } from "../text-edit.js";
import type { Command } from "./types.js";

/** The VML shape a position's run holds. Throws when there is none (a stale selection). */
export function shapeAt(doc: Docx, at: DocPosition): XmlElement {
  const run = runAtPath(doc, at);
  const shape = run && runShape(run);
  if (!shape) throw new Error("No shape at the selected position");
  return shape;
}

/** The SmartArt graphic a position's run holds. */
export function smartArtAt(doc: Docx, at: DocPosition): SmartArtRef {
  const run = runAtPath(doc, at);
  const ref = run && runSmartArt(doc, run);
  if (!ref) throw new Error("No SmartArt at the selected position");
  return ref;
}

function anchorParagraph(model: EditorModel, at: DocPosition | undefined): WmlParagraph {
  const pos = at ?? model.selection?.focus ?? { block: 0 };
  const para = paragraphAt(model.doc, pos);
  if (!para) throw new Error("Shapes are anchored to a paragraph; place the caret in one");
  return para;
}

/** The position of the run just appended to the anchor paragraph. */
function positionOfLastRun(anchor: DocPosition, para: WmlParagraph, run: WmlRun): DocPosition {
  const runs = para.children.filter((c): c is WmlRun => c.kind === "run");
  return { ...anchor, inline: runs.indexOf(run), offset: 0 };
}

function lastRunOf(para: WmlParagraph): WmlRun {
  const run = para.children.at(-1);
  if (run?.kind !== "run") throw new Error("The new object's run is missing");
  return run;
}

/** Where a new object goes: the anchor paragraph's position (defaults to the caret's). */
function anchorPosition(model: EditorModel, at: DocPosition | undefined): DocPosition {
  const pos = at ?? model.selection?.focus ?? { block: 0 };
  return pos.cell
    ? { block: pos.block, cell: pos.cell, para: pos.para ?? 0 }
    : { block: pos.block };
}

// ---------------------------------------------------------------------------
// Insert
// ---------------------------------------------------------------------------

/**
 * Insert ▸ Shapes / Text Box: add a shape anchored to the paragraph at `at`
 * (the caret's by default), appended to its end so existing run positions
 * stay valid. Returns the new shape's position.
 */
export const insertShapeCommand: Command<
  { options: AddShapeOptions; at?: DocPosition },
  DocPosition
> = {
  id: "shape.insert",
  group: "shape",
  label: "Insert shape",
  run(model, { options, at }) {
    const anchor = anchorPosition(model, at);
    const para = anchorParagraph(model, anchor);
    addShape(model.doc, para, options);
    return positionOfLastRun(anchor, para, lastRunOf(para));
  },
};

/** Insert ▸ WordArt. */
export const insertWordArtCommand: Command<
  { options: AddWordArtOptions; at?: DocPosition },
  DocPosition
> = {
  id: "shape.insertWordArt",
  group: "shape",
  label: "Insert WordArt",
  run(model, { options, at }) {
    const anchor = anchorPosition(model, at);
    const para = anchorParagraph(model, anchor);
    addWordArt(model.doc, para, options);
    return positionOfLastRun(anchor, para, lastRunOf(para));
  },
};

/** Insert ▸ Shapes ▸ New Drawing Canvas. */
export const insertDrawingCanvasCommand: Command<
  { width: number; height: number; at?: DocPosition },
  DocPosition
> = {
  id: "shape.insertCanvas",
  group: "shape",
  label: "New drawing canvas",
  run(model, { width, height, at }) {
    const anchor = anchorPosition(model, at);
    const para = anchorParagraph(model, anchor);
    addDrawingCanvas(model.doc, para, { width, height });
    return positionOfLastRun(anchor, para, lastRunOf(para));
  },
};

/**
 * Draw ▸ pens: one ink stroke. Inside a drawing canvas (`canvas` set) the
 * stroke is drawn into it; otherwise it floats over the page, anchored to the
 * paragraph at `at`.
 */
export const insertInkCommand: Command<
  { options: Omit<AddShapeOptions, "preset" | "ink">; at?: DocPosition; canvas?: DocPosition },
  DocPosition | undefined
> = {
  id: "shape.insertInk",
  group: "shape",
  label: "Ink",
  run(model, { options, at, canvas }) {
    if (canvas) {
      addShapeToGroup(model.doc, shapeAt(model.doc, canvas), {
        ...options,
        preset: "scribble",
        ink: true,
      });
      return canvas;
    }
    const anchor = anchorPosition(model, at);
    const para = anchorParagraph(model, anchor);
    addShape(model.doc, para, { ...options, preset: "scribble", ink: true });
    return positionOfLastRun(anchor, para, lastRunOf(para));
  },
};

/** Insert ▸ SmartArt. */
export const insertSmartArtCommand: Command<
  { options: AddSmartArtOptions; at?: DocPosition },
  DocPosition
> = {
  id: "smartArt.insert",
  group: "shape",
  label: "Insert SmartArt",
  run(model, { options, at }) {
    const anchor = anchorPosition(model, at);
    const para = anchorParagraph(model, anchor);
    const run = addSmartArt(model.doc, para, options);
    return positionOfLastRun(anchor, para, run);
  },
};

// ---------------------------------------------------------------------------
// Edit a shape
// ---------------------------------------------------------------------------

type At<P = object> = { at: DocPosition } & P;

function shapeCommand<P>(
  id: string,
  label: string,
  apply: (doc: Docx, shape: XmlElement, params: At<P>) => void,
): Command<At<P>> {
  return {
    id,
    group: "shape",
    label,
    run(model, params) {
      apply(model.doc, shapeAt(model.doc, params.at), params);
    },
  };
}

export const deleteShapeCommand = shapeCommand("shape.delete", "Delete shape", (doc, shape) => {
  removeShape(doc, shape);
});

/** Shape Fill. */
export const shapeFillCommand = shapeCommand<{ fill: ShapeFill }>(
  "shape.fill",
  "Shape Fill",
  (doc, shape, { fill }) => setShapeFill(doc, shape, fill),
);

/** Shape Outline (`null` = No Outline). */
export const shapeOutlineCommand = shapeCommand<{ stroke: ShapeStroke | null }>(
  "shape.outline",
  "Shape Outline",
  (_doc, shape, { stroke }) => setShapeStroke(shape, stroke),
);

/** Shape Effects ▸ Shadow (`null` = No Shadow). */
export const shapeShadowCommand = shapeCommand<{ shadow: ShapeShadow | null }>(
  "shape.shadow",
  "Shape Effects",
  (_doc, shape, { shadow }) => setShapeShadow(shape, shadow),
);

/** Move, resize, rotate, flip, align or hide (Size group, Align, Rotate, Selection Pane). */
export const shapeLayoutCommand = shapeCommand<{ layout: Partial<ShapeLayout> }>(
  "shape.layout",
  "Shape layout",
  (_doc, shape, { layout }) => setShapeLayout(shape, layout),
);

/** Wrap Text. */
export const shapeWrapCommand = shapeCommand<{ wrap: ShapeWrap }>(
  "shape.wrap",
  "Wrap Text",
  (doc, shape, { wrap }) => setShapeWrap(doc, shape, wrap),
);

/** Bring Forward / Send Backward. */
export const shapeOrderCommand = shapeCommand<{ order: ShapeOrder }>(
  "shape.order",
  "Arrange order",
  (doc, shape, { order }) => setShapeOrder(doc, shape, order),
);

/** Position ▸ Move with Text. */
export const shapeMoveWithTextCommand = shapeCommand<{ on: boolean }>(
  "shape.moveWithText",
  "Move with Text",
  (_doc, shape, { on }) => setShapeMoveWithText(shape, on),
);

/** Edit Shape ▸ Change Shape. */
export const changeShapeCommand = shapeCommand<{ preset: ShapePreset }>(
  "shape.change",
  "Change Shape",
  (_doc, shape, { preset }) => changeShapePreset(shape, preset),
);

/** Edit Shape ▸ Edit Points. */
export const shapePointsCommand = shapeCommand<{
  points: ReadonlyArray<readonly [number, number]>;
}>("shape.points", "Edit Points", (_doc, shape, { points }) => setShapePoints(shape, points));

/** Replace a shape's text (Add Text, the Text Box dialog). */
export const shapeTextCommand = shapeCommand<{ text: string | readonly WmlParagraph[] }>(
  "shape.text",
  "Edit text",
  (_doc, shape, { text }) => setShapeText(shape, text),
);

/**
 * Typing inside a text box on the canvas: the new text of each run, paragraph
 * by paragraph, as the canvas shows it. A paragraph without runs that gained
 * text gets one; runs holding more than text (tabs, fields) are left alone,
 * as body typing does.
 */
export const shapeTextEditCommand = shapeCommand<{ paragraphs: ReadonlyArray<readonly string[]> }>(
  "shape.textEdit",
  "Typing",
  (_doc, shape, { paragraphs }) => {
    const current = shapeText(shape);
    current.forEach((para, p) => {
      const texts = paragraphs[p];
      if (!texts) return;
      const runs = para.children.filter((c): c is WmlRun => c.kind === "run");
      if (runs.length === 0) {
        const text = texts.join("");
        if (text) para.children.push({ kind: "run", pieces: [], extras: [] });
      }
      const live = para.children.filter((c): c is WmlRun => c.kind === "run");
      live.forEach((run, r) => {
        const text = runs.length === 0 ? texts.join("") : texts[r];
        if (text !== undefined) setSimpleRunText(run, text);
      });
    });
    setShapeText(shape, current);
  },
);

/** Enter inside a text box: split the paragraph at the caret. */
export const shapeSplitTextCommand = shapeCommand<{ para: number; run: number; offset: number }>(
  "shape.splitText",
  "New paragraph",
  (_doc, shape, { para, run, offset }) => {
    const paragraphs = shapeText(shape);
    const target = paragraphs[para];
    if (!target) return;
    const runs = target.children.filter((c): c is WmlRun => c.kind === "run");
    const split = runs[run];
    const tail: WmlRun[] = runs.slice(run + 1);
    if (split) {
      const text = split.pieces.map((p) => (p.kind === "text" ? p.value : "")).join("");
      const after: WmlRun = { ...structuredClone(split), pieces: [] };
      setSimpleRunText(after, text.slice(offset));
      setSimpleRunText(split, text.slice(0, offset));
      tail.unshift(after);
    }
    target.children = target.children.filter((c) => c.kind !== "run" || !tail.includes(c));
    const next: WmlParagraph = {
      kind: "paragraph",
      ...(target.pPr ? { pPr: structuredClone(target.pPr) } : {}),
      children: tail,
      extras: [],
    };
    paragraphs.splice(para + 1, 0, next);
    setShapeText(shape, paragraphs);
  },
);

/** Backspace at the start of a text box paragraph: join it onto the previous one. */
export const shapeMergeTextCommand = shapeCommand<{ para: number }>(
  "shape.mergeText",
  "Join paragraphs",
  (_doc, shape, { para }) => {
    const paragraphs = shapeText(shape);
    const prev = paragraphs[para - 1];
    const target = paragraphs[para];
    if (!prev || !target) return;
    prev.children.push(...target.children);
    paragraphs.splice(para, 1);
    setShapeText(shape, paragraphs);
  },
);

/** Text Direction, Align Text, margins, Resize shape to fit text. */
export const textBoxLayoutCommand = shapeCommand<{ layout: Partial<TextBoxLayout> }>(
  "shape.textLayout",
  "Text box layout",
  (_doc, shape, { layout }) => setTextBoxLayout(shape, layout),
);

/** Create Link / Break Link between text boxes. */
export const linkTextBoxCommand = shapeCommand<{ to: DocPosition | null }>(
  "shape.link",
  "Create Link",
  (doc, shape, { to }) => linkTextBoxes(shape, to ? shapeAt(doc, to) : undefined),
);

/** Alt Text. */
export const shapeAltTextCommand = shapeCommand<{ alt: string }>(
  "shape.altText",
  "Alt Text",
  (_doc, shape, { alt }) => setShapeAltText(shape, alt),
);

/** Selection Pane ▸ rename. */
export const shapeNameCommand = shapeCommand<{ name: string }>(
  "shape.name",
  "Rename",
  (_doc, shape, { name }) => setShapeName(shape, name),
);

/** WordArt text and font. */
export const wordArtTextCommand = shapeCommand<{ text: Partial<WordArtText> }>(
  "shape.wordArt",
  "WordArt text",
  (_doc, shape, { text }) => setWordArt(shape, text),
);

/** Arrange ▸ Group: the shapes at `ats` become one group. Returns the group's position. */
export const groupShapesCommand: Command<{ ats: readonly DocPosition[] }, DocPosition> = {
  id: "shape.group",
  group: "shape",
  label: "Group",
  run(model, { ats }) {
    const shapes = ats.map((at) => shapeAt(model.doc, at));
    const first = ats[0];
    if (!first) throw new Error("Nothing to group");
    groupShapes(model.doc, shapes);
    const para = anchorParagraph(model, first);
    return positionOfLastRun(anchorPosition(model, first), para, lastRunOf(para));
  },
  isEnabled: (model) => !!model.selection,
};

/** Arrange ▸ Ungroup. */
export const ungroupShapesCommand = shapeCommand("shape.ungroup", "Ungroup", (doc, shape) => {
  ungroupShapes(doc, shape);
});

// ---------------------------------------------------------------------------
// SmartArt
// ---------------------------------------------------------------------------

function smartArtCommand<P>(
  id: string,
  label: string,
  apply: (doc: Docx, ref: SmartArtRef, params: At<P>) => void,
): Command<At<P>> {
  return {
    id,
    group: "shape",
    label,
    run(model, params) {
      apply(model.doc, smartArtAt(model.doc, params.at), params);
    },
  };
}

/** Delete a selected SmartArt graphic along with its diagram parts. */
export const deleteSmartArtCommand: Command<{ at: DocPosition }> = {
  id: "smartArt.delete",
  group: "shape",
  label: "Delete SmartArt",
  run(model, { at }) {
    const run = runAtPath(model.doc, at);
    const para = paragraphAt(model.doc, at);
    if (!run || !para || !removeSmartArt(model.doc, para, run)) {
      throw new Error("No SmartArt at the selected position");
    }
  },
};

/** Text Pane. */
export const smartArtNodesCommand = smartArtCommand<{ nodes: readonly SmartArtNode[] }>(
  "smartArt.nodes",
  "Text Pane",
  (doc, ref, { nodes }) => setSmartArtNodes(doc, ref, nodes),
);
export const smartArtLayoutCommand = smartArtCommand<{ layout: SmartArtLayout }>(
  "smartArt.layout",
  "Layouts",
  (doc, ref, { layout }) => setSmartArtLayout(doc, ref, layout),
);
export const smartArtColorsCommand = smartArtCommand<{ colors: SmartArtColors }>(
  "smartArt.colors",
  "Change Colors",
  (doc, ref, { colors }) => setSmartArtColors(doc, ref, colors),
);
export const smartArtStyleCommand = smartArtCommand<{ style: SmartArtStyle }>(
  "smartArt.style",
  "SmartArt Styles",
  (doc, ref, { style }) => setSmartArtStyle(doc, ref, style),
);
export const smartArtSizeCommand = smartArtCommand<{ width: number; height: number }>(
  "smartArt.size",
  "Size",
  (_doc, ref, { width, height }) => setSmartArtSize(ref, width, height),
);

export const shapeCommands = [
  insertShapeCommand,
  insertWordArtCommand,
  insertDrawingCanvasCommand,
  insertInkCommand,
  insertSmartArtCommand,
  deleteShapeCommand,
  shapeFillCommand,
  shapeOutlineCommand,
  shapeShadowCommand,
  shapeLayoutCommand,
  shapeWrapCommand,
  shapeOrderCommand,
  shapeMoveWithTextCommand,
  changeShapeCommand,
  shapePointsCommand,
  shapeTextCommand,
  shapeTextEditCommand,
  shapeSplitTextCommand,
  shapeMergeTextCommand,
  textBoxLayoutCommand,
  linkTextBoxCommand,
  shapeAltTextCommand,
  shapeNameCommand,
  wordArtTextCommand,
  groupShapesCommand,
  ungroupShapesCommand,
  deleteSmartArtCommand,
  smartArtNodesCommand,
  smartArtLayoutCommand,
  smartArtColorsCommand,
  smartArtStyleCommand,
  smartArtSizeCommand,
];
