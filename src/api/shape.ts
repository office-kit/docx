/**
 * Shapes, text boxes, WordArt, drawing canvases and ink strokes.
 *
 * Word writes these with the `wps:` / `wpg:` DrawingML extensions, which are
 * not part of ECMA-376. This module writes them in the standard's own
 * representation for WordprocessingML drawing objects: VML inside `w:pict`
 * (ECMA-376 Part 4, §19.1 VML, §19.2 Office extensions, §19.3 Word
 * extensions), which Word reads, edits and round-trips.
 *
 * Shape handles are the VML elements themselves (`v:shape`, `v:rect`,
 * `v:group` …, as returned by {@link vmlShapes} / {@link addShape}); every
 * setter edits that element in place. Lengths are in points.
 */

import type { HorizontalAlign, VerticalAlign } from "./picture.js";
import {
  addPart,
  addRelationship,
  getPart,
  hasPart,
  partRelationships,
  relationshipById,
  setContentTypeDefault,
} from "../internal/opc/index.js";
import {
  buildTextParagraph,
  paragraphToElement,
  parseParagraph,
  WML_RELATIONSHIPS,
  type WmlParagraph,
  type WmlRun,
  type WmlRunPiece,
} from "../internal/wordprocessingml/index.js";
import type { XmlElement, XmlNode } from "../internal/xml/index.js";
import { VML_NS } from "../internal/vml/namespaces.js";
import { TEXT_PATH_SHAPETYPE_REF, textPathShapetype } from "../internal/vml/shapetypes.js";
import { FILL_PATTERNS, type FillPattern, patternBitmap } from "../internal/vml/pattern.js";
import {
  COORD_SIZE,
  SHAPE_PRESETS,
  type ShapePreset,
  type ShapePresetDef,
} from "../internal/vml/presets.js";
import { lengthToPoints, parseVmlStyle, points, serializeVmlStyle } from "../internal/vml/style.js";
import {
  child,
  children,
  element,
  getAttr,
  isElement,
  nsDecl,
  rename,
  setAttr,
  setChild,
} from "../internal/vml/xml.js";
import type { Docx } from "./docx.js";

export {
  SHAPE_CATEGORIES,
  SHAPE_PRESETS,
  type ShapeCategory,
  type ShapePreset,
  type ShapePresetDef,
} from "../internal/vml/presets.js";
export { FILL_PATTERNS, type FillPattern } from "../internal/vml/pattern.js";

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

/** `mso-position-horizontal-relative` (§19.1.2.19 style): what `left` is measured from. */
export type ShapeHorizontalRelativeTo = "margin" | "page" | "column" | "character";
/** `mso-position-vertical-relative`: what `top` is measured from. */
export type ShapeVerticalRelativeTo = "margin" | "page" | "paragraph" | "line";

/** Word's Wrap Text choices. */
export type ShapeWrap =
  | "inline"
  | "square"
  | "tight"
  | "through"
  | "topAndBottom"
  | "behindText"
  | "inFrontOfText";

/** A shape's box and anchoring. `left` / `top` are ignored while inline. */
export interface ShapeLayout {
  readonly inline: boolean;
  readonly left: number;
  readonly top: number;
  readonly width: number;
  readonly height: number;
  readonly horizontalRelativeTo: ShapeHorizontalRelativeTo;
  readonly verticalRelativeTo: ShapeVerticalRelativeTo;
  readonly horizontalAlign?: HorizontalAlign;
  readonly verticalAlign?: VerticalAlign;
  /** Clockwise rotation in degrees. */
  readonly rotation: number;
  readonly flipH: boolean;
  readonly flipV: boolean;
  readonly zIndex: number;
  /** Hidden through the Selection Pane (`visibility:hidden`). */
  readonly hidden: boolean;
}

/** Shape Fill. Colors are `RRGGBB` hex. */
export type ShapeFill =
  | { readonly type: "none" }
  | { readonly type: "solid"; readonly color: string; readonly opacity?: number }
  | {
      readonly type: "gradient";
      readonly color: string;
      readonly color2: string;
      /** `linear` runs along `angle`; `radial` spreads from the centre. */
      readonly style?: "linear" | "radial";
      /** Degrees; 0 = `color` at the top, `color2` at the bottom. */
      readonly angle?: number;
      readonly opacity?: number;
    }
  | {
      readonly type: "pattern";
      readonly pattern: FillPattern;
      /** Foreground (the pattern's dark pixels). */
      readonly color: string;
      /** Background. */
      readonly color2: string;
    }
  | {
      readonly type: "picture";
      readonly bytes: Uint8Array;
      readonly contentType: string;
      /** Repeat the picture instead of stretching it over the shape. */
      readonly tile?: boolean;
    };

/** What {@link getShapeFill} reports (a picture fill is reported by relationship). */
export type ShapeFillInfo =
  | Exclude<ShapeFill, { type: "picture" }>
  | { readonly type: "picture"; readonly relId: string; readonly tile: boolean };

/** VML `dashstyle` values (§19.1.2.21 `ST_StrokeDashStyle`). */
export type ShapeDash =
  | "solid"
  | "shortdash"
  | "shortdot"
  | "shortdashdot"
  | "shortdashdotdot"
  | "dot"
  | "dash"
  | "longdash"
  | "dashdot"
  | "longdashdot"
  | "longdashdotdot";

/** VML arrowheads (`ST_StrokeArrowType`). */
export type ShapeArrow = "none" | "block" | "classic" | "oval" | "diamond" | "open";

/** Shape Outline. */
export interface ShapeStroke {
  readonly color: string;
  /** Points. */
  readonly weight: number;
  readonly dash?: ShapeDash;
  readonly startArrow?: ShapeArrow;
  readonly endArrow?: ShapeArrow;
  /** 0 – 1; ink highlighter strokes are translucent. */
  readonly opacity?: number;
}

/** Shape Effects ▸ Shadow (`v:shadow`). */
export interface ShapeShadow {
  readonly color: string;
  /** Offset in points (positive = right / down). */
  readonly offsetX: number;
  readonly offsetY: number;
  readonly opacity?: number;
}

/** Text ▸ Text Direction / Align Text and the text box's internal margins. */
export interface TextBoxLayout {
  readonly direction: "horizontal" | "vertical" | "vertical270";
  readonly anchor: "top" | "middle" | "bottom";
  /** Internal margins in points (left, top, right, bottom). */
  readonly inset: readonly [number, number, number, number];
  /** Resize shape to fit text. */
  readonly autoFit: boolean;
}

/** WordArt text and its font (`v:textpath`). */
export interface WordArtText {
  readonly text: string;
  readonly font: string;
  /** Points. */
  readonly size: number;
  readonly bold: boolean;
  readonly italic: boolean;
}

export type ShapeKind = "shape" | "textBox" | "wordArt" | "ink" | "group" | "canvas";

export interface AddShapeOptions {
  readonly preset: ShapePreset;
  readonly width: number;
  readonly height: number;
  /** Omit both `left` and `top` together with `wrap: "inline"` for an inline shape. */
  readonly left?: number;
  readonly top?: number;
  readonly horizontalRelativeTo?: ShapeHorizontalRelativeTo;
  readonly verticalRelativeTo?: ShapeVerticalRelativeTo;
  /** Defaults to In Front of Text, Word's default for a drawn shape. */
  readonly wrap?: ShapeWrap;
  readonly fill?: ShapeFill;
  /** `null` = No Outline. */
  readonly stroke?: ShapeStroke | null;
  readonly shadow?: ShapeShadow | null;
  /** Text inside the shape (a text box for presets that hold text). */
  readonly text?: string | readonly WmlParagraph[];
  /** Freeform / scribble vertices in points from the box's top-left corner. */
  readonly points?: ReadonlyArray<readonly [number, number]>;
  readonly rotation?: number;
  readonly flipH?: boolean;
  readonly flipV?: boolean;
  readonly altText?: string;
  /** The shape's name in the Selection Pane (defaults to Word's "Rectangle 3" style). */
  readonly name?: string;
  /** Text box layout (direction, alignment, margins). */
  readonly textLayout?: Partial<TextBoxLayout>;
  /**
   * An ink stroke drawn with a pen (Draw tab) rather than a shape: named
   * "Ink n" like Word's, which is how {@link shapeKind} tells the two apart
   * (VML has no ink marker; InkML is an extension).
   */
  readonly ink?: boolean;
}

export interface AddWordArtOptions extends Omit<AddShapeOptions, "preset" | "text" | "points"> {
  readonly text: string;
  readonly font?: string;
  readonly size?: number;
  readonly bold?: boolean;
  readonly italic?: boolean;
}

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

const DRAWABLE = new Set([
  "shape",
  "rect",
  "roundrect",
  "oval",
  "line",
  "polyline",
  "curve",
  "arc",
  "group",
  "image",
]);

// Word's z-index bases for the two drawing layers: objects in front of text
// count up from the first, objects behind text from the second.
const FRONT_LAYER_BASE = 251659264;
const BEHIND_LAYER_BASE = -251658240;
// Word numbers shape ids from 1025 (`_x0000_s1025`).
const FIRST_SHAPE_ID = 1025;
const SHAPE_ID = /^_x0000_s(\d+)$/;
// `v:roundrect` corner radius as a fraction of the shorter side (Word's 1/6).
const DEFAULT_ARC_SIZE = "10923f";
// Group child coordinates are kept in twentieths of a point so they stay integers.
const GROUP_UNITS_PER_POINT = 20;
const DEFAULT_FONT = "Calibri";
const DEFAULT_WORDART_SIZE = 36;
// WordArt "plain text" (shape type 136): two guide lines the text is fitted between.
const WORDART_SPT = "136";
const WORDART_PATH = `m0,0l${COORD_SIZE},0m0,${COORD_SIZE}l${COORD_SIZE},${COORD_SIZE}e`;
const INK_NAME = "Ink";
// Default text box margins: 0.1" left/right, 0.05" top/bottom.
const DEFAULT_INSET: readonly [number, number, number, number] = [7.2, 3.6, 7.2, 3.6];

const RELATIVE_H_TO_VML: Readonly<Record<ShapeHorizontalRelativeTo, string>> = {
  margin: "margin",
  page: "page",
  column: "text",
  character: "char",
};
const RELATIVE_V_TO_VML: Readonly<Record<ShapeVerticalRelativeTo, string>> = {
  margin: "margin",
  page: "page",
  paragraph: "text",
  line: "line",
};
const VML_TO_RELATIVE_H: Readonly<Record<string, ShapeHorizontalRelativeTo>> = {
  margin: "margin",
  page: "page",
  text: "column",
  char: "character",
  "left-margin-area": "margin",
  "right-margin-area": "margin",
  "inner-margin-area": "margin",
  "outer-margin-area": "margin",
};
const VML_TO_RELATIVE_V: Readonly<Record<string, ShapeVerticalRelativeTo>> = {
  margin: "margin",
  page: "page",
  text: "paragraph",
  line: "line",
  "top-margin-area": "margin",
  "bottom-margin-area": "margin",
  "inner-margin-area": "margin",
  "outer-margin-area": "margin",
};

/** Word's default look for a newly drawn shape: theme Accent 1 with a darker outline. */
const DEFAULT_FILL: ShapeFill = { type: "solid", color: "4472C4" };
const DEFAULT_STROKE: ShapeStroke = { color: "2F528F", weight: 1 };
const LINE_STROKE: ShapeStroke = { color: "4472C4", weight: 0.5 };
const TEXT_BOX_FILL: ShapeFill = { type: "solid", color: "FFFFFF" };
const TEXT_BOX_STROKE: ShapeStroke = { color: "000000", weight: 0.5 };

// ---------------------------------------------------------------------------
// Finding shapes
// ---------------------------------------------------------------------------

/** Where a `w:pict` lives in the body. */
interface PictLocation {
  readonly paragraph: WmlParagraph;
  readonly run: WmlRun;
  readonly pict: XmlElement;
}

function* bodyParagraphs(doc: Docx): Generator<WmlParagraph> {
  for (const block of doc.document.body.blocks) {
    if (block.kind === "paragraph") yield block;
    else if (block.kind === "table") {
      for (const row of block.rows) for (const cell of row.cells) yield* cell.paragraphs;
    }
  }
}

function* pictLocations(doc: Docx): Generator<PictLocation> {
  for (const paragraph of bodyParagraphs(doc)) {
    for (const inline of paragraph.children) {
      if (inline.kind !== "run") continue;
      for (const piece of inline.pieces) {
        if (piece.kind === "pict") yield { paragraph, run: inline, pict: piece.node };
      }
    }
  }
}

function isDrawable(node: XmlNode): node is XmlElement {
  return node.kind === "element" && node.name.uri === VML_NS && DRAWABLE.has(node.name.local);
}

/** The drawing objects of a `w:pict` (its shape type definitions excluded). */
function drawablesOf(pict: XmlElement): XmlElement[] {
  return pict.children.filter(isDrawable);
}

/** Every top-level VML drawing object in the body, in document order. */
export function vmlShapes(doc: Docx): XmlElement[] {
  const out: XmlElement[] = [];
  for (const { pict } of pictLocations(doc)) out.push(...drawablesOf(pict));
  return out;
}

/** The VML drawing object a run holds (`w:r/w:pict/v:*`), if any. */
export function runShape(run: WmlRun): XmlElement | undefined {
  for (const piece of run.pieces) {
    if (piece.kind !== "pict") continue;
    const found = drawablesOf(piece.node)[0];
    if (found) return found;
  }
  return undefined;
}

function locate(doc: Docx, shape: XmlElement): PictLocation | undefined {
  for (const loc of pictLocations(doc)) if (loc.pict.children.includes(shape)) return loc;
  return undefined;
}

// ---------------------------------------------------------------------------
// Ids and names
// ---------------------------------------------------------------------------

function* allVmlElements(el: XmlElement): Generator<XmlElement> {
  yield el;
  for (const c of children(el)) yield* allVmlElements(c);
}

function nextShapeId(doc: Docx): number {
  let max = FIRST_SHAPE_ID - 1;
  for (const { pict } of pictLocations(doc)) {
    for (const el of allVmlElements(pict)) {
      const m = SHAPE_ID.exec(getAttr(el, "o:spid") ?? getAttr(el, "id") ?? "");
      if (m) max = Math.max(max, Number(m[1]));
    }
  }
  return max + 1;
}

/** "rightArrow" → "Right Arrow": Word's default shape names are the preset's name. */
function presetLabel(preset: ShapePreset): string {
  if (preset === "textBox") return "Text Box";
  if (preset === "ellipse") return "Oval";
  if (preset === "rect") return "Rectangle";
  if (preset === "line") return "Straight Connector";
  const words = preset.replace(/([A-Z0-9]+)/g, " $1");
  return words.charAt(0).toUpperCase() + words.slice(1);
}

// ---------------------------------------------------------------------------
// Building
// ---------------------------------------------------------------------------

function color(hex: string): string {
  return `#${hex.replace(/^#/, "")}`;
}

function readColor(value: string | undefined): string | undefined {
  if (!value) return undefined;
  const hex = /^#?([0-9a-f]{6})/i.exec(value.trim());
  if (hex?.[1]) return hex[1].toUpperCase();
  const short = /^#([0-9a-f])([0-9a-f])([0-9a-f])$/i.exec(value.trim());
  if (short)
    return `${short[1]}${short[1]}${short[2]}${short[2]}${short[3]}${short[3]}`.toUpperCase();
  return NAMED_COLORS[value.trim().toLowerCase()];
}

// VML accepts the 16 HTML color names (§19.1.2.x ST_ColorType).
const NAMED_COLORS: Readonly<Record<string, string>> = {
  black: "000000",
  white: "FFFFFF",
  red: "FF0000",
  green: "008000",
  blue: "0000FF",
  yellow: "FFFF00",
  silver: "C0C0C0",
  gray: "808080",
  maroon: "800000",
  purple: "800080",
  fuchsia: "FF00FF",
  lime: "00FF00",
  olive: "808000",
  navy: "000080",
  teal: "008080",
  aqua: "00FFFF",
};

/** VML fractions are written as plain decimals or as 1/65536ths with an `f` suffix. */
function readFraction(value: string | undefined): number | undefined {
  if (value === undefined) return undefined;
  const v = value.trim();
  if (v.endsWith("f")) return Number(v.slice(0, -1)) / 65536;
  if (v.endsWith("%")) return Number(v.slice(0, -1)) / 100;
  const n = Number(v);
  return Number.isFinite(n) ? n : undefined;
}

function pathFromPoints(
  points: ReadonlyArray<readonly [number, number]>,
  width: number,
  height: number,
): string {
  const scale = (v: number, size: number): number => Math.round((v / (size || 1)) * COORD_SIZE);
  const coords = points.map(([x, y]) => `${scale(x, width)},${scale(y, height)}`);
  const [first, ...rest] = coords;
  if (!first) throw new Error("A freeform needs at least one point");
  return rest.length ? `m${first}l${rest.join(",")}e` : `m${first}e`;
}

function newShapeElement(preset: ShapePreset): XmlElement {
  const def: ShapePresetDef = SHAPE_PRESETS[preset];
  const el = element(`v:${def.element}`);
  applyPresetGeometry(el, def);
  return el;
}

function applyPresetGeometry(el: XmlElement, def: ShapePresetDef): void {
  setAttr(el, "o:spt", def.spt === undefined ? undefined : String(def.spt));
  setAttr(el, "coordsize", def.element === "shape" ? `${COORD_SIZE},${COORD_SIZE}` : undefined);
  setAttr(el, "path", def.element === "shape" ? def.path : undefined);
  setAttr(el, "arcsize", def.element === "roundrect" ? DEFAULT_ARC_SIZE : undefined);
  setAttr(el, "o:oned", def.oneD ? "t" : undefined);
}

function shapePiece(shape: XmlElement): WmlRunPiece {
  const pict = element("w:pict", {}, [shape]);
  (pict.attrs as XmlElement["attrs"][number][]).push(nsDecl("v"), nsDecl("o"), nsDecl("w10"));
  return { kind: "pict", node: pict };
}

function initShape(
  doc: Docx,
  el: XmlElement,
  options: Omit<AddShapeOptions, "preset">,
  defaults: { fill?: ShapeFill | undefined; stroke: ShapeStroke | null; label: string },
): void {
  const id = nextShapeId(doc);
  setAttr(el, "id", options.name ?? `${defaults.label} ${id - FIRST_SHAPE_ID + 1}`);
  setAttr(el, "o:spid", `_x0000_s${id}`);
  const wrap = options.wrap ?? "inFrontOfText";
  const style = new Map<string, string>();
  if (wrap !== "inline") {
    style.set("position", "absolute");
    style.set("margin-left", points(options.left ?? 0));
    style.set("margin-top", points(options.top ?? 0));
  }
  style.set("width", points(options.width));
  style.set("height", points(options.height));
  setAttr(el, "style", serializeVmlStyle(style));
  setShapeLayout(el, {
    ...(options.horizontalRelativeTo ? { horizontalRelativeTo: options.horizontalRelativeTo } : {}),
    ...(options.verticalRelativeTo ? { verticalRelativeTo: options.verticalRelativeTo } : {}),
    ...(options.rotation === undefined ? {} : { rotation: options.rotation }),
    ...(options.flipH === undefined ? {} : { flipH: options.flipH }),
    ...(options.flipV === undefined ? {} : { flipV: options.flipV }),
  });
  if (wrap !== "inline") applyWrap(doc, el, wrap);
  const fill = options.fill ?? defaults.fill;
  if (fill) setShapeFill(doc, el, fill);
  else setAttr(el, "filled", "f");
  setShapeStroke(el, options.stroke === undefined ? defaults.stroke : options.stroke);
  if (options.shadow) setShapeShadow(el, options.shadow);
  if (options.altText !== undefined) setShapeAltText(el, options.altText);
}

/**
 * Add a shape, text box, freeform or ink stroke as a new run at the end of
 * `paragraph` (or at inline index `index`), anchored to that paragraph.
 * Returns the VML element, the handle every other function in this module
 * takes.
 */
export function addShape(
  doc: Docx,
  paragraph: WmlParagraph,
  options: AddShapeOptions,
  index: number = paragraph.children.length,
): XmlElement {
  const def: ShapePresetDef = SHAPE_PRESETS[options.preset];
  const el = newShapeElement(options.preset);
  if (options.points)
    setAttr(el, "path", pathFromPoints(options.points, options.width, options.height));
  const textBox = options.preset === "textBox";
  const unfilled = def.filled === false;
  initShape(doc, el, options, {
    fill: textBox ? TEXT_BOX_FILL : unfilled ? undefined : DEFAULT_FILL,
    stroke: textBox ? TEXT_BOX_STROKE : unfilled ? LINE_STROKE : DEFAULT_STROKE,
    label: options.ink ? INK_NAME : presetLabel(options.preset),
  });
  if (def.startArrow || def.endArrow) {
    const stroke = child(el, "v:stroke") ?? element("v:stroke");
    if (def.startArrow) setAttr(stroke, "startarrow", "block");
    if (def.endArrow) setAttr(stroke, "endarrow", "block");
    setChild(el, "v:stroke", stroke);
  }
  if (options.text !== undefined || textBox) setShapeText(el, options.text ?? "");
  if (options.textLayout) setTextBoxLayout(el, options.textLayout);
  insertRun(doc, paragraph, el, index);
  return el;
}

function insertRun(doc: Docx, paragraph: WmlParagraph, shape: XmlElement, index: number): WmlRun {
  const run: WmlRun = { kind: "run", pieces: [shapePiece(shape)], extras: [] };
  paragraph.children.splice(index, 0, run);
  doc.dirty = true;
  return run;
}

/**
 * Add WordArt: a text path shape (`v:textpath`, shape type 136) whose text is
 * stretched to fill the shape, like Insert ▸ WordArt.
 */
export function addWordArt(
  doc: Docx,
  paragraph: WmlParagraph,
  options: AddWordArtOptions,
  index: number = paragraph.children.length,
): XmlElement {
  const el = element("v:shape", {
    type: TEXT_PATH_SHAPETYPE_REF,
    coordsize: `${COORD_SIZE},${COORD_SIZE}`,
    "o:spt": WORDART_SPT,
    path: WORDART_PATH,
  });
  initShape(doc, el, options, { fill: DEFAULT_FILL, stroke: null, label: "WordArt" });
  setChild(el, "v:path", element("v:path", { textpathok: "t" }));
  setChild(el, "v:textpath", element("v:textpath", { on: "t", fitshape: "t" }));
  setWordArt(el, {
    text: options.text,
    font: options.font ?? DEFAULT_FONT,
    size: options.size ?? DEFAULT_WORDART_SIZE,
    bold: options.bold ?? false,
    italic: options.italic ?? false,
  });
  const run = insertRun(doc, paragraph, el, index);
  // Word shows a text path only through its shape type definition; it is
  // repeated in each WordArt's pict, as Word itself does for pasted shapes.
  const pict = run.pieces[0];
  if (pict?.kind === "pict") (pict.node.children as XmlNode[]).unshift(textPathShapetype());
  return el;
}

/**
 * Add a drawing canvas (Insert ▸ Shapes ▸ New Drawing Canvas): a VML group
 * marked `editas="canvas"` that shapes and ink can be drawn into with
 * {@link addShapeToGroup}. It is inline, as Word inserts it.
 */
export function addDrawingCanvas(
  doc: Docx,
  paragraph: WmlParagraph,
  size: { readonly width: number; readonly height: number },
  index: number = paragraph.children.length,
): XmlElement {
  const id = nextShapeId(doc);
  const w = Math.round(size.width * GROUP_UNITS_PER_POINT);
  const h = Math.round(size.height * GROUP_UNITS_PER_POINT);
  const background = element("v:rect", {
    id: `Canvas Background ${id - FIRST_SHAPE_ID + 1}`,
    "o:spid": `_x0000_s${id + 1}`,
    style: `position:absolute;left:0;top:0;width:${w};height:${h}`,
    filled: "f",
    stroked: "f",
  });
  const group = element(
    "v:group",
    {
      id: `Canvas ${id - FIRST_SHAPE_ID + 1}`,
      "o:spid": `_x0000_s${id}`,
      editas: "canvas",
      style: `width:${points(size.width)};height:${points(size.height)}`,
      coordorigin: "0,0",
      coordsize: `${w},${h}`,
    },
    [background],
  );
  insertRun(doc, paragraph, group, index);
  return group;
}

/** Group coordinate space: the group's box in points and its child coordinate system. */
function groupFrame(group: XmlElement): {
  originX: number;
  originY: number;
  scaleX: number;
  scaleY: number;
} {
  const layout = getShapeLayout(group);
  const [ox = 0, oy = 0] = (getAttr(group, "coordorigin") ?? "0,0").split(",").map(Number);
  const [cw = 1000, ch = 1000] = (getAttr(group, "coordsize") ?? "1000,1000")
    .split(",")
    .map(Number);
  return {
    originX: ox,
    originY: oy,
    scaleX: cw / (layout.width || 1),
    scaleY: ch / (layout.height || 1),
  };
}

/**
 * Draw a shape (or ink stroke) inside a group or drawing canvas. `left` /
 * `top` are points from the group's top-left corner.
 */
export function addShapeToGroup(
  doc: Docx,
  group: XmlElement,
  options: AddShapeOptions,
): XmlElement {
  const el = newShapeElement(options.preset);
  if (options.points)
    setAttr(el, "path", pathFromPoints(options.points, options.width, options.height));
  const def: ShapePresetDef = SHAPE_PRESETS[options.preset];
  const unfilled = def.filled === false;
  initShape(
    doc,
    el,
    { ...options, wrap: "inline" },
    {
      fill: unfilled ? undefined : DEFAULT_FILL,
      stroke: unfilled ? LINE_STROKE : DEFAULT_STROKE,
      label: options.ink ? INK_NAME : presetLabel(options.preset),
    },
  );
  const frame = groupFrame(group);
  const style = parseVmlStyle(getAttr(el, "style"));
  style.set("position", "absolute");
  style.set("left", String(Math.round(frame.originX + (options.left ?? 0) * frame.scaleX)));
  style.set("top", String(Math.round(frame.originY + (options.top ?? 0) * frame.scaleY)));
  style.set("width", String(Math.round(options.width * frame.scaleX)));
  style.set("height", String(Math.round(options.height * frame.scaleY)));
  setAttr(el, "style", serializeVmlStyle(style));
  if (options.text !== undefined) setShapeText(el, options.text);
  (group.children as XmlNode[]).push(el);
  (group as { selfClosing: boolean }).selfClosing = false;
  doc.dirty = true;
  return el;
}

// ---------------------------------------------------------------------------
// Reading
// ---------------------------------------------------------------------------

/** What kind of drawing object a VML element is, for choosing the right UI. */
export function shapeKind(shape: XmlElement): ShapeKind {
  if (isElement(shape, "v:group"))
    return getAttr(shape, "editas") === "canvas" ? "canvas" : "group";
  const textpath = child(shape, "v:textpath");
  if (textpath && getAttr(textpath, "on") !== "f" && getAttr(textpath, "on") !== "false")
    return "wordArt";
  if (child(shape, "v:textbox")) return "textBox";
  if ((getAttr(shape, "id") ?? "").startsWith(`${INK_NAME} `)) return "ink";
  return "shape";
}

let presetBySpt: Map<string, ShapePreset> | undefined;

/**
 * The gallery preset a shape was drawn from: by shape type number, by VML
 * primitive, or by its exact path. `undefined` for custom geometry.
 */
export function shapePreset(shape: XmlElement): ShapePreset | undefined {
  if (isElement(shape, "v:rect")) return "rect";
  if (isElement(shape, "v:roundrect")) return "roundRect";
  if (isElement(shape, "v:oval")) return "ellipse";
  if (isElement(shape, "v:line")) return "line";
  if (isElement(shape, "v:polyline")) return "freeform";
  if (!isElement(shape, "v:shape")) return undefined;
  if (!presetBySpt) {
    presetBySpt = new Map();
    for (const [name, def] of Object.entries(SHAPE_PRESETS) as Array<
      [ShapePreset, ShapePresetDef]
    >) {
      if (def.spt !== undefined && !presetBySpt.has(String(def.spt)))
        presetBySpt.set(String(def.spt), name);
    }
  }
  const spt = getAttr(shape, "o:spt") ?? sptFromType(shape);
  if (spt === WORDART_SPT) return undefined;
  const bySpt = spt === undefined ? undefined : presetBySpt.get(spt);
  if (bySpt) return bySpt;
  const path = getAttr(shape, "path");
  for (const [name, def] of Object.entries(SHAPE_PRESETS) as Array<[ShapePreset, ShapePresetDef]>) {
    if (def.path !== undefined && def.path === path) return name;
  }
  return path ? "freeform" : undefined;
}

// Word's own shape type ids are `_x0000_t<spt>`.
function sptFromType(shape: XmlElement): string | undefined {
  return /^#_x0000_t(\d+)$/.exec(getAttr(shape, "type") ?? "")?.[1];
}

function readStyle(shape: XmlElement): Map<string, string> {
  return parseVmlStyle(getAttr(shape, "style"));
}

const H_ALIGNS: ReadonlySet<string> = new Set(["left", "center", "right", "inside", "outside"]);
const V_ALIGNS: ReadonlySet<string> = new Set(["top", "center", "bottom", "inside", "outside"]);
const isHorizontalAlign = (v: string): v is HorizontalAlign => H_ALIGNS.has(v);
const isVerticalAlign = (v: string): v is VerticalAlign => V_ALIGNS.has(v);

/** The shape's box, anchoring, rotation, flip and z-order. */
export function getShapeLayout(shape: XmlElement): ShapeLayout {
  const style = readStyle(shape);
  const inGroup = !style.has("margin-left") && (style.has("left") || style.has("top"));
  const unit = inGroup ? "raw" : "px";
  const flip = style.get("flip") ?? "";
  const hAlign = style.get("mso-position-horizontal");
  const vAlign = style.get("mso-position-vertical");
  return {
    inline: style.get("position") !== "absolute",
    left: lengthToPoints(style.get("margin-left") ?? style.get("left"), unit) ?? 0,
    top: lengthToPoints(style.get("margin-top") ?? style.get("top"), unit) ?? 0,
    width: lengthToPoints(style.get("width"), unit) ?? 0,
    height: lengthToPoints(style.get("height"), unit) ?? 0,
    horizontalRelativeTo:
      VML_TO_RELATIVE_H[style.get("mso-position-horizontal-relative") ?? ""] ?? "column",
    verticalRelativeTo:
      VML_TO_RELATIVE_V[style.get("mso-position-vertical-relative") ?? ""] ?? "paragraph",
    ...(hAlign && isHorizontalAlign(hAlign) ? { horizontalAlign: hAlign } : {}),
    ...(vAlign && isVerticalAlign(vAlign) ? { verticalAlign: vAlign } : {}),
    rotation: Number.parseFloat(style.get("rotation") ?? "0") || 0,
    flipH: flip.includes("x"),
    flipV: flip.includes("y"),
    zIndex: Number.parseInt(style.get("z-index") ?? "0", 10) || 0,
    hidden: style.get("visibility") === "hidden",
  };
}

/**
 * Change any part of a shape's layout. Setting `left` / `top` clears the
 * matching alignment (an aligned shape ignores its offset); pass
 * `horizontalAlign: undefined` explicitly to clear one without moving.
 */
export function setShapeLayout(shape: XmlElement, layout: Partial<ShapeLayout>): void {
  const style = readStyle(shape);
  const inGroup = !style.has("margin-left") && (style.has("left") || style.has("top"));
  const len = (v: number): string => (inGroup ? String(Math.round(v)) : points(v));
  if (layout.inline !== undefined) {
    if (layout.inline) {
      for (const k of [
        "position",
        "margin-left",
        "margin-top",
        "z-index",
        "mso-position-horizontal",
        "mso-position-horizontal-relative",
        "mso-position-vertical",
        "mso-position-vertical-relative",
      ])
        style.delete(k);
    } else if (style.get("position") !== "absolute") {
      style.set("position", "absolute");
      style.set("margin-left", style.get("margin-left") ?? "0");
      style.set("margin-top", style.get("margin-top") ?? "0");
    }
  }
  if (layout.left !== undefined) {
    style.set(inGroup ? "left" : "margin-left", len(layout.left));
    style.delete("mso-position-horizontal");
  }
  if (layout.top !== undefined) {
    style.set(inGroup ? "top" : "margin-top", len(layout.top));
    style.delete("mso-position-vertical");
  }
  if (layout.width !== undefined) style.set("width", len(Math.max(0, layout.width)));
  if (layout.height !== undefined) style.set("height", len(Math.max(0, layout.height)));
  if (layout.horizontalRelativeTo !== undefined)
    style.set("mso-position-horizontal-relative", RELATIVE_H_TO_VML[layout.horizontalRelativeTo]);
  if (layout.verticalRelativeTo !== undefined)
    style.set("mso-position-vertical-relative", RELATIVE_V_TO_VML[layout.verticalRelativeTo]);
  if ("horizontalAlign" in layout) {
    if (layout.horizontalAlign && isHorizontalAlign(layout.horizontalAlign))
      style.set("mso-position-horizontal", layout.horizontalAlign);
    else style.delete("mso-position-horizontal");
  }
  if ("verticalAlign" in layout) {
    if (layout.verticalAlign && isVerticalAlign(layout.verticalAlign))
      style.set("mso-position-vertical", layout.verticalAlign);
    else style.delete("mso-position-vertical");
  }
  if (layout.rotation !== undefined) {
    const deg = ((Math.round(layout.rotation) % 360) + 360) % 360;
    if (deg) style.set("rotation", String(deg));
    else style.delete("rotation");
  }
  if (layout.flipH !== undefined || layout.flipV !== undefined) {
    const current = style.get("flip") ?? "";
    const x = layout.flipH ?? current.includes("x");
    const y = layout.flipV ?? current.includes("y");
    const flip = [x ? "x" : "", y ? "y" : ""].filter(Boolean).join(" ");
    if (flip) style.set("flip", flip);
    else style.delete("flip");
  }
  if (layout.zIndex !== undefined) style.set("z-index", String(layout.zIndex));
  if (layout.hidden !== undefined) {
    if (layout.hidden) style.set("visibility", "hidden");
    else style.delete("visibility");
  }
  setAttr(shape, "style", serializeVmlStyle(style));
}

// ---------------------------------------------------------------------------
// Wrap and z-order
// ---------------------------------------------------------------------------

const WRAP_TYPES: ReadonlySet<string> = new Set(["square", "tight", "through", "topAndBottom"]);

/** Wrap Text: how body text flows around the shape. */
export function getShapeWrap(shape: XmlElement): ShapeWrap {
  if (getShapeLayout(shape).inline) return "inline";
  const wrapEl = child(shape, "w10:wrap");
  const type = wrapEl && getAttr(wrapEl, "type");
  if (type && WRAP_TYPES.has(type)) return type as ShapeWrap;
  return getShapeLayout(shape).zIndex < 0 ? "behindText" : "inFrontOfText";
}

/**
 * Set Wrap Text. Behind / In Front of Text move the shape to the matching
 * drawing layer (negative / positive z-index, as Word does); In Line with
 * Text drops the floating position.
 */
export function setShapeWrap(doc: Docx, shape: XmlElement, wrap: ShapeWrap): void {
  if (wrap === "inline") {
    setShapeLayout(shape, { inline: true });
    setChild(shape, "w10:wrap", undefined);
    return;
  }
  if (getShapeLayout(shape).inline) setShapeLayout(shape, { inline: false });
  applyWrap(doc, shape, wrap);
}

function applyWrap(doc: Docx, shape: XmlElement, wrap: Exclude<ShapeWrap, "inline">): void {
  setChild(
    shape,
    "w10:wrap",
    WRAP_TYPES.has(wrap) ? element("w10:wrap", { type: wrap }) : undefined,
  );
  const behind = wrap === "behindText";
  const z = getShapeLayout(shape).zIndex;
  if (z !== 0 && z < 0 === behind) return;
  setShapeLayout(shape, { zIndex: topOfLayer(doc, behind ? "behind" : "front", shape) });
}

function floatingShapes(doc: Docx): XmlElement[] {
  return vmlShapes(doc).filter((s) => !getShapeLayout(s).inline);
}

function topOfLayer(doc: Docx, layer: "front" | "behind", except?: XmlElement): number {
  const zs = floatingShapes(doc)
    .filter((s) => s !== except)
    .map((s) => getShapeLayout(s).zIndex)
    .filter((z) => (layer === "front" ? z >= 0 : z < 0));
  const base = layer === "front" ? FRONT_LAYER_BASE : BEHIND_LAYER_BASE;
  return zs.length ? Math.max(...zs) + 1 : base;
}

function bottomOfLayer(doc: Docx, layer: "front" | "behind", except: XmlElement): number {
  const zs = floatingShapes(doc)
    .filter((s) => s !== except)
    .map((s) => getShapeLayout(s).zIndex)
    .filter((z) => (layer === "front" ? z >= 0 : z < 0));
  const base = layer === "front" ? FRONT_LAYER_BASE : BEHIND_LAYER_BASE;
  return zs.length ? Math.min(...zs) - 1 : base;
}

export type ShapeOrder =
  | "bringForward"
  | "bringToFront"
  | "bringInFrontOfText"
  | "sendBackward"
  | "sendToBack"
  | "sendBehindText";

/**
 * Arrange ▸ Bring Forward / Send Backward. Forward / backward swap places with
 * the next shape in the same layer; front / back move to that layer's end;
 * In Front of / Behind Text change layer.
 */
export function setShapeOrder(doc: Docx, shape: XmlElement, order: ShapeOrder): void {
  const z = getShapeLayout(shape).zIndex;
  const layer = z < 0 ? "behind" : "front";
  switch (order) {
    case "bringToFront":
      setShapeLayout(shape, { zIndex: topOfLayer(doc, layer, shape) });
      return;
    case "sendToBack":
      setShapeLayout(shape, { zIndex: bottomOfLayer(doc, layer, shape) });
      return;
    case "bringInFrontOfText":
      setShapeWrap(doc, shape, "inFrontOfText");
      return;
    case "sendBehindText":
      setShapeWrap(doc, shape, "behindText");
      return;
    case "bringForward":
    case "sendBackward": {
      const sameLayer = floatingShapes(doc).filter(
        (s) => s !== shape && getShapeLayout(s).zIndex < 0 === (layer === "behind"),
      );
      const forward = order === "bringForward";
      let neighbour: XmlElement | undefined;
      let neighbourZ = forward ? Infinity : -Infinity;
      for (const s of sameLayer) {
        const sz = getShapeLayout(s).zIndex;
        if (forward ? sz >= z && sz < neighbourZ : sz <= z && sz > neighbourZ) {
          neighbour = s;
          neighbourZ = sz;
        }
      }
      if (!neighbour) return;
      // Equal z-indexes stack in document order; nudge past the neighbour.
      const target = neighbourZ === z ? z + (forward ? 1 : -1) : neighbourZ;
      setShapeLayout(neighbour, { zIndex: z });
      setShapeLayout(shape, { zIndex: target });
      return;
    }
  }
}

/** Arrange ▸ Position ▸ Move with Text: anchored to its paragraph rather than the page. */
export function setShapeMoveWithText(shape: XmlElement, on: boolean): void {
  const layout = getShapeLayout(shape);
  if (on === (layout.verticalRelativeTo === "paragraph")) return;
  // Keep the shape where it is on the page: Word converts the offset too, but
  // without layout information the margin is the closest stable reference.
  setShapeLayout(shape, { verticalRelativeTo: on ? "paragraph" : "margin" });
}

// ---------------------------------------------------------------------------
// Fill, outline, shadow
// ---------------------------------------------------------------------------

const PICTURE_EXTENSIONS: Readonly<Record<string, string>> = {
  "image/png": "png",
  "image/jpeg": "jpeg",
  "image/gif": "gif",
  "image/bmp": "bmp",
  "image/tiff": "tiff",
};

function addImagePart(doc: Docx, bytes: Uint8Array, contentType: string): string {
  const ext = PICTURE_EXTENSIONS[contentType];
  if (!ext) throw new Error(`Unsupported picture fill type "${contentType}"`);
  let n = 1;
  while (hasPart(doc.opc, `/word/media/image${n}.${ext}`)) n++;
  const name = `/word/media/image${n}.${ext}`;
  addPart(doc.opc, { name, contentType, data: bytes });
  setContentTypeDefault(doc.opc.contentTypes, ext, contentType);
  const rels = partRelationships(doc.opc, doc.partName);
  return addRelationship(rels, { type: WML_RELATIONSHIPS.image, target: `media/image${n}.${ext}` })
    .id;
}

/** Shape Fill: no fill, a solid colour, a gradient, a pattern or a picture. */
export function setShapeFill(doc: Docx, shape: XmlElement, fill: ShapeFill): void {
  setAttr(shape, "filled", fill.type === "none" ? "f" : undefined);
  if (fill.type === "none") {
    setChild(shape, "v:fill", undefined);
    return;
  }
  let fillEl: XmlElement | undefined;
  switch (fill.type) {
    case "solid":
      setAttr(shape, "fillcolor", color(fill.color));
      if (fill.opacity !== undefined && fill.opacity < 1)
        fillEl = element("v:fill", { opacity: String(fill.opacity) });
      break;
    case "gradient":
      setAttr(shape, "fillcolor", color(fill.color));
      fillEl = element("v:fill", {
        type: fill.style === "radial" ? "gradientRadial" : "gradient",
        color2: color(fill.color2),
        angle: fill.style === "radial" ? undefined : String(fill.angle ?? 0),
        focusposition: fill.style === "radial" ? ".5,.5" : undefined,
        focussize: fill.style === "radial" ? "" : undefined,
        opacity: fill.opacity === undefined || fill.opacity >= 1 ? undefined : String(fill.opacity),
      });
      break;
    case "pattern":
      setAttr(shape, "fillcolor", color(fill.color));
      fillEl = element("v:fill", {
        type: "pattern",
        "o:title": fill.pattern,
        color2: color(fill.color2),
        "r:id": addImagePart(doc, patternBitmap(fill.pattern), "image/bmp"),
      });
      break;
    case "picture":
      fillEl = element("v:fill", {
        type: fill.tile ? "tile" : "frame",
        "r:id": addImagePart(doc, fill.bytes, fill.contentType),
        recolor: "t",
      });
      break;
  }
  if (fillEl) (fillEl.attrs as XmlElement["attrs"][number][]).push(nsDecl("r"));
  setChild(shape, "v:fill", fillEl);
  doc.dirty = true;
}

/** The shape's fill as written (VML defaults: filled, white). */
export function getShapeFill(shape: XmlElement): ShapeFillInfo {
  const filled = getAttr(shape, "filled");
  const fillEl = child(shape, "v:fill");
  if (filled === "f" || filled === "false" || (fillEl && getAttr(fillEl, "on") === "f"))
    return { type: "none" };
  const base =
    readColor(getAttr(shape, "fillcolor")) ??
    readColor(fillEl && getAttr(fillEl, "color")) ??
    "FFFFFF";
  const type = fillEl ? getAttr(fillEl, "type") : undefined;
  const opacity = readFraction(fillEl && getAttr(fillEl, "opacity"));
  const color2 = readColor(fillEl && getAttr(fillEl, "color2")) ?? "FFFFFF";
  if (fillEl && (type === "gradient" || type === "gradientRadial")) {
    return {
      type: "gradient",
      color: base,
      color2,
      style: type === "gradientRadial" ? "radial" : "linear",
      angle: Number(getAttr(fillEl, "angle") ?? 0),
      ...(opacity === undefined ? {} : { opacity }),
    };
  }
  if (fillEl && type === "pattern") {
    const title = getAttr(fillEl, "o:title") ?? "";
    return {
      type: "pattern",
      pattern: isFillPattern(title) ? title : "pct50",
      color: base,
      color2,
    };
  }
  if (fillEl && (type === "frame" || type === "tile")) {
    return {
      type: "picture",
      relId: getAttr(fillEl, "r:id") ?? getAttr(fillEl, "o:relid") ?? "",
      tile: type === "tile",
    };
  }
  return { type: "solid", color: base, ...(opacity === undefined ? {} : { opacity }) };
}

function isFillPattern(name: string): name is FillPattern {
  return Object.hasOwn(FILL_PATTERNS, name);
}

/**
 * The image a VML element references by relationship id (a picture or
 * pattern fill's `r:id`, or `v:imagedata`), for renderers. `partName` is the
 * part holding the VML (a header, for a watermark). `undefined` when the
 * relationship or its target is missing.
 */
export function vmlImageData(
  doc: Docx,
  relId: string,
  partName: string = doc.partName,
): { bytes: Uint8Array; contentType: string } | undefined {
  const rel = relationshipById(partRelationships(doc.opc, partName), relId);
  if (!rel || rel.targetMode === "External") return undefined;
  const folder = partName.slice(0, partName.lastIndexOf("/") + 1);
  const target = rel.target.startsWith("/")
    ? rel.target
    : `${folder}${rel.target.replace(/^\.\//, "")}`;
  const part = getPart(doc.opc, target);
  return part ? { bytes: part.data, contentType: part.contentType ?? "" } : undefined;
}

/** Shape Outline; `null` = No Outline. */
export function setShapeStroke(shape: XmlElement, stroke: ShapeStroke | null): void {
  if (stroke === null) {
    setAttr(shape, "stroked", "f");
    return;
  }
  setAttr(shape, "stroked", undefined);
  setAttr(shape, "strokecolor", color(stroke.color));
  setAttr(shape, "strokeweight", points(stroke.weight));
  const existing = child(shape, "v:stroke");
  const strokeEl = existing ?? element("v:stroke");
  const set = (name: string, value: string | undefined, fallback: string): void =>
    setAttr(strokeEl, name, value === undefined || value === fallback ? undefined : value);
  set("dashstyle", stroke.dash, "solid");
  if (stroke.startArrow !== undefined) set("startarrow", stroke.startArrow, "none");
  if (stroke.endArrow !== undefined) set("endarrow", stroke.endArrow, "none");
  set("opacity", stroke.opacity === undefined ? undefined : String(stroke.opacity), "1");
  setChild(shape, "v:stroke", strokeEl.attrs.length ? strokeEl : undefined);
}

/** The shape's outline (`null` = none). VML defaults: black, ¾ pt. */
export function getShapeStroke(shape: XmlElement): ShapeStroke | null {
  const stroked = getAttr(shape, "stroked");
  const strokeEl = child(shape, "v:stroke");
  if (stroked === "f" || stroked === "false" || (strokeEl && getAttr(strokeEl, "on") === "f"))
    return null;
  const read = (name: string): string | undefined =>
    strokeEl ? getAttr(strokeEl, name) : undefined;
  const opacity = readFraction(read("opacity"));
  return {
    color: readColor(getAttr(shape, "strokecolor") ?? read("color")) ?? "000000",
    weight: lengthToPoints(getAttr(shape, "strokeweight") ?? read("weight")) ?? 0.75,
    dash: (read("dashstyle") as ShapeDash | undefined) ?? "solid",
    startArrow: (read("startarrow") as ShapeArrow | undefined) ?? "none",
    endArrow: (read("endarrow") as ShapeArrow | undefined) ?? "none",
    ...(opacity === undefined ? {} : { opacity }),
  };
}

/** Shape Effects ▸ Shadow; `null` = No Shadow. */
export function setShapeShadow(shape: XmlElement, shadow: ShapeShadow | null): void {
  setChild(
    shape,
    "v:shadow",
    shadow
      ? element("v:shadow", {
          on: "t",
          color: color(shadow.color),
          offset: `${points(shadow.offsetX)},${points(shadow.offsetY)}`,
          opacity: shadow.opacity === undefined ? undefined : String(shadow.opacity),
        })
      : undefined,
  );
}

export function getShapeShadow(shape: XmlElement): ShapeShadow | null {
  const el = child(shape, "v:shadow");
  if (!el || !["t", "true"].includes(getAttr(el, "on") ?? "")) return null;
  // VML's default shadow offset is 2pt down and right.
  const [x = "2pt", y = "2pt"] = (getAttr(el, "offset") ?? "2pt,2pt").split(",");
  const opacity = readFraction(getAttr(el, "opacity"));
  return {
    color: readColor(getAttr(el, "color")) ?? "808080",
    offsetX: lengthToPoints(x) ?? 2,
    offsetY: lengthToPoints(y) ?? 2,
    ...(opacity === undefined ? {} : { opacity }),
  };
}

// ---------------------------------------------------------------------------
// Text
// ---------------------------------------------------------------------------

function txbxContent(shape: XmlElement): XmlElement | undefined {
  const box = child(shape, "v:textbox");
  return box && child(box, "w:txbxContent");
}

/** The paragraphs of a shape's text box (`v:textbox/w:txbxContent`); empty when it has none. */
export function shapeText(shape: XmlElement): WmlParagraph[] {
  const content = txbxContent(shape);
  if (!content) return [];
  return children(content)
    .filter((c) => isElement(c, "w:p"))
    .map(parseParagraph);
}

/**
 * Replace the shape's text. A string becomes one paragraph per line. Adds a
 * text box to the shape if it has none (Add Text on a shape).
 */
export function setShapeText(shape: XmlElement, content: string | readonly WmlParagraph[]): void {
  const paragraphs =
    typeof content === "string" ? content.split(/\r?\n/).map(buildTextParagraph) : content;
  const box = child(shape, "v:textbox") ?? element("v:textbox");
  const others = box.children.filter((c) => !isElement(c, "w:txbxContent"));
  const body = element("w:txbxContent", {}, paragraphs.map(paragraphToElement));
  (box.children as XmlNode[]).splice(0, box.children.length, ...others, body);
  (box as { selfClosing: boolean }).selfClosing = false;
  setChild(shape, "v:textbox", box);
}

/** Text Direction, Align Text, internal margins and Resize shape to fit text. */
export function getTextBoxLayout(shape: XmlElement): TextBoxLayout {
  const box = child(shape, "v:textbox");
  const boxStyle = parseVmlStyle(box && getAttr(box, "style"));
  const flow = boxStyle.get("layout-flow");
  const alt = boxStyle.get("mso-layout-flow-alt");
  const anchor = readStyle(shape).get("v-text-anchor") ?? "top";
  const inset = (box && getAttr(box, "inset"))?.split(",") ?? [];
  const margin = (i: number): number =>
    lengthToPoints(inset[i]?.trim() || undefined) ?? DEFAULT_INSET[i] ?? 0;
  return {
    direction:
      flow === "vertical" ? (alt === "bottom-to-top" ? "vertical270" : "vertical") : "horizontal",
    anchor: anchor.startsWith("middle") ? "middle" : anchor.startsWith("bottom") ? "bottom" : "top",
    inset: [margin(0), margin(1), margin(2), margin(3)],
    autoFit: boxStyle.get("mso-fit-shape-to-text") === "t",
  };
}

export function setTextBoxLayout(shape: XmlElement, layout: Partial<TextBoxLayout>): void {
  const box = child(shape, "v:textbox") ?? element("v:textbox");
  const boxStyle = parseVmlStyle(getAttr(box, "style"));
  if (layout.direction !== undefined) {
    boxStyle.delete("layout-flow");
    boxStyle.delete("mso-layout-flow-alt");
    if (layout.direction !== "horizontal") boxStyle.set("layout-flow", "vertical");
    if (layout.direction === "vertical270") boxStyle.set("mso-layout-flow-alt", "bottom-to-top");
  }
  if (layout.autoFit !== undefined) {
    if (layout.autoFit) boxStyle.set("mso-fit-shape-to-text", "t");
    else boxStyle.delete("mso-fit-shape-to-text");
  }
  setAttr(box, "style", boxStyle.size ? serializeVmlStyle(boxStyle) : undefined);
  if (layout.inset !== undefined) setAttr(box, "inset", layout.inset.map(points).join(","));
  setChild(shape, "v:textbox", box);
  if (layout.anchor !== undefined) {
    const style = readStyle(shape);
    if (layout.anchor === "top") style.delete("v-text-anchor");
    else style.set("v-text-anchor", layout.anchor);
    setAttr(shape, "style", serializeVmlStyle(style));
  }
}

/**
 * Text ▸ Create Link: overflow text continues in `next` (`mso-next-textbox`
 * on the text box). `undefined` breaks the link.
 */
export function linkTextBoxes(from: XmlElement, next: XmlElement | undefined): void {
  const box = child(from, "v:textbox") ?? element("v:textbox");
  const style = parseVmlStyle(getAttr(box, "style"));
  if (next) {
    const id = getAttr(next, "id");
    if (!id) throw new Error("The target text box has no id");
    style.set("mso-next-textbox", `#${id}`);
  } else {
    style.delete("mso-next-textbox");
  }
  setAttr(box, "style", style.size ? serializeVmlStyle(style) : undefined);
  setChild(from, "v:textbox", box);
}

/** The id of the text box `shape` flows into, if linked. */
export function linkedTextBox(shape: XmlElement): string | undefined {
  const box = child(shape, "v:textbox");
  return parseVmlStyle(box && getAttr(box, "style"))
    .get("mso-next-textbox")
    ?.replace(/^#/, "");
}

/** The WordArt text and font, or `undefined` for a shape that is not WordArt. */
export function getWordArt(shape: XmlElement): WordArtText | undefined {
  const path = child(shape, "v:textpath");
  if (!path) return undefined;
  const style = parseVmlStyle(getAttr(path, "style"));
  return {
    text: getAttr(path, "string") ?? "",
    font: (style.get("font-family") ?? DEFAULT_FONT).replace(/^["']|["']$/g, ""),
    size: lengthToPoints(style.get("font-size")) ?? DEFAULT_WORDART_SIZE,
    bold: style.get("font-weight") === "bold",
    italic: style.get("font-style") === "italic",
  };
}

export function setWordArt(shape: XmlElement, text: Partial<WordArtText>): void {
  const path = child(shape, "v:textpath");
  if (!path) throw new Error("The shape is not WordArt");
  if (text.text !== undefined) setAttr(path, "string", text.text);
  const style = parseVmlStyle(getAttr(path, "style"));
  if (text.font !== undefined) style.set("font-family", `"${text.font.replace(/["\\]/g, "")}"`);
  if (text.size !== undefined) style.set("font-size", points(text.size));
  if (text.bold !== undefined) {
    if (text.bold) style.set("font-weight", "bold");
    else style.delete("font-weight");
  }
  if (text.italic !== undefined) {
    if (text.italic) style.set("font-style", "italic");
    else style.delete("font-style");
  }
  style.set("v-text-kern", "t");
  setAttr(path, "style", serializeVmlStyle(style));
}

// ---------------------------------------------------------------------------
// Alt text, preset, points
// ---------------------------------------------------------------------------

/** Alt Text (the VML `alt` attribute). */
export function setShapeAltText(shape: XmlElement, alt: string): void {
  setAttr(shape, "alt", alt === "" ? undefined : alt);
}

export function getShapeAltText(shape: XmlElement): string {
  return getAttr(shape, "alt") ?? "";
}

/** The shape's name (its VML `id`), as the Selection Pane shows it. */
export function getShapeName(shape: XmlElement): string {
  return getAttr(shape, "id") ?? "";
}

export function setShapeName(shape: XmlElement, name: string): void {
  setAttr(shape, "id", name);
}

/** Edit Shape ▸ Change Shape: swap the geometry, keeping box, fill, outline and text. */
export function changeShapePreset(shape: XmlElement, preset: ShapePreset): void {
  const def: ShapePresetDef = SHAPE_PRESETS[preset];
  setAttr(shape, "type", undefined);
  rename(shape, `v:${def.element}`);
  applyPresetGeometry(shape, def);
}

/**
 * Vertices of a freeform built from straight segments (Edit Points), in points
 * from the box's top-left corner. `undefined` when the path has curves.
 */
export function getShapePoints(shape: XmlElement): Array<[number, number]> | undefined {
  const path = getAttr(shape, "path");
  if (!path || /[^mlxe0-9,\s-]/i.test(path)) return undefined;
  const [cw = COORD_SIZE, ch = COORD_SIZE] = (getAttr(shape, "coordsize") ?? "")
    .split(",")
    .map(Number);
  const nums = path
    .replace(/[mlxe]/gi, " ")
    .split(/[\s,]+/)
    .filter(Boolean)
    .map(Number);
  const { width, height } = getShapeLayout(shape);
  const out: Array<[number, number]> = [];
  for (let i = 0; i + 1 < nums.length; i += 2) {
    out.push([((nums[i] ?? 0) / cw) * width, ((nums[i + 1] ?? 0) / ch) * height]);
  }
  return out;
}

/** Replace a freeform's vertices (Edit Points). */
export function setShapePoints(
  shape: XmlElement,
  pts: ReadonlyArray<readonly [number, number]>,
): void {
  const { width, height } = getShapeLayout(shape);
  const closed = /x/i.test(getAttr(shape, "path") ?? "");
  const path = pathFromPoints(pts, width, height);
  setAttr(shape, "coordsize", `${COORD_SIZE},${COORD_SIZE}`);
  setAttr(shape, "path", closed ? path.replace(/e$/, "xe") : path);
  setAttr(shape, "o:spt", undefined);
}

// ---------------------------------------------------------------------------
// Delete, group, ungroup
// ---------------------------------------------------------------------------

/** Delete a top-level shape (and its run, when the run held nothing else). */
export function removeShape(doc: Docx, shape: XmlElement): boolean {
  const loc = locate(doc, shape);
  if (!loc) return false;
  detach(loc, shape);
  doc.dirty = true;
  return true;
}

function detach(loc: PictLocation, shape: XmlElement): void {
  const list = loc.pict.children as XmlNode[];
  list.splice(list.indexOf(shape), 1);
  if (drawablesOf(loc.pict).length > 0) return;
  loc.run.pieces = loc.run.pieces.filter((p) => !(p.kind === "pict" && p.node === loc.pict));
  if (loc.run.pieces.length === 0) {
    const idx = loc.paragraph.children.indexOf(loc.run);
    if (idx >= 0) loc.paragraph.children.splice(idx, 1);
  }
}

/**
 * Arrange ▸ Group: combine floating shapes into one `v:group`, which takes the
 * first shape's place, anchoring and wrap. Child coordinates are kept in
 * twentieths of a point relative to the group's box.
 */
export function groupShapes(doc: Docx, shapes: readonly XmlElement[]): XmlElement {
  if (shapes.length < 2) throw new Error("Grouping needs at least two shapes");
  const locs = shapes.map((s) => {
    const loc = locate(doc, s);
    if (!loc) throw new Error("Only top-level shapes can be grouped");
    if (getShapeLayout(s).inline) throw new Error("Inline shapes cannot be grouped");
    return loc;
  });
  const layouts = shapes.map(getShapeLayout);
  const left = Math.min(...layouts.map((l) => l.left));
  const top = Math.min(...layouts.map((l) => l.top));
  const right = Math.max(...layouts.map((l) => l.left + l.width));
  const bottom = Math.max(...layouts.map((l) => l.top + l.height));
  const first = shapes[0];
  const firstLoc = locs[0];
  const firstLayout = layouts[0];
  if (!first || !firstLoc || !firstLayout) throw new Error("Grouping needs at least two shapes");
  const u = GROUP_UNITS_PER_POINT;
  const id = nextShapeId(doc);
  const groupStyle = new Map<string, string>([
    ["position", "absolute"],
    ["margin-left", points(left)],
    ["margin-top", points(top)],
    ["width", points(right - left)],
    ["height", points(bottom - top)],
    ["z-index", String(Math.max(...layouts.map((l) => l.zIndex)))],
    ["mso-position-horizontal-relative", RELATIVE_H_TO_VML[firstLayout.horizontalRelativeTo]],
    ["mso-position-vertical-relative", RELATIVE_V_TO_VML[firstLayout.verticalRelativeTo]],
  ]);
  const wrap = child(first, "w10:wrap");
  const group = element(
    "v:group",
    {
      id: `Group ${id - FIRST_SHAPE_ID + 1}`,
      "o:spid": `_x0000_s${id}`,
      style: serializeVmlStyle(groupStyle),
      coordorigin: `${Math.round(left * u)},${Math.round(top * u)}`,
      coordsize: `${Math.round((right - left) * u)},${Math.round((bottom - top) * u)}`,
    },
    wrap ? [wrap] : [],
  );
  shapes.forEach((shape, i) => {
    const l = layouts[i];
    const loc = locs[i];
    if (!l || !loc) return;
    detach(loc, shape);
    setChild(shape, "w10:wrap", undefined);
    const style = readStyle(shape);
    for (const k of [
      "margin-left",
      "margin-top",
      "z-index",
      "mso-position-horizontal",
      "mso-position-horizontal-relative",
      "mso-position-vertical",
      "mso-position-vertical-relative",
    ])
      style.delete(k);
    style.set("position", "absolute");
    style.set("left", String(Math.round(l.left * u)));
    style.set("top", String(Math.round(l.top * u)));
    style.set("width", String(Math.round(l.width * u)));
    style.set("height", String(Math.round(l.height * u)));
    setAttr(shape, "style", serializeVmlStyle(style));
    (group.children as XmlNode[]).push(shape);
  });
  (group as { selfClosing: boolean }).selfClosing = false;
  insertRun(doc, firstLoc.paragraph, group, Math.max(0, firstLoc.paragraph.children.length));
  return group;
}

/** Arrange ▸ Ungroup: the group's shapes become top-level shapes in its paragraph. */
export function ungroupShapes(doc: Docx, group: XmlElement): XmlElement[] {
  if (!isElement(group, "v:group")) throw new Error("Not a group");
  const loc = locate(doc, group);
  if (!loc) throw new Error("Only top-level groups can be ungrouped");
  const gl = getShapeLayout(group);
  const frame = groupFrame(group);
  const wrap = child(group, "w10:wrap");
  const members = children(group).filter(isDrawable);
  const at = loc.paragraph.children.indexOf(loc.run) + 1;
  detach(loc, group);
  members.forEach((shape, i) => {
    const l = getShapeLayout(shape);
    const style = readStyle(shape);
    style.delete("left");
    style.delete("top");
    style.set("position", "absolute");
    style.set("margin-left", points(gl.left + (l.left - frame.originX) / frame.scaleX));
    style.set("margin-top", points(gl.top + (l.top - frame.originY) / frame.scaleY));
    style.set("width", points(l.width / frame.scaleX));
    style.set("height", points(l.height / frame.scaleY));
    style.set("z-index", String(gl.zIndex));
    style.set("mso-position-horizontal-relative", RELATIVE_H_TO_VML[gl.horizontalRelativeTo]);
    style.set("mso-position-vertical-relative", RELATIVE_V_TO_VML[gl.verticalRelativeTo]);
    setAttr(shape, "style", serializeVmlStyle(style));
    if (wrap) setChild(shape, "w10:wrap", element("w10:wrap", { type: getAttr(wrap, "type") }));
    insertRun(doc, loc.paragraph, shape, Math.min(at + i, loc.paragraph.children.length));
  });
  return members;
}

/** The members of a group or drawing canvas. */
export function groupMembers(group: XmlElement): XmlElement[] {
  return children(group).filter(isDrawable);
}
