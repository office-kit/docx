/**
 * Placement of a `<w:drawing>`: inline (`wp:inline`) versus floating
 * (`wp:anchor`), text wrapping, position, z-order and distance from text
 * (ECMA-376 Part 1, §20.4.2).
 */

import type { XmlElement, XmlNode } from "../xml/index.js";
import type {
  DrawingAnchor,
  HorizontalPosition,
  VerticalPosition,
  WrapDistance,
  WrapPoint,
  WrapSide,
  WrapStyle,
} from "./types.js";
import {
  attr,
  boolAttr,
  child,
  elementChildren,
  ensureDrawingNamespaces,
  fragment,
  numAttr,
  orderAttrs,
  placeChild,
  setAttr,
  WP_NS,
} from "./xml.js";

/** CT_Anchor's child sequence; the wrap elements form one choice. */
export const ANCHOR_ORDER = [
  "simplePos",
  "positionH",
  "positionV",
  "extent",
  "effectExtent",
  "wrapNone|wrapSquare|wrapTight|wrapThrough|wrapTopAndBottom",
  "docPr",
  "cNvGraphicFramePr",
  "graphic",
] as const;
const ANCHOR_ATTR_ORDER = [
  "distT",
  "distB",
  "distL",
  "distR",
  "simplePos",
  "relativeHeight",
  "behindDoc",
  "locked",
  "layoutInCell",
  "hidden",
  "allowOverlap",
];
const INLINE_ORDER = ["extent", "effectExtent", "docPr", "cNvGraphicFramePr", "graphic"] as const;

// Word's default distance from text for a newly floated picture: 0.13" left
// and right (114300 EMU), none above or below.
export const DEFAULT_WRAP_DISTANCE: WrapDistance = {
  top: 0,
  bottom: 0,
  left: 114300,
  right: 114300,
};
// wrapPolygon coordinates live in a 21600 × 21600 square mapped onto the object.
export const WRAP_POLYGON_UNITS = 21600;
// Word numbers floating objects from this relativeHeight upward.
export const BASE_RELATIVE_HEIGHT = 251659264;

const WRAP_ELEMENT: Readonly<Record<Exclude<WrapStyle, "inline">, string>> = {
  square: "wrapSquare",
  tight: "wrapTight",
  through: "wrapThrough",
  topAndBottom: "wrapTopAndBottom",
  behindText: "wrapNone",
  inFrontOfText: "wrapNone",
};

/** The `wp:inline` or `wp:anchor` directly under `<w:drawing>`. */
export function placementOf(drawing: XmlElement): XmlElement | undefined {
  return child(drawing, WP_NS, "anchor") ?? child(drawing, WP_NS, "inline");
}

function wrapElementOf(anchor: XmlElement): XmlElement | undefined {
  return elementChildren(anchor).find(
    (c) => c.name.uri === WP_NS && c.name.local.startsWith("wrap"),
  );
}

export function readWrap(drawing: XmlElement): WrapStyle {
  const placement = placementOf(drawing);
  if (!placement || placement.name.local === "inline") return "inline";
  switch (wrapElementOf(placement)?.name.local) {
    case "wrapSquare":
      return "square";
    case "wrapTight":
      return "tight";
    case "wrapThrough":
      return "through";
    case "wrapTopAndBottom":
      return "topAndBottom";
    default:
      return boolAttr(placement, "behindDoc") ? "behindText" : "inFrontOfText";
  }
}

function readPosition<P extends HorizontalPosition | VerticalPosition>(
  el: XmlElement | undefined,
  fallbackFrame: P["relativeTo"],
): P {
  const relativeTo = (attr(el, "relativeFrom") ?? fallbackFrame) as P["relativeTo"];
  const align = child(el, WP_NS, "align");
  if (align) return { relativeTo, align: textOf(align) } as P;
  const offset = child(el, WP_NS, "posOffset");
  return { relativeTo, offsetEmu: offset ? Number(textOf(offset)) || 0 : 0 } as P;
}

function textOf(el: XmlElement): string {
  return el.children
    .map((c) => (c.kind === "text" ? c.value : ""))
    .join("")
    .trim();
}

export function readAnchor(drawing: XmlElement): DrawingAnchor | undefined {
  const anchor = child(drawing, WP_NS, "anchor");
  if (!anchor) return undefined;
  const wrap = wrapElementOf(anchor);
  const polygon = child(wrap, WP_NS, "wrapPolygon");
  const points = polygon
    ? elementChildren(polygon).map((p) => ({ x: numAttr(p, "x") ?? 0, y: numAttr(p, "y") ?? 0 }))
    : undefined;
  return {
    horizontal: readPosition<HorizontalPosition>(child(anchor, WP_NS, "positionH"), "column"),
    vertical: readPosition<VerticalPosition>(child(anchor, WP_NS, "positionV"), "paragraph"),
    wrapSide: (attr(wrap, "wrapText") ?? "bothSides") as WrapSide,
    ...(points ? { wrapPolygon: points } : {}),
    distance: {
      top: numAttr(anchor, "distT") ?? 0,
      bottom: numAttr(anchor, "distB") ?? 0,
      left: numAttr(anchor, "distL") ?? 0,
      right: numAttr(anchor, "distR") ?? 0,
    },
    allowOverlap: boolAttr(anchor, "allowOverlap") ?? true,
    locked: boolAttr(anchor, "locked") ?? false,
    layoutInCell: boolAttr(anchor, "layoutInCell") ?? true,
    relativeHeight: numAttr(anchor, "relativeHeight") ?? 0,
  };
}

function positionMarkup(
  local: "positionH" | "positionV",
  pos: HorizontalPosition | VerticalPosition,
): string {
  const inner =
    "align" in pos
      ? `<wp:align>${pos.align}</wp:align>`
      : `<wp:posOffset>${Math.round(pos.offsetEmu)}</wp:posOffset>`;
  return `<wp:${local} relativeFrom="${pos.relativeTo}">${inner}</wp:${local}>`;
}

function pointMarkup(local: string, pt: WrapPoint): string {
  return `<wp:${local} x="${Math.round(pt.x)}" y="${Math.round(pt.y)}"/>`;
}

function polygonMarkup(points: readonly WrapPoint[], edited: boolean): string {
  const [first, ...rest] = points;
  if (!first || rest.length < 2) throw new Error("A wrap polygon needs at least three points.");
  return `<wp:wrapPolygon edited="${edited ? 1 : 0}">${pointMarkup("start", first)}${rest.map((pt) => pointMarkup("lineTo", pt)).join("")}</wp:wrapPolygon>`;
}

/** The object's bounding box as a closed wrap polygon (Word's default for Tight / Through). */
export const RECT_POLYGON: readonly WrapPoint[] = [
  { x: 0, y: 0 },
  { x: 0, y: WRAP_POLYGON_UNITS },
  { x: WRAP_POLYGON_UNITS, y: WRAP_POLYGON_UNITS },
  { x: WRAP_POLYGON_UNITS, y: 0 },
  { x: 0, y: 0 },
];

function wrapMarkup(
  wrap: Exclude<WrapStyle, "inline">,
  side: WrapSide,
  polygon: readonly WrapPoint[] | undefined,
): string {
  const local = WRAP_ELEMENT[wrap];
  switch (local) {
    case "wrapSquare":
      return `<wp:wrapSquare wrapText="${side}"/>`;
    case "wrapTight":
    case "wrapThrough":
      return `<wp:${local} wrapText="${side}">${polygonMarkup(polygon ?? RECT_POLYGON, polygon !== undefined)}</wp:${local}>`;
    default:
      return `<wp:${local}/>`;
  }
}

function replaceElement(parent: XmlElement, old: XmlElement, next: XmlElement): void {
  const kids = parent.children as XmlNode[];
  kids[kids.indexOf(old)] = next;
}

export interface FloatOptions {
  readonly horizontal: HorizontalPosition;
  readonly vertical: VerticalPosition;
  readonly relativeHeight: number;
}

/**
 * Change the wrapping style. Going from inline to floating builds a
 * `wp:anchor` around the drawing's extent, docPr and graphic at
 * `float.horizontal` / `float.vertical`; going back to inline drops the
 * anchor-only parts. Between floating styles only the wrap element and
 * `behindDoc` change, so the position is kept.
 */
export function applyWrap(drawing: XmlElement, wrap: WrapStyle, float: FloatOptions): void {
  ensureDrawingNamespaces(drawing);
  const placement = placementOf(drawing);
  if (!placement) throw new Error("The drawing has neither wp:inline nor wp:anchor.");
  // By local name: the sequence mixes namespaces (wp:extent … a:graphic).
  const keep = (local: string): XmlElement | undefined =>
    elementChildren(placement).find((c) => c.name.local === local);
  if (wrap === "inline") {
    if (placement.name.local === "inline") return;
    const inline = fragment(`<wp:inline distT="0" distB="0" distL="0" distR="0"/>`);
    for (const local of INLINE_ORDER) {
      const el = keep(local);
      if (el) (inline.children as XmlNode[]).push(el);
    }
    replaceElement(drawing, placement, inline);
    return;
  }
  let anchor = placement;
  if (placement.name.local === "inline") {
    const d = DEFAULT_WRAP_DISTANCE;
    anchor = fragment(
      `<wp:anchor distT="${d.top}" distB="${d.bottom}" distL="${d.left}" distR="${d.right}" simplePos="0" relativeHeight="${float.relativeHeight}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">` +
        `<wp:simplePos x="0" y="0"/>${positionMarkup("positionH", float.horizontal)}${positionMarkup("positionV", float.vertical)}</wp:anchor>`,
    );
    for (const local of INLINE_ORDER) {
      const el = keep(local);
      if (el) (anchor.children as XmlNode[]).push(el);
    }
    replaceElement(drawing, placement, anchor);
  }
  const current = wrapElementOf(anchor);
  const side = (attr(current, "wrapText") ?? "bothSides") as WrapSide;
  const existingPolygon = readAnchor(drawing)?.wrapPolygon;
  placeChild(anchor, fragment(wrapMarkup(wrap, side, existingPolygon)), ANCHOR_ORDER);
  setAttr(anchor, "behindDoc", wrap === "behindText" ? "1" : "0");
  orderAttrs(anchor, ANCHOR_ATTR_ORDER);
}

/** The largest `wp:docPr` id among `drawings` (ids must be unique per document). */
export function maxDocPrId(drawings: readonly XmlElement[]): number {
  let max = 0;
  for (const drawing of drawings) {
    const id = numAttr(child(placementOf(drawing), WP_NS, "docPr"), "id");
    if (id !== undefined && id > max) max = id;
  }
  return max;
}

function requireAnchor(drawing: XmlElement): XmlElement {
  const anchor = child(drawing, WP_NS, "anchor");
  if (!anchor)
    throw new Error("The object is in line with text; give it a text wrapping style first.");
  return anchor;
}

export function applyPosition(
  drawing: XmlElement,
  pos: { readonly horizontal?: HorizontalPosition; readonly vertical?: VerticalPosition },
): void {
  ensureDrawingNamespaces(drawing);
  const anchor = requireAnchor(drawing);
  if (pos.horizontal)
    placeChild(anchor, fragment(positionMarkup("positionH", pos.horizontal)), ANCHOR_ORDER);
  if (pos.vertical)
    placeChild(anchor, fragment(positionMarkup("positionV", pos.vertical)), ANCHOR_ORDER);
}

export interface AnchorOptions {
  readonly wrapSide?: WrapSide;
  /** `null` restores the default (bounding box) polygon. */
  readonly wrapPolygon?: readonly WrapPoint[] | null;
  readonly distance?: Partial<WrapDistance>;
  readonly allowOverlap?: boolean;
  readonly locked?: boolean;
  readonly layoutInCell?: boolean;
  readonly relativeHeight?: number;
}

const DISTANCE_ATTR: Readonly<Record<keyof WrapDistance, string>> = {
  top: "distT",
  bottom: "distB",
  left: "distL",
  right: "distR",
};

export function applyAnchorOptions(drawing: XmlElement, options: AnchorOptions): void {
  ensureDrawingNamespaces(drawing);
  const anchor = requireAnchor(drawing);
  const flag = (local: string, value: boolean | undefined): void => {
    if (value !== undefined) setAttr(anchor, local, value ? "1" : "0");
  };
  flag("allowOverlap", options.allowOverlap);
  flag("locked", options.locked);
  flag("layoutInCell", options.layoutInCell);
  if (options.relativeHeight !== undefined) {
    setAttr(anchor, "relativeHeight", String(Math.max(0, Math.round(options.relativeHeight))));
  }
  for (const [side, value] of Object.entries(options.distance ?? {})) {
    if (value === undefined) continue;
    if (!Number.isFinite(value) || value < 0)
      throw new Error(`Invalid distance from text: ${value}.`);
    setAttr(anchor, DISTANCE_ATTR[side as keyof WrapDistance], String(Math.round(value)));
  }
  const wrap = wrapElementOf(anchor);
  if (
    wrap &&
    options.wrapSide !== undefined &&
    wrap.name.local !== "wrapNone" &&
    wrap.name.local !== "wrapTopAndBottom"
  ) {
    setAttr(wrap, "wrapText", options.wrapSide);
  }
  if (options.wrapPolygon !== undefined) {
    const style = readWrap(drawing);
    if (style !== "tight" && style !== "through") {
      throw new Error("Wrap points apply to Tight and Through wrapping only.");
    }
    const side = (attr(wrap, "wrapText") ?? "bothSides") as WrapSide;
    const polygon = options.wrapPolygon ?? undefined;
    placeChild(anchor, fragment(wrapMarkup(style, side, polygon)), ANCHOR_ORDER);
  }
  orderAttrs(anchor, ANCHOR_ATTR_ORDER);
}
