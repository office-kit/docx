/**
 * VML → HTML/SVG for the canvas: shapes, text boxes, WordArt, ink, groups and
 * pictures inside `w:pict` (ECMA-376 Part 4, §19.1), including Word-authored
 * VML that references shape types and their formulas.
 *
 * Floating objects (`position:absolute` in their style) are positioned against
 * the nearest positioned ancestor, which must be the page box and carry the
 * page CSS variables (`--page-w`, `--page-h`, `--m-top`, `--m-right`,
 * `--m-bottom`, `--m-left`). Offsets from the paragraph use the object's static
 * position, so the renderer emits floats at the start of their paragraph.
 *
 * Text wrapping is approximated with CSS: Square / Tight / Through become a
 * float on the shape's side of the column, Top and Bottom a block, and the
 * rest (and anything positioned against the page vertically) an overlay.
 */

import {
  type Docx,
  getShapeFill,
  getShapeLayout,
  getShapeShadow,
  getShapeStroke,
  getShapeWrap,
  getTextBoxLayout,
  getWordArt,
  groupMembers,
  type ShapeFillInfo,
  type ShapeLayout,
  type ShapeStroke,
  shapeKind,
  shapeText,
  vmlImageData,
  type WmlParagraph,
  type WmlRun,
  type XmlElement,
  type XmlNode,
  FILL_PATTERNS,
  type FillPattern,
} from "@office-kit/docx";
import { createStyleResolver, type StyleResolver } from "./resolve.js";
import { paragraphCss, runCss } from "./render.js";
import type { DocPosition } from "./selection.js";
import { renderRunSmartArt } from "./render-smartart.js";
import { vmlPathToSvg } from "./vml-path.js";

export const VML_NS = "urn:schemas-microsoft-com:vml";
const OFFICE_NS = "urn:schemas-microsoft-com:office:office";
const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

export interface VmlRenderContext {
  readonly doc: Docx;
  /** The part holding the VML (for `r:id` image lookups); the main document by default. */
  readonly partName?: string;
  /** Where the object's run is; omit for objects that cannot be selected (headers). */
  readonly at?: DocPosition;
}

const DRAWABLE: ReadonlySet<string> = new Set([
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
const DEFAULT_COORD = 1000;
const ROUNDRECT_DEFAULT_ARC = 0.2;
// A WordArt string fills its box; its glyphs are scaled to the box height.
const WORDART_EM = 0.82;
const POINTS_PER_PX = 0.75;

const escapeHtml = (s: string): string =>
  s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

function attr(el: XmlElement, local: string, uri = ""): string | undefined {
  return el.attrs.find((a) => !a.isNamespaceDecl && a.name.local === local && a.name.uri === uri)
    ?.value;
}

function vmlChild(el: XmlElement, local: string): XmlElement | undefined {
  return el.children.find(
    (c): c is XmlElement => c.kind === "element" && c.name.uri === VML_NS && c.name.local === local,
  );
}

function isDrawable(node: XmlNode): node is XmlElement {
  return node.kind === "element" && node.name.uri === VML_NS && DRAWABLE.has(node.name.local);
}

// ---------------------------------------------------------------------------
// Shape types
// ---------------------------------------------------------------------------

// Shape types are defined once per part and referenced by `type="#id"` from
// any later shape, so they are indexed per document body.
const shapeTypeCache = new WeakMap<object, Map<string, XmlElement>>();

function collectShapeTypes(node: XmlElement, into: Map<string, XmlElement>): void {
  for (const c of node.children) {
    if (c.kind !== "element") continue;
    if (c.name.uri === VML_NS && c.name.local === "shapetype") {
      const id = attr(c, "id");
      if (id) into.set(id, c);
    } else collectShapeTypes(c, into);
  }
}

function shapeTypesFor(doc: Docx, pict: XmlElement): Map<string, XmlElement> {
  let map = shapeTypeCache.get(doc.document);
  if (!map) {
    map = new Map();
    for (const block of doc.document.body.blocks) {
      if (block.kind !== "paragraph") continue;
      for (const inline of block.children) {
        if (inline.kind !== "run") continue;
        for (const piece of inline.pieces)
          if (piece.kind === "pict") collectShapeTypes(piece.node, map);
      }
    }
    shapeTypeCache.set(doc.document, map);
  }
  // Header / footer VML (watermarks) carries its own shape types.
  const local = new Map<string, XmlElement>();
  collectShapeTypes(pict, local);
  return local.size ? new Map([...map, ...local]) : map;
}

/**
 * The shape with its shape type folded in: the type supplies defaults for any
 * attribute and child element the shape does not set itself.
 */
function effective(shape: XmlElement, types: Map<string, XmlElement>): XmlElement {
  const ref = attr(shape, "type")?.replace(/^#/, "");
  const type = ref ? types.get(ref) : undefined;
  if (!type) return shape;
  const own = new Set(shape.attrs.map((a) => `${a.name.uri}|${a.name.local}`));
  const ownChildren = new Set(
    shape.children
      .filter((c): c is XmlElement => c.kind === "element")
      .map((c) => `${c.name.uri}|${c.name.local}`),
  );
  return {
    ...shape,
    attrs: [
      ...shape.attrs,
      ...type.attrs.filter(
        (a) => a.name.local !== "id" && !own.has(`${a.name.uri}|${a.name.local}`),
      ),
    ],
    children: [
      ...shape.children,
      ...type.children.filter(
        (c) => c.kind === "element" && !ownChildren.has(`${c.name.uri}|${c.name.local}`),
      ),
    ],
  };
}

// ---------------------------------------------------------------------------
// Colors, fills, strokes
// ---------------------------------------------------------------------------

let idCounter = 0;
const nextId = (): string => `wkv${++idCounter}`;

const hex = (c: string): string => `#${c}`;

const DASHES: Readonly<Record<string, readonly number[]>> = {
  shortdash: [3, 1],
  shortdot: [1, 1],
  shortdashdot: [3, 1, 1, 1],
  shortdashdotdot: [3, 1, 1, 1, 1, 1],
  dot: [1, 3],
  dash: [4, 3],
  longdash: [8, 3],
  dashdot: [4, 3, 1, 3],
  longdashdot: [8, 3, 1, 3],
  longdashdotdot: [8, 3, 1, 3, 1, 3],
};

const base64Cache = new WeakMap<Uint8Array, string>();

function dataUrl(bytes: Uint8Array, contentType: string): string {
  let b64 = base64Cache.get(bytes);
  if (b64 === undefined) {
    let binary = "";
    const CHUNK = 0x8000;
    for (let i = 0; i < bytes.length; i += CHUNK) {
      binary += String.fromCharCode(...bytes.subarray(i, i + CHUNK));
    }
    b64 = btoa(binary);
    base64Cache.set(bytes, b64);
  }
  return `data:${contentType};base64,${b64}`;
}

/** SVG `<defs>` content and the paint reference for a fill. */
function fillPaint(
  fill: ShapeFillInfo,
  ctx: VmlRenderContext,
  width: number,
  height: number,
): { defs: string; paint: string; opacity?: number } {
  switch (fill.type) {
    case "none":
      return { defs: "", paint: "none" };
    case "solid":
      return {
        defs: "",
        paint: hex(fill.color),
        ...(fill.opacity === undefined ? {} : { opacity: fill.opacity }),
      };
    case "gradient": {
      const id = nextId();
      const stops = `<stop offset="0" stop-color="${hex(fill.color)}"/><stop offset="1" stop-color="${hex(fill.color2)}"/>`;
      if (fill.style === "radial") {
        return {
          defs: `<radialGradient id="${id}">${stops}</radialGradient>`,
          paint: `url(#${id})`,
          ...(fill.opacity === undefined ? {} : { opacity: fill.opacity }),
        };
      }
      // 0° runs top → bottom; the angle turns the gradient clockwise.
      const a = (((fill.angle ?? 0) + 90) * Math.PI) / 180;
      const dx = Math.cos(a) / 2;
      const dy = Math.sin(a) / 2;
      return {
        defs:
          `<linearGradient id="${id}" x1="${0.5 - dx}" y1="${0.5 - dy}" x2="${0.5 + dx}" y2="${0.5 + dy}">` +
          `${stops}</linearGradient>`,
        paint: `url(#${id})`,
        ...(fill.opacity === undefined ? {} : { opacity: fill.opacity }),
      };
    }
    case "pattern": {
      const id = nextId();
      return {
        defs: patternDef(id, fill.pattern, fill.color, fill.color2),
        paint: `url(#${id})`,
      };
    }
    case "picture": {
      const image = vmlImageData(ctx.doc, fill.relId, ctx.partName);
      if (!image) return { defs: "", paint: "#FFFFFF" };
      const id = nextId();
      const href = dataUrl(image.bytes, image.contentType);
      const tile = fill.tile;
      const w = tile ? 48 : width;
      const h = tile ? 48 : height;
      return {
        defs:
          `<pattern id="${id}" patternUnits="userSpaceOnUse" width="${w}" height="${h}">` +
          `<image href="${href}" width="${w}" height="${h}" preserveAspectRatio="none"/></pattern>`,
        paint: `url(#${id})`,
      };
    }
  }
}

export function patternDef(id: string, pattern: FillPattern, fg: string, bg: string): string {
  const rows = FILL_PATTERNS[pattern];
  let cells = "";
  rows.forEach((bits, y) => {
    for (let x = 0; x < 8; x++) {
      if (bits & (0x80 >> x)) cells += `<rect x="${x}" y="${y}" width="1" height="1"/>`;
    }
  });
  return (
    `<pattern id="${id}" patternUnits="userSpaceOnUse" width="6" height="6" viewBox="0 0 8 8">` +
    `<rect width="8" height="8" fill="${hex(bg)}"/><g fill="${hex(fg)}">${cells}</g></pattern>`
  );
}

function markerDef(id: string, kind: string, color: string, start: boolean): string {
  const shapes: Readonly<Record<string, string>> = {
    block: `<path d="M0,0 L10,5 L0,10 z"/>`,
    classic: `<path d="M0,0 L10,5 L0,10 L3,5 z"/>`,
    open: `<path d="M0,0 L10,5 L0,10" fill="none" stroke="${hex(color)}" stroke-width="1.5"/>`,
    oval: `<circle cx="5" cy="5" r="4"/>`,
    diamond: `<path d="M0,5 L5,0 L10,5 L5,10 z"/>`,
  };
  const body = shapes[kind];
  if (!body) return "";
  return (
    `<marker id="${id}" viewBox="0 0 10 10" refX="${kind === "oval" || kind === "diamond" ? 5 : 9}" refY="5" ` +
    `markerWidth="4" markerHeight="4" orient="${start ? "auto-start-reverse" : "auto"}" fill="${hex(color)}">${body}</marker>`
  );
}

function strokeAttrs(stroke: ShapeStroke | null): { attrs: string; defs: string } {
  if (!stroke) return { attrs: `stroke="none"`, defs: "" };
  const parts = [
    `stroke="${hex(stroke.color)}"`,
    `stroke-width="${stroke.weight}"`,
    `stroke-linejoin="round"`,
    `stroke-linecap="round"`,
  ];
  if (stroke.opacity !== undefined) parts.push(`stroke-opacity="${stroke.opacity}"`);
  const dash = DASHES[stroke.dash ?? "solid"];
  if (dash) parts.push(`stroke-dasharray="${dash.map((n) => n * stroke.weight).join(" ")}"`);
  let defs = "";
  if (stroke.startArrow && stroke.startArrow !== "none") {
    const id = nextId();
    defs += markerDef(id, stroke.startArrow, stroke.color, true);
    parts.push(`marker-start="url(#${id})"`);
  }
  if (stroke.endArrow && stroke.endArrow !== "none") {
    const id = nextId();
    defs += markerDef(id, stroke.endArrow, stroke.color, false);
    parts.push(`marker-end="url(#${id})"`);
  }
  return { attrs: parts.join(" "), defs };
}

// ---------------------------------------------------------------------------
// Geometry
// ---------------------------------------------------------------------------

function pair(value: string | undefined, fallback: readonly [number, number]): [number, number] {
  if (!value) return [fallback[0], fallback[1]];
  const [a, b] = value.split(",").map((s) => Number.parseFloat(s));
  return [Number.isFinite(a) ? (a ?? 0) : fallback[0], Number.isFinite(b) ? (b ?? 0) : fallback[1]];
}

/** A length attribute (`from="10pt,20pt"`) in points. */
function lengthPair(value: string | undefined): [number, number] {
  const out: number[] = (value ?? "0,0").split(",").map((v) => {
    const m = /^(-?[0-9.]+)\s*(pt|px|in|cm|mm)?$/.exec(v.trim());
    if (!m) return 0;
    const n = Number(m[1]);
    const unit = m[2] ?? "px";
    const factor: Record<string, number> = {
      pt: 1,
      px: POINTS_PER_PX,
      in: 72,
      cm: 72 / 2.54,
      mm: 72 / 25.4,
    };
    return n * (factor[unit] ?? 1);
  });
  return [out[0] ?? 0, out[1] ?? 0];
}

/** SVG path data (in points on the shape's box) for a VML element. */
function geometryPaths(
  shape: XmlElement,
  width: number,
  height: number,
): Array<{ d: string; fill: boolean; stroke: boolean }> {
  const local = shape.name.local;
  if (local === "rect" || local === "image") {
    return [{ d: `M0 0 H${width} V${height} H0 Z`, fill: true, stroke: true }];
  }
  if (local === "oval") {
    const rx = width / 2;
    const ry = height / 2;
    return [
      {
        d: `M0 ${ry} A${rx} ${ry} 0 1 0 ${width} ${ry} A${rx} ${ry} 0 1 0 0 ${ry} Z`,
        fill: true,
        stroke: true,
      },
    ];
  }
  if (local === "roundrect") {
    const raw = attr(shape, "arcsize");
    const fraction =
      raw === undefined
        ? ROUNDRECT_DEFAULT_ARC
        : raw.endsWith("f")
          ? Number(raw.slice(0, -1)) / 65536
          : Number(raw);
    const r =
      (Math.min(width, height) / 2) *
      (Number.isFinite(fraction) ? fraction : ROUNDRECT_DEFAULT_ARC);
    return [
      {
        d:
          `M${r} 0 H${width - r} A${r} ${r} 0 0 1 ${width} ${r} V${height - r} A${r} ${r} 0 0 1 ${width - r} ${height} ` +
          `H${r} A${r} ${r} 0 0 1 0 ${height - r} V${r} A${r} ${r} 0 0 1 ${r} 0 Z`,
        fill: true,
        stroke: true,
      },
    ];
  }
  if (local === "line") {
    const [x1, y1] = lengthPair(attr(shape, "from"));
    const [x2, y2] = lengthPair(attr(shape, "to"));
    return [{ d: `M${x1} ${y1} L${x2} ${y2}`, fill: false, stroke: true }];
  }
  if (local === "polyline") {
    const nums = (attr(shape, "points") ?? "").split(/[\s,]+/).filter(Boolean);
    const pts: string[] = [];
    for (let i = 0; i + 1 < nums.length; i += 2) {
      const [x, y] = lengthPair(`${nums[i]},${nums[i + 1]}`);
      pts.push(`${pts.length ? "L" : "M"}${x} ${y}`);
    }
    return [{ d: pts.join(" "), fill: true, stroke: true }];
  }
  if (local === "curve") {
    const [x1, y1] = lengthPair(attr(shape, "from"));
    const [c1x, c1y] = lengthPair(attr(shape, "control1"));
    const [c2x, c2y] = lengthPair(attr(shape, "control2"));
    const [x2, y2] = lengthPair(attr(shape, "to"));
    return [
      { d: `M${x1} ${y1} C${c1x} ${c1y} ${c2x} ${c2y} ${x2} ${y2}`, fill: true, stroke: true },
    ];
  }
  const path = attr(shape, "path") ?? vmlChildAttr(shape, "path", "v");
  if (!path) return [{ d: `M0 0 H${width} V${height} H0 Z`, fill: true, stroke: true }];
  const formulas = vmlChild(shape, "formulas");
  return vmlPathToSvg(path, {
    coordSize: pair(attr(shape, "coordsize"), [DEFAULT_COORD, DEFAULT_COORD]),
    coordOrigin: pair(attr(shape, "coordorigin"), [0, 0]),
    width,
    height,
    adj: (attr(shape, "adj") ?? "").split(",").map((s) => Number.parseFloat(s) || 0),
    formulas: formulas
      ? formulas.children
          .filter((c): c is XmlElement => c.kind === "element" && c.name.local === "f")
          .map((f) => attr(f, "eqn") ?? "")
      : [],
  });
}

function vmlChildAttr(shape: XmlElement, child: string, local: string): string | undefined {
  const el = vmlChild(shape, child);
  return el ? attr(el, local) : undefined;
}

// ---------------------------------------------------------------------------
// Text
// ---------------------------------------------------------------------------

function runText(run: WmlRun): string {
  let out = "";
  for (const piece of run.pieces) {
    if (piece.kind === "text") out += piece.value;
    else if (piece.kind === "tab") out += "\t";
    else if (piece.kind === "break") out += "\n";
  }
  return out;
}

/**
 * A text box's paragraphs. Runs carry `data-wk-txbx-run` so the canvas can map
 * typing inside an editable text box back to the run it changed.
 */
function textBoxHtml(paragraphs: readonly WmlParagraph[], styles: StyleResolver): string {
  return paragraphs
    .map((para, p) => {
      let r = 0;
      const runs = para.children
        .map((inline) => {
          if (inline.kind !== "run") return "";
          const html = `<span class="wk-txbx-run" data-wk-txbx-run="${p},${r}" style="${escapeHtml(runCss(styles.run(para, inline)))}">${escapeHtml(runText(inline))}</span>`;
          r++;
          return html;
        })
        .join("");
      const css = paragraphCss(styles.paragraph(para), styles.run(para));
      return `<p class="wk-txbx-p" data-wk-txbx-para="${p}" style="${escapeHtml(css)}">${runs || "\u200b"}</p>`;
    })
    .join("");
}

// ---------------------------------------------------------------------------
// Placement
// ---------------------------------------------------------------------------

const pt = (n: number): string => `${Math.round(n * 100) / 100}pt`;

/** CSS for the object's wrapper: where it sits and how text flows around it. */
function placementCss(layout: ShapeLayout, wrap: string): string {
  const css: string[] = [`width:${pt(layout.width)}`, `height:${pt(layout.height)}`];
  const transforms: string[] = [];
  if (layout.rotation) transforms.push(`rotate(${layout.rotation}deg)`);
  if (layout.flipH || layout.flipV)
    transforms.push(`scale(${layout.flipH ? -1 : 1},${layout.flipV ? -1 : 1})`);
  if (transforms.length) css.push(`transform:${transforms.join(" ")}`);
  if (layout.hidden) css.push("visibility:hidden");
  if (layout.inline) {
    css.push("display:inline-block", "position:relative", "vertical-align:bottom");
    return css.join(";");
  }
  const page = layout.horizontalRelativeTo === "page";
  const columnWidth = "(var(--page-w) - var(--m-left) - var(--m-right))";
  const w = pt(layout.width);
  let left: string;
  switch (layout.horizontalAlign) {
    case "center":
      left = page
        ? `calc((var(--page-w) - ${w}) / 2)`
        : `calc(var(--m-left) + (${columnWidth} - ${w}) / 2)`;
      break;
    case "right":
    case "outside":
      left = page ? `calc(var(--page-w) - ${w})` : `calc(var(--page-w) - var(--m-right) - ${w})`;
      break;
    case "left":
    case "inside":
      left = page ? "0pt" : "var(--m-left)";
      break;
    default:
      left = page ? pt(layout.left) : `calc(var(--m-left) + ${pt(layout.left)})`;
  }
  const verticalToPage =
    layout.verticalRelativeTo === "page" || layout.verticalRelativeTo === "margin";
  const flows =
    !verticalToPage &&
    (wrap === "square" || wrap === "tight" || wrap === "through" || wrap === "topAndBottom");
  if (flows) {
    // Offsets are from the column's left edge for a float.
    const offset = `calc(${left} - var(--m-left))`;
    css.push(`margin-top:${pt(layout.top)}`);
    if (wrap === "topAndBottom") {
      css.push("display:block", "clear:both", `margin-left:${offset}`, "position:relative");
    } else {
      css.push("float:left", `margin-left:${offset}`, "margin-right:9pt", "position:relative");
    }
    return css.join(";");
  }
  css.push("position:absolute", `left:${left}`);
  const h = pt(layout.height);
  if (layout.verticalRelativeTo === "page" || layout.verticalRelativeTo === "margin") {
    const onPage = layout.verticalRelativeTo === "page";
    const rowHeight = "(var(--page-h) - var(--m-top) - var(--m-bottom))";
    switch (layout.verticalAlign) {
      case "center":
        css.push(
          onPage
            ? `top:calc((var(--page-h) - ${h}) / 2)`
            : `top:calc(var(--m-top) + (${rowHeight} - ${h}) / 2)`,
        );
        break;
      case "bottom":
      case "outside":
        css.push(
          onPage
            ? `top:calc(var(--page-h) - ${h})`
            : `top:calc(var(--page-h) - var(--m-bottom) - ${h})`,
        );
        break;
      case "top":
      case "inside":
        css.push(onPage ? "top:0pt" : "top:var(--m-top)");
        break;
      default:
        css.push(onPage ? `top:${pt(layout.top)}` : `top:calc(var(--m-top) + ${pt(layout.top)})`);
    }
  } else {
    // Paragraph-relative: the static position is the top of the paragraph.
    css.push(`margin-top:${pt(layout.top)}`);
  }
  css.push(`z-index:${layout.zIndex < 0 ? -1 : 1}`);
  if (layout.zIndex < 0) css.push("pointer-events:auto");
  return css.join(";");
}

// ---------------------------------------------------------------------------
// Shapes
// ---------------------------------------------------------------------------

interface RenderedBox {
  /** Inner HTML (SVG geometry, text) drawn on a box of the given size. */
  readonly html: string;
}

function shapeBody(
  shape: XmlElement,
  width: number,
  height: number,
  ctx: VmlRenderContext,
  styles: () => StyleResolver,
  types: Map<string, XmlElement>,
  editable: boolean,
): RenderedBox {
  if (shape.name.local === "group")
    return { html: groupBody(shape, width, height, ctx, styles, types) };
  const kind = shapeKind(shape);
  const fill = getShapeFill(shape);
  const stroke = getShapeStroke(shape);
  const shadow = getShapeShadow(shape);
  const imagedata = vmlChild(shape, "imagedata");
  const parts: string[] = [];
  let defs = "";
  const svgStyle = shadow
    ? ` style="filter:drop-shadow(${shadow.offsetX}pt ${shadow.offsetY}pt 0 ${hex(shadow.color)}${shadow.opacity === undefined ? "" : opacityHex(shadow.opacity)})"`
    : "";

  if (imagedata) {
    const relId = attr(imagedata, "id", REL_NS) ?? attr(imagedata, "relid", OFFICE_NS);
    const image = relId ? vmlImageData(ctx.doc, relId, ctx.partName) : undefined;
    if (image) {
      parts.push(
        `<image href="${dataUrl(image.bytes, image.contentType)}" width="${width}" height="${height}" preserveAspectRatio="none"/>`,
      );
    }
  }

  const wordArt = kind === "wordArt" ? getWordArt(shape) : undefined;
  if (wordArt) {
    const paint = fillPaint(fill, ctx, width, height);
    const s = strokeAttrs(stroke);
    defs += paint.defs + s.defs;
    const fontSize = height * WORDART_EM;
    const font = wordArt.font.replace(/['"\\<>]/g, "");
    parts.push(
      `<text x="${width / 2}" y="${height / 2}" text-anchor="middle" dominant-baseline="central" ` +
        `textLength="${width}" lengthAdjust="spacingAndGlyphs" font-family="'${font}',sans-serif" font-size="${fontSize}" ` +
        `font-weight="${wordArt.bold ? "bold" : "normal"}" font-style="${wordArt.italic ? "italic" : "normal"}" ` +
        `fill="${paint.paint}"${paint.opacity === undefined ? "" : ` fill-opacity="${paint.opacity}"`} ${s.attrs}>` +
        `${escapeHtml(wordArt.text)}</text>`,
    );
  } else if (!imagedata) {
    const paint = fillPaint(fill, ctx, width, height);
    const s = strokeAttrs(stroke);
    defs += paint.defs + s.defs;
    const evenOdd = ` fill-rule="evenodd"`;
    for (const sub of geometryPaths(shape, width, height)) {
      const fillAttr = sub.fill
        ? `fill="${paint.paint}"${paint.opacity === undefined ? "" : ` fill-opacity="${paint.opacity}"`}${evenOdd}`
        : `fill="none"`;
      parts.push(`<path d="${sub.d}" ${fillAttr} ${sub.stroke ? s.attrs : `stroke="none"`}/>`);
    }
  }
  let html =
    `<svg class="wk-shape-svg" width="100%" height="100%" viewBox="0 0 ${Math.max(width, 0.01)} ${Math.max(height, 0.01)}" ` +
    `preserveAspectRatio="none" overflow="visible"${svgStyle}>${defs ? `<defs>${defs}</defs>` : ""}${parts.join("")}</svg>`;

  const paragraphs = kind === "textBox" ? shapeText(shape) : [];
  if (kind === "textBox") {
    const layout = getTextBoxLayout(shape);
    const [l, t, r, b] = layout.inset;
    const css = [
      `inset:${pt(t)} ${pt(r)} ${pt(b)} ${pt(l)}`,
      `justify-content:${layout.anchor === "middle" ? "center" : layout.anchor === "bottom" ? "flex-end" : "flex-start"}`,
    ];
    if (layout.direction !== "horizontal") css.push("writing-mode:vertical-rl");
    if (layout.direction === "vertical270") css.push("transform:rotate(180deg)");
    html +=
      `<div class="wk-txbx"${editable ? ` contenteditable="true" spellcheck="false"` : ""} style="${css.join(";")}">` +
      `${textBoxHtml(paragraphs, styles())}</div>`;
  }
  return { html };
}

function opacityHex(opacity: number): string {
  return Math.round(Math.max(0, Math.min(1, opacity)) * 255)
    .toString(16)
    .padStart(2, "0");
}

/** A group's members, placed by the group's child coordinate system. */
function groupBody(
  group: XmlElement,
  width: number,
  height: number,
  ctx: VmlRenderContext,
  styles: () => StyleResolver,
  types: Map<string, XmlElement>,
): string {
  const [ox, oy] = pair(attr(group, "coordorigin"), [0, 0]);
  const [cw, ch] = pair(attr(group, "coordsize"), [DEFAULT_COORD, DEFAULT_COORD]);
  const sx = width / (cw || 1);
  const sy = height / (ch || 1);
  return groupMembers(group)
    .map((member) => {
      const m = effective(member, types);
      const l = getShapeLayout(m);
      const w = l.width * sx;
      const h = l.height * sy;
      const css = [
        "position:absolute",
        `left:${pt((l.left - ox) * sx)}`,
        `top:${pt((l.top - oy) * sy)}`,
        `width:${pt(w)}`,
        `height:${pt(h)}`,
      ];
      if (l.rotation) css.push(`transform:rotate(${l.rotation}deg)`);
      const body = shapeBody(m, w, h, ctx, styles, types, false);
      return `<span class="wk-shape-part" style="${css.join(";")}">${body.html}</span>`;
    })
    .join("");
}

const KIND_TO_OBJECT: Readonly<Record<string, string>> = {
  shape: "shape",
  textBox: "textBox",
  wordArt: "shape",
  ink: "ink",
  group: "shape",
  canvas: "shape",
};

function renderShape(
  shape: XmlElement,
  ctx: VmlRenderContext,
  types: Map<string, XmlElement>,
  styles: () => StyleResolver,
): string {
  const merged = effective(shape, types);
  const layout = getShapeLayout(merged);
  let { width, height } = layout;
  if (merged.name.local === "line" && (!width || !height)) {
    const [x1, y1] = lengthPair(attr(merged, "from"));
    const [x2, y2] = lengthPair(attr(merged, "to"));
    width = Math.max(width, Math.abs(x2 - x1));
    height = Math.max(height, Math.abs(y2 - y1));
  }
  const kind = shapeKind(merged);
  const editable = ctx.at !== undefined && kind === "textBox";
  const body = shapeBody(merged, width, height, ctx, styles, types, editable);
  const css = placementCss({ ...layout, width, height }, getShapeWrap(merged));
  const objectKind = KIND_TO_OBJECT[kind] ?? "shape";
  const at = ctx.at
    ? ` data-wk-object="${objectKind}" data-wk-at="${escapeHtml(JSON.stringify(ctx.at))}"`
    : "";
  const name = attr(shape, "id");
  return (
    `<span class="wk-shape" contenteditable="false"${at}${name ? ` data-wk-name="${escapeHtml(name)}"` : ""} ` +
    `style="${css}">${body.html}</span>`
  );
}

/** Whether a VML drawing object floats (is positioned) rather than sitting in the line. */
function isFloating(shape: XmlElement): boolean {
  return /(^|;)\s*position\s*:\s*absolute/i.test(attr(shape, "style") ?? "");
}

/**
 * Render a `w:pict`'s drawing objects. `only` limits the output to the
 * floating or the inline ones (the body renderer emits floats at the start of
 * their paragraph).
 */
export function renderPictHtml(
  pict: XmlElement,
  ctx: VmlRenderContext,
  only?: "floating" | "inline",
): string {
  const types = shapeTypesFor(ctx.doc, pict);
  let resolver: StyleResolver | undefined;
  const styles = (): StyleResolver => (resolver ??= createStyleResolver(ctx.doc));
  return pict.children
    .filter(isDrawable)
    .filter((s) => only === undefined || (only === "floating") === isFloating(s))
    .map((s) => renderShape(s, ctx, types, styles))
    .join("");
}

/** Whether a run holds VML drawing objects. */
export function runHasPict(run: WmlRun): boolean {
  return run.pieces.some((p) => p.kind === "pict");
}

/** The floating VML objects of a paragraph's runs, for emitting at the paragraph start. */
export function paragraphFloatsHtml(
  para: WmlParagraph,
  doc: Docx,
  at: (inline: number) => DocPosition,
): string {
  let out = "";
  let runIndex = 0;
  for (const inline of para.children) {
    if (inline.kind !== "run") continue;
    for (const piece of inline.pieces) {
      if (piece.kind === "pict")
        out += renderPictHtml(piece.node, { doc, at: at(runIndex) }, "floating");
    }
    runIndex++;
  }
  return out;
}

/**
 * A run's inline VML objects and SmartArt graphics (floating VML is emitted
 * by {@link paragraphFloatsHtml}).
 */
export function runObjectsHtml(run: WmlRun, doc: Docx, at: DocPosition): string {
  let out = "";
  for (const piece of run.pieces) {
    if (piece.kind === "pict") out += renderPictHtml(piece.node, { doc, at }, "inline");
  }
  if (run.pieces.some((p) => p.kind === "drawing")) out += renderRunSmartArt(run, doc, at);
  return out;
}

/** SVG markup for a preset's geometry (gallery thumbnails). */
export function vmlPathPreviewSvg(path: string, size: number, coordSize = 21600): string {
  return vmlPathToSvg(path, { coordSize: [coordSize, coordSize], width: size, height: size })
    .map(
      (s) => `<path d="${s.d}"${s.fill ? "" : ` fill="none"`}${s.stroke ? "" : ` stroke="none"`}/>`,
    )
    .join("");
}
