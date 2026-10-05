/**
 * DrawingML pictures and charts on the canvas.
 *
 * A picture renders as an inline SVG so the DrawingML look maps onto SVG
 * primitives instead of approximations: the crop (`a:srcRect`) is the image's
 * placement, Crop to Shape is a clip path, Picture Border is a stroke of that
 * path, the blip effects (`a:lum`, `a:grayscl`, `a:duotone`, `a:biLevel`,
 * `a:clrChange`, `a:hsl`, `a:alphaModFix`) are an SVG filter on the image,
 * and shadow / glow / soft edges / reflection are filters on the whole shape.
 * Bevel and 3-D rotation are approximated (inner highlights, a CSS 3-D
 * transform).
 *
 * Floating objects (`wp:anchor`) are emitted where their anchor run is, with
 * their position, wrap and size in `data-wk-float`; {@link layoutFloatingObjects}
 * then places them against the page once the browser has laid out the text.
 */

import {
  type ChartSpec,
  type Docx,
  type DrawingInfo,
  imageDrawings,
  type PictureColorAdjustments,
  type PictureEffects,
  type PictureInfo,
  readChart,
  readDrawing,
  type WmlRun,
  type XmlElement,
} from "@office-kit/docx";
import { presetPath } from "./preset-geometry.js";
import { chartSvg } from "./render-chart.js";
import type { DocPosition } from "./selection.js";

/** EMU per CSS pixel (914400 per inch, 96 px per inch). */
export const EMU_PER_PX = 9525;
const ANGLE_FULL = 360;

const px = (emu: number): number => Math.round((emu / EMU_PER_PX) * 100) / 100;

function escapeAttr(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

const HEX = /^[0-9A-Fa-f]{6}$/;
const color = (hex: string): string => (HEX.test(hex) ? `#${hex}` : "#000");

function rgb(hex: string): [number, number, number] {
  const n = Number.parseInt(HEX.test(hex) ? hex : "000000", 16);
  return [((n >> 16) & 255) / 255, ((n >> 8) & 255) / 255, (n & 255) / 255];
}

// --- image data URLs -----------------------------------------------------------

const dataUrls = new WeakMap<Uint8Array, string>();

function base64(bytes: Uint8Array): string {
  let binary = "";
  const CHUNK = 0x8000;
  for (let i = 0; i < bytes.length; i += CHUNK) {
    binary += String.fromCharCode(...bytes.subarray(i, i + CHUNK));
  }
  return btoa(binary);
}

/** A `data:` URL for image bytes, memoized per byte array (re-renders reuse it). */
export function imageDataUrl(bytes: Uint8Array, contentType: string): string {
  const cached = dataUrls.get(bytes);
  if (cached) return cached;
  const url = `data:${contentType};base64,${base64(bytes)}`;
  dataUrls.set(bytes, url);
  return url;
}

// --- blip effects -------------------------------------------------------------

// a:clrChange matches colors within this distance per channel (0–255), as
// Word's Set Transparent Color tolerates JPEG noise around the picked color.
const TRANSPARENT_COLOR_TOLERANCE = 12;
// Resolution of the biLevel threshold step table.
const BILEVEL_STEPS = 100;
const LUMINANCE = "0.2126 0.7152 0.0722 0 0";

function luminanceMatrix(): string {
  return `<feColorMatrix type="matrix" values="${LUMINANCE} ${LUMINANCE} ${LUMINANCE} 0 0 0 1 0"/>`;
}

/** A 256-entry discrete table: 0 near channel value `t`, 1 elsewhere. */
function matchTable(t: number): string {
  return Array.from({ length: 256 }, (_, i) =>
    Math.abs(i - t) <= TRANSPARENT_COLOR_TOLERANCE ? 0 : 1,
  ).join(" ");
}

function adjustmentFilter(id: string, adj: PictureColorAdjustments): string | undefined {
  const steps: string[] = [];
  if (adj.transparentColor) {
    const target = rgb(adj.transparentColor).map((c) => Math.round(c * 255));

    steps.push(
      `<feComponentTransfer in="SourceGraphic" result="d"><feFuncR type="discrete" tableValues="${matchTable(target[0] ?? 0)}"/><feFuncG type="discrete" tableValues="${matchTable(target[1] ?? 0)}"/><feFuncB type="discrete" tableValues="${matchTable(target[2] ?? 0)}"/></feComponentTransfer>`,
      `<feColorMatrix in="d" type="matrix" values="0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 1 1 1 0 0" result="keep"/>`,
      `<feComposite in="SourceGraphic" in2="keep" operator="in"/>`,
    );
  }
  if (adj.saturation !== undefined && adj.saturation !== 100) {
    steps.push(`<feColorMatrix type="saturate" values="${Math.max(0, adj.saturation / 100)}"/>`);
  }
  if (adj.grayscale) steps.push(luminanceMatrix());
  if (adj.duotone) {
    const [a, b] = [rgb(adj.duotone[0]), rgb(adj.duotone[1])];
    steps.push(
      luminanceMatrix(),
      `<feComponentTransfer><feFuncR type="table" tableValues="${a[0]} ${b[0]}"/><feFuncG type="table" tableValues="${a[1]} ${b[1]}"/><feFuncB type="table" tableValues="${a[2]} ${b[2]}"/></feComponentTransfer>`,
    );
  }
  if (adj.biLevelThreshold !== undefined) {
    const t = adj.biLevelThreshold / 100;
    const table = Array.from({ length: BILEVEL_STEPS }, (_, i) =>
      i / BILEVEL_STEPS < t ? 0 : 1,
    ).join(" ");
    steps.push(
      luminanceMatrix(),
      `<feComponentTransfer><feFuncR type="discrete" tableValues="${table}"/><feFuncG type="discrete" tableValues="${table}"/><feFuncB type="discrete" tableValues="${table}"/></feComponentTransfer>`,
    );
  }
  const bright = (adj.brightness ?? 0) / 100;
  const contrast = (adj.contrast ?? 0) / 100;
  if (bright !== 0 || contrast !== 0) {
    // Contrast scales around mid-grey (steeper above 0, flatter below), then
    // brightness shifts: Word's Corrections presets read the same way.
    const slope = contrast >= 0 ? 1 / Math.max(0.01, 1 - contrast) : 1 + contrast;
    const intercept = 0.5 - 0.5 * slope + bright;
    const f = (c: string): string =>
      `<feFunc${c} type="linear" slope="${slope}" intercept="${intercept}"/>`;
    steps.push(`<feComponentTransfer>${f("R")}${f("G")}${f("B")}</feComponentTransfer>`);
  }
  if (!steps.length) return undefined;
  return `<filter id="${id}" color-interpolation-filters="sRGB">${steps.join("")}</filter>`;
}

// --- shape effects ------------------------------------------------------------

function effectFilter(id: string, fx: PictureEffects): string | undefined {
  const parts: string[] = [];
  let base = "SourceGraphic";
  if (fx.softEdgeEmu) {
    const r = px(fx.softEdgeEmu) / 2;
    parts.push(
      `<feMorphology in="SourceAlpha" operator="erode" radius="${r}" result="se1"/>`,
      `<feGaussianBlur in="se1" stdDeviation="${r}" result="se2"/>`,
      `<feComposite in="SourceGraphic" in2="se2" operator="in" result="soft"/>`,
    );
    base = "soft";
  }
  if (fx.bevel) {
    // A bevel lit from the top-left: a light inner edge top-left, a dark one bottom-right.
    const w = Math.max(1, px(fx.bevel.widthEmu) / 2);
    const inner = (name: string, dx: number, colorCss: string, opacity: number): string =>
      `<feComponentTransfer in="${base}" result="${name}i"><feFuncA type="table" tableValues="1 0"/></feComponentTransfer>` +
      `<feOffset in="${name}i" dx="${dx}" dy="${dx}" result="${name}o"/><feGaussianBlur in="${name}o" stdDeviation="${w / 2}" result="${name}b"/>` +
      `<feFlood flood-color="${colorCss}" flood-opacity="${opacity}"/><feComposite in2="${name}b" operator="in" result="${name}c"/>` +
      `<feComposite in="${name}c" in2="${base}" operator="in" result="${name}"/>`;
    parts.push(inner("hl", w, "#fff", 0.6), inner("sh", -w, "#000", 0.35));
    parts.push(
      `<feMerge result="bevel"><feMergeNode in="${base}"/><feMergeNode in="hl"/><feMergeNode in="sh"/></feMerge>`,
    );
    base = "bevel";
  }
  const under: string[] = [];
  const over: string[] = [];
  if (fx.glow) {
    const r = px(fx.glow.radiusEmu) / 2;
    parts.push(
      `<feMorphology in="${base}" operator="dilate" radius="${r}" result="gd"/>`,
      `<feGaussianBlur in="gd" stdDeviation="${r}" result="gb"/>`,
      `<feFlood flood-color="${color(fx.glow.color)}" flood-opacity="${fx.glow.opacity / 100}"/>`,
      `<feComposite in2="gb" operator="in" result="glow"/>`,
    );
    under.push("glow");
  }
  const s = fx.shadow;
  if (s) {
    const rad = (s.directionDeg * Math.PI) / (ANGLE_FULL / 2);
    const dx = Math.round(px(s.distanceEmu) * Math.cos(rad) * 100) / 100;
    const dy = Math.round(px(s.distanceEmu) * Math.sin(rad) * 100) / 100;
    const blur = px(s.blurEmu) / 2;
    const flood = `<feFlood flood-color="${color(s.color)}" flood-opacity="${s.opacity / 100}"/>`;
    if (s.kind === "outer") {
      parts.push(
        `<feGaussianBlur in="${base}" stdDeviation="${blur}" result="sb"/><feOffset in="sb" dx="${dx}" dy="${dy}" result="so"/>`,
        `${flood}<feComposite in2="so" operator="in" result="shadow"/>`,
      );
      under.push("shadow");
    } else {
      parts.push(
        `<feComponentTransfer in="${base}" result="si"><feFuncA type="table" tableValues="1 0"/></feComponentTransfer>`,
        `<feGaussianBlur in="si" stdDeviation="${blur}" result="sb"/><feOffset in="sb" dx="${dx}" dy="${dy}" result="so"/>`,
        `${flood}<feComposite in2="so" operator="in" result="sc"/><feComposite in="sc" in2="${base}" operator="in" result="inner"/>`,
      );
      over.push("inner");
    }
  }
  if (!parts.length) return undefined;
  const nodes = [...under, base, ...over].map((n) => `<feMergeNode in="${n}"/>`).join("");
  // The filter region grows to hold glow and shadow outside the picture.
  return `<filter id="${id}" x="-50%" y="-50%" width="200%" height="200%" color-interpolation-filters="sRGB">${parts.join("")}<feMerge>${nodes}</feMerge></filter>`;
}

// Approximate camera orientations (degrees) of the 3-D Rotation gallery presets.
const CAMERA_ANGLES: Readonly<
  Record<string, readonly [lat: number, lon: number, perspective: boolean]>
> = {
  isometricOffAxis1Left: [10, 45, false],
  isometricOffAxis1Right: [10, -45, false],
  isometricOffAxis1Top: [-45, 10, false],
  isometricOffAxis2Left: [10, 30, false],
  isometricOffAxis2Right: [10, -30, false],
  isometricOffAxis2Top: [-30, -10, false],
  isometricLeftDown: [30, 45, false],
  isometricRightUp: [-30, -45, false],
  isometricTopUp: [-45, 0, false],
  isometricBottomDown: [45, 0, false],
  perspectiveFront: [0, 0, true],
  perspectiveLeft: [0, 20, true],
  perspectiveRight: [0, -20, true],
  perspectiveAbove: [-20, 0, true],
  perspectiveBelow: [20, 0, true],
  perspectiveRelaxed: [-30, 0, true],
  perspectiveRelaxedModerately: [-15, 0, true],
  perspectiveContrastingLeftFacing: [10, 40, true],
  perspectiveContrastingRightFacing: [10, -40, true],
  perspectiveHeroicExtremeLeftFacing: [15, 35, true],
  perspectiveHeroicExtremeRightFacing: [15, -35, true],
  obliqueTopLeft: [-20, 20, false],
  obliqueTopRight: [-20, -20, false],
  obliqueBottomLeft: [20, 20, false],
  obliqueBottomRight: [20, -20, false],
};

const PERSPECTIVE_PX = 600;

function rotation3dCss(fx: PictureEffects): string {
  const r = fx.rotation3d;
  if (!r) return "";
  const [lat, lon, perspective] = CAMERA_ANGLES[r.camera] ?? [
    0,
    0,
    r.camera.startsWith("perspective"),
  ];
  const x = r.latitude ?? lat;
  const y = r.longitude ?? lon;
  const z = r.revolution ?? 0;
  return `${perspective ? `perspective(${PERSPECTIVE_PX}px) ` : ""}rotateX(${-x}deg) rotateY(${-y}deg) rotateZ(${z}deg)`;
}

const DASH: Readonly<Record<string, readonly number[]>> = {
  dot: [1, 1],
  sysDot: [1, 1],
  dash: [4, 3],
  sysDash: [3, 1],
  lgDash: [8, 3],
  dashDot: [4, 3, 1, 3],
  sysDashDot: [3, 1, 1, 1],
  lgDashDot: [8, 3, 1, 3],
  lgDashDotDot: [8, 3, 1, 3, 1, 3],
  sysDashDotDot: [3, 1, 1, 1, 1, 1],
};

export interface PictureSvgInput {
  readonly width: number;
  readonly height: number;
  readonly href: string;
  readonly picture: PictureInfo;
  readonly rotation: number;
  readonly flipH: boolean;
  readonly flipV: boolean;
}

/**
 * The SVG for one picture, `width` × `height` CSS px. `idPrefix` keeps the
 * filter / clip ids unique in the page (gallery previews render pictures too).
 */
export function pictureSvg(input: PictureSvgInput, idPrefix: string): string {
  const { width: w, height: h, picture: p } = input;
  const crop = p.crop;
  const visW = Math.max(0.01, 1 - (crop.left + crop.right) / 100);
  const visH = Math.max(0.01, 1 - (crop.top + crop.bottom) / 100);
  const imgW = w / visW;
  const imgH = h / visH;
  const x = (-crop.left / 100) * imgW;
  const y = (-crop.top / 100) * imgH;
  const geom = presetPath(p.geometry, w, h);
  const defs: string[] = [`<clipPath id="${idPrefix}c"><path d="${geom}"/></clipPath>`];
  const adj = adjustmentFilter(`${idPrefix}a`, p.adjustments);
  if (adj) defs.push(adj);
  const fx = effectFilter(`${idPrefix}f`, p.effects);
  if (fx) defs.push(fx);
  const opacity = p.adjustments.transparency
    ? ` opacity="${1 - p.adjustments.transparency / 100}"`
    : "";
  const image = `<image href="${escapeAttr(input.href)}" x="${x}" y="${y}" width="${imgW}" height="${imgH}" preserveAspectRatio="none"${adj ? ` filter="url(#${idPrefix}a)"` : ""}${opacity}/>`;
  let outline = "";
  if (p.outline) {
    const sw = Math.max(px(p.outline.widthEmu), 0.75);
    const dash = DASH[p.outline.dash];
    const dashAttr = dash ? ` stroke-dasharray="${dash.map((d) => d * sw).join(" ")}"` : "";
    if (
      p.outline.compound === "dbl" ||
      p.outline.compound === "thickThin" ||
      p.outline.compound === "thinThick"
    ) {
      // Two lines inside one stroke: draw the full width, then cut a gap with a thinner white-out mask.
      outline = `<path d="${geom}" fill="none" stroke="${color(p.outline.color)}" stroke-width="${sw}"${dashAttr}/><path d="${geom}" fill="none" stroke="#fff" stroke-width="${sw / 3}"${dashAttr}/>`;
    } else {
      outline = `<path d="${geom}" fill="none" stroke="${color(p.outline.color)}" stroke-width="${sw}"${dashAttr}/>`;
    }
  }
  const body = `<g id="${idPrefix}b"${fx ? ` filter="url(#${idPrefix}f)"` : ""}><g clip-path="url(#${idPrefix}c)">${image}</g>${outline}</g>`;
  let reflection = "";
  const r = p.effects.reflection;
  if (r) {
    const dist = px(r.distanceEmu);
    const endY = (h * r.endPosition) / 100;
    defs.push(
      `<linearGradient id="${idPrefix}g" gradientUnits="userSpaceOnUse" x1="0" y1="${h + dist}" x2="0" y2="${h + dist + Math.max(1, endY)}"><stop offset="0" stop-color="#fff" stop-opacity="${r.startOpacity / 100}"/><stop offset="1" stop-color="#fff" stop-opacity="${r.endOpacity / 100}"/></linearGradient>`,
      `<mask id="${idPrefix}m" maskUnits="userSpaceOnUse" x="${-w}" y="${h}" width="${3 * w}" height="${2 * h}"><rect x="${-w}" y="${h + dist}" width="${3 * w}" height="${endY}" fill="url(#${idPrefix}g)"/></mask>`,
    );
    if (r.blurEmu)
      defs.push(
        `<filter id="${idPrefix}rb"><feGaussianBlur stdDeviation="${px(r.blurEmu) / 2}"/></filter>`,
      );
    reflection = `<g mask="url(#${idPrefix}m)"><g transform="translate(0 ${2 * h + dist}) scale(1 -1)"${r.blurEmu ? ` filter="url(#${idPrefix}rb)"` : ""}><use href="#${idPrefix}b"/></g></g>`;
  }
  const transforms: string[] = [];
  if (input.rotation) transforms.push(`rotate(${input.rotation}deg)`);
  if (input.flipH || input.flipV)
    transforms.push(`scale(${input.flipH ? -1 : 1},${input.flipV ? -1 : 1})`);
  const threeD = rotation3dCss(p.effects);
  if (threeD) transforms.push(threeD);
  const style = transforms.length ? ` style="transform:${transforms.join(" ")}"` : "";
  return `<svg class="wk-pic" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}" overflow="visible"${style} aria-hidden="true"><defs>${defs.join("")}</defs>${body}${reflection}</svg>`;
}

// --- drawings in runs ------------------------------------------------------------

/** What the float layout pass needs, serialized into `data-wk-float`. */
export interface FloatData {
  readonly h: NonNullable<DrawingInfo["anchor"]>["horizontal"];
  readonly v: NonNullable<DrawingInfo["anchor"]>["vertical"];
  readonly wrap: DrawingInfo["wrap"];
  readonly side: NonNullable<DrawingInfo["anchor"]>["wrapSide"];
  readonly dist: readonly [top: number, right: number, bottom: number, left: number];
  readonly w: number;
  readonly hgt: number;
  readonly polygon?: ReadonlyArray<readonly [number, number]>;
}

/** Per-render state: drawing indices and z-order ranks, computed once per pass. */
export interface DrawingRenderContext {
  readonly doc: Docx;
  readonly index: ReadonlyMap<XmlElement, number>;
  readonly zRank: ReadonlyMap<XmlElement, number>;
  readonly charts: Map<string, ChartSpec | undefined>;
}

export function createDrawingContext(doc: Docx): DrawingRenderContext {
  const all = imageDrawings(doc);
  const index = new Map(all.map((d, i) => [d, i] as const));
  const heights = all.flatMap((d) => {
    const placement = d.children.find((c) => c.kind === "element" && c.name.local === "anchor");
    const raw =
      placement?.kind === "element"
        ? placement.attrs.find((a) => a.name.local === "relativeHeight")?.value
        : undefined;
    return raw === undefined ? [] : [[d, Number(raw) || 0] as const];
  });
  const zRank = new Map(
    heights.toSorted((a, b) => a[1] - b[1]).map(([d], rank) => [d, rank] as const),
  );
  return { doc, index, zRank, charts: new Map() };
}

const DIAGRAM_URI = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
// Floating objects stack above the text by z-rank; behind-text ones go under it.
const FRONT_Z_BASE = 10;
const BEHIND_Z_BASE = -1000;

function isDiagram(node: XmlElement): boolean {
  const stack: XmlElement[] = [node];
  while (stack.length) {
    const el = stack.pop();
    if (!el) break;
    if (el.name.local === "graphicData")
      return el.attrs.some((a) => a.name.local === "uri" && a.value === DIAGRAM_URI);
    for (const c of el.children) if (c.kind === "element") stack.push(c);
  }
  return false;
}

function chartFor(
  ctx: DrawingRenderContext,
  drawing: XmlElement,
  partName: string,
): ChartSpec | undefined {
  if (!ctx.charts.has(partName)) ctx.charts.set(partName, readChart(ctx.doc, drawing));
  return ctx.charts.get(partName);
}

function renderOne(ctx: DrawingRenderContext, node: XmlElement, at: DocPosition): string {
  if (isDiagram(node)) return "";
  const info = readDrawing(ctx.doc, node);
  if (info.kind === "other") return "";
  const index = ctx.index.get(node) ?? -1;
  const w = px(info.widthEmu);
  const h = px(info.heightEmu);
  let inner: string;
  if (info.kind === "picture" && info.picture) {
    const image = info.picture.image;
    const href = image ? imageDataUrl(image.data, image.contentType) : "";
    inner = pictureSvg(
      {
        width: w,
        height: h,
        href,
        picture: info.picture,
        rotation: info.rotation,
        flipH: info.flipH,
        flipV: info.flipV,
      },
      `wkd${index}-`,
    );
  } else {
    const spec = info.chartPartName ? chartFor(ctx, node, info.chartPartName) : undefined;
    inner = spec ? chartSvg(spec, w, h) : "";
  }
  // inline-block so the box (what selection measures) is the object's size, not the line's.
  const css: string[] = ["display:inline-block", `width:${w}px`, `height:${h}px`];
  let floatAttr = "";
  if (info.anchor) {
    const a = info.anchor;
    const data: FloatData = {
      h: a.horizontal,
      v: a.vertical,
      wrap: info.wrap,
      side: a.wrapSide,
      dist: [px(a.distance.top), px(a.distance.right), px(a.distance.bottom), px(a.distance.left)],
      w,
      hgt: h,
      ...(a.wrapPolygon ? { polygon: a.wrapPolygon.map((p) => [p.x, p.y] as const) } : {}),
    };
    floatAttr = ` data-wk-float="${escapeAttr(JSON.stringify(data))}"`;
    const rank = ctx.zRank.get(node) ?? 0;
    css.push(`z-index:${info.wrap === "behindText" ? BEHIND_Z_BASE + rank : FRONT_Z_BASE + rank}`);
  }
  const title = info.hyperlink
    ? ` title="${escapeAttr(info.hyperlink)}" data-wk-href="${escapeAttr(info.hyperlink)}"`
    : "";
  const label = info.description || info.name;
  // `hidden` rather than `display:none`: the float layout pass rewrites `display`.
  const hidden = info.hidden ? " hidden" : "";
  return (
    `<span class="wk-obj ${info.anchor ? "wk-obj-float" : "wk-obj-inline"}" contenteditable="false" data-wk-object="${info.kind}"` +
    ` data-wk-drawing="${index}" data-wk-at="${escapeAttr(JSON.stringify(at))}"${floatAttr}${title}` +
    ` role="img" aria-label="${escapeAttr(label)}"${hidden} style="${css.join(";")}">${inner}</span>`
  );
}

/**
 * HTML for the pictures and charts in `run` (anything else in a `<w:drawing>`
 * renders nothing here). `at` is the run's position, kept on the object so a
 * click can select it.
 */
export function runDrawingsHtml(ctx: DrawingRenderContext, run: WmlRun, at: DocPosition): string {
  let out = "";
  for (const piece of run.pieces) {
    if (piece.kind === "drawing") out += renderOne(ctx, piece.node, at);
  }
  return out;
}

/** Whether a run holds only drawings (no text a caret could sit in). */
export function isDrawingOnlyRun(run: WmlRun): boolean {
  return run.pieces.length > 0 && run.pieces.every((p) => p.kind === "drawing");
}
