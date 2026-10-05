/**
 * The look of a `pic:pic` (ECMA-376 Part 1, §19.3.1 and §20.1): crop
 * (`a:srcRect`), shape (`a:prstGeom`), border (`a:ln`), effects
 * (`a:effectLst`, `a:scene3d`, `a:sp3d`), transform (`a:xfrm`) and the blip
 * effects behind Corrections, Color and Transparency.
 */

import type { XmlAttr, XmlElement, XmlNode } from "../xml/index.js";
import { colorElements, readColor, resolveColor, srgbMarkup } from "./color.js";
import type {
  CompoundLine,
  LineDash,
  Picture3dRotation,
  PictureColorAdjustments,
  PictureCrop,
  PictureEffects,
  PictureOutline,
  PictureShadow,
  RectAlignment,
  BevelPreset,
} from "./types.js";
import {
  A_NS,
  attr,
  boolAttr,
  child,
  clearSlot,
  descendant,
  elementChildren,
  ensureDrawingNamespaces,
  fragment,
  numAttr,
  path,
  PIC_NS,
  placeChild,
  R_NS,
  removeChildren,
  setAttr,
} from "./xml.js";

// DrawingML units: angles in 60000ths of a degree, percentages in 1000ths.
export const ANGLE_UNIT = 60000;
export const PERCENT_UNIT = 1000;

const SP_PR_ORDER = [
  "xfrm",
  "custGeom|prstGeom",
  "noFill|solidFill|gradFill|blipFill|pattFill|grpFill",
  "ln",
  "effectLst|effectDag",
  "scene3d",
  "sp3d",
  "extLst",
] as const;
const BLIP_FILL_ORDER = ["blip", "srcRect", "tile|stretch"] as const;
const XFRM_ORDER = ["off", "ext"] as const;
// The blip effects are one repeating choice, then extLst; this module owns the
// ones listed in BLIP_EFFECTS and leaves any others (and extLst) untouched.
const BLIP_EFFECTS: ReadonlySet<string> = new Set([
  "lum",
  "grayscl",
  "biLevel",
  "duotone",
  "clrChange",
  "alphaModFix",
  "hsl",
]);
const SHADOW_TYPES: ReadonlySet<string> = new Set(["outerShdw", "innerShdw"]);

export function picOf(drawing: XmlElement): XmlElement | undefined {
  return descendant(drawing, PIC_NS, "pic");
}

function spPrOf(pic: XmlElement): XmlElement {
  const spPr = child(pic, PIC_NS, "spPr");
  if (!spPr) throw new Error("The picture has no pic:spPr.");
  return spPr;
}

function blipFillOf(pic: XmlElement): XmlElement {
  const fill = child(pic, PIC_NS, "blipFill");
  if (!fill) throw new Error("The picture has no pic:blipFill.");
  return fill;
}

export function blipOf(pic: XmlElement): XmlElement | undefined {
  return child(child(pic, PIC_NS, "blipFill"), A_NS, "blip");
}

// --- transform ----------------------------------------------------------------

export function readTransform(pic: XmlElement | undefined): {
  rotation: number;
  flipH: boolean;
  flipV: boolean;
} {
  const xfrm = path(pic, [PIC_NS, "spPr"], [A_NS, "xfrm"]);
  return {
    rotation: (numAttr(xfrm, "rot") ?? 0) / ANGLE_UNIT,
    flipH: boolAttr(xfrm, "flipH") ?? false,
    flipV: boolAttr(xfrm, "flipV") ?? false,
  };
}

function ensureXfrm(spPr: XmlElement, cx: number, cy: number): XmlElement {
  const existing = child(spPr, A_NS, "xfrm");
  if (existing) return existing;
  const xfrm = fragment(`<a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>`);
  placeChild(spPr, xfrm, SP_PR_ORDER);
  return xfrm;
}

/** Normalize degrees into [0, 360). */
export function normalizeDegrees(deg: number): number {
  return ((deg % 360) + 360) % 360;
}

export function applyTransform(
  drawing: XmlElement,
  pic: XmlElement,
  size: { cx: number; cy: number },
  t: { rotation?: number; flipH?: boolean; flipV?: boolean },
): void {
  ensureDrawingNamespaces(drawing);
  const xfrm = ensureXfrm(spPrOf(pic), size.cx, size.cy);
  if (t.rotation !== undefined) {
    if (!Number.isFinite(t.rotation)) throw new Error(`Invalid rotation: ${t.rotation}.`);
    const rot = Math.round(normalizeDegrees(t.rotation) * ANGLE_UNIT);
    setAttr(xfrm, "rot", rot === 0 ? undefined : String(rot));
  }
  if (t.flipH !== undefined) setAttr(xfrm, "flipH", t.flipH ? "1" : undefined);
  if (t.flipV !== undefined) setAttr(xfrm, "flipV", t.flipV ? "1" : undefined);
}

/** Keep `a:xfrm/a:ext` in step with `wp:extent` after a resize. */
export function syncXfrmExtent(pic: XmlElement, cx: number, cy: number): void {
  const xfrm = path(pic, [PIC_NS, "spPr"], [A_NS, "xfrm"]);
  if (!xfrm) return;
  placeChild(xfrm, fragment(`<a:ext cx="${Math.round(cx)}" cy="${Math.round(cy)}"/>`), XFRM_ORDER);
}

// --- crop ----------------------------------------------------------------------

export function readCrop(pic: XmlElement): PictureCrop {
  const rect = child(child(pic, PIC_NS, "blipFill"), A_NS, "srcRect");
  const pct = (local: string): number => (numAttr(rect, local) ?? 0) / PERCENT_UNIT;
  return { left: pct("l"), top: pct("t"), right: pct("r"), bottom: pct("b") };
}

export function applyCrop(drawing: XmlElement, pic: XmlElement, crop: PictureCrop): void {
  ensureDrawingNamespaces(drawing);
  for (const v of Object.values(crop)) {
    if (!Number.isFinite(v)) throw new Error(`Invalid crop value: ${v}.`);
  }
  if (crop.left + crop.right >= 100 || crop.top + crop.bottom >= 100) {
    throw new Error("A crop must leave part of the picture visible.");
  }
  const fill = blipFillOf(pic);
  const empty = crop.left === 0 && crop.top === 0 && crop.right === 0 && crop.bottom === 0;
  if (empty) {
    clearSlot(fill, A_NS, "srcRect", BLIP_FILL_ORDER);
    // An empty <a:srcRect/> is what Word writes for an uncropped picture.
    placeChild(fill, fragment(`<a:srcRect/>`), BLIP_FILL_ORDER);
    return;
  }
  const v = (n: number): string => String(Math.round(n * PERCENT_UNIT));
  placeChild(
    fill,
    fragment(
      `<a:srcRect l="${v(crop.left)}" t="${v(crop.top)}" r="${v(crop.right)}" b="${v(crop.bottom)}"/>`,
    ),
    BLIP_FILL_ORDER,
  );
}

// --- geometry -----------------------------------------------------------------

export function readGeometry(pic: XmlElement): string {
  return attr(path(pic, [PIC_NS, "spPr"], [A_NS, "prstGeom"]), "prst") ?? "rect";
}

const SHAPE_NAME = /^[A-Za-z0-9]+$/;

export function applyGeometry(drawing: XmlElement, pic: XmlElement, preset: string): void {
  if (!SHAPE_NAME.test(preset)) throw new Error(`Invalid shape preset ${JSON.stringify(preset)}.`);
  ensureDrawingNamespaces(drawing);
  placeChild(
    spPrOf(pic),
    fragment(`<a:prstGeom prst="${preset}"><a:avLst/></a:prstGeom>`),
    SP_PR_ORDER,
  );
}

// --- outline ------------------------------------------------------------------

export function readOutline(pic: XmlElement): PictureOutline | undefined {
  const ln = path(pic, [PIC_NS, "spPr"], [A_NS, "ln"]);
  const fill = child(ln, A_NS, "solidFill");
  const color = readColor(fill);
  if (!ln || !color) return undefined;
  return {
    color: color.hex,
    widthEmu: numAttr(ln, "w") ?? 9525,
    dash: (attr(child(ln, A_NS, "prstDash"), "val") ?? "solid") as LineDash,
    compound: (attr(ln, "cmpd") ?? "sng") as CompoundLine,
  };
}

/** Set the picture border; `undefined` removes it (`a:ln` with `a:noFill`, as Word writes No Outline). */
export function applyOutline(
  drawing: XmlElement,
  pic: XmlElement,
  outline: PictureOutline | undefined,
): void {
  ensureDrawingNamespaces(drawing);
  const spPr = spPrOf(pic);
  if (!outline) {
    placeChild(spPr, fragment(`<a:ln><a:noFill/></a:ln>`), SP_PR_ORDER);
    return;
  }
  if (!Number.isFinite(outline.widthEmu) || outline.widthEmu < 0) {
    throw new Error(`Invalid border weight: ${outline.widthEmu}.`);
  }
  const cmpd = outline.compound === "sng" ? "" : ` cmpd="${outline.compound}"`;
  placeChild(
    spPr,
    fragment(
      `<a:ln w="${Math.round(outline.widthEmu)}"${cmpd}><a:solidFill>${srgbMarkup(outline.color)}</a:solidFill><a:prstDash val="${outline.dash}"/></a:ln>`,
    ),
    SP_PR_ORDER,
  );
}

// --- effects ------------------------------------------------------------------

function readShadow(el: XmlElement): PictureShadow {
  const color = readColor(el);
  const kind = el.name.local === "innerShdw" ? "inner" : "outer";
  const sx = numAttr(el, "sx");
  const sy = numAttr(el, "sy");
  const kx = numAttr(el, "kx");
  const algn = attr(el, "algn");
  return {
    kind,
    color: color?.hex ?? "000000",
    opacity: color?.alpha ?? 100,
    blurEmu: numAttr(el, "blurRad") ?? 0,
    distanceEmu: numAttr(el, "dist") ?? 0,
    directionDeg: (numAttr(el, "dir") ?? 0) / ANGLE_UNIT,
    ...(sx !== undefined ? { scaleX: sx / PERCENT_UNIT } : {}),
    ...(sy !== undefined ? { scaleY: sy / PERCENT_UNIT } : {}),
    ...(kx !== undefined ? { skewX: kx / ANGLE_UNIT } : {}),
    ...(algn !== undefined ? { align: algn as RectAlignment } : {}),
  };
}

export function readEffects(pic: XmlElement): PictureEffects {
  const spPr = child(pic, PIC_NS, "spPr");
  const lst = child(spPr, A_NS, "effectLst");
  const shadowEl = lst
    ? elementChildren(lst).find((c) => c.name.uri === A_NS && SHADOW_TYPES.has(c.name.local))
    : undefined;
  const refl = child(lst, A_NS, "reflection");
  const glow = child(lst, A_NS, "glow");
  const soft = child(lst, A_NS, "softEdge");
  const bevel = path(spPr, [A_NS, "sp3d"], [A_NS, "bevelT"]);
  const camera = path(spPr, [A_NS, "scene3d"], [A_NS, "camera"]);
  const rot = child(camera, A_NS, "rot");
  const glowColor = readColor(glow);
  const rotation3d: Picture3dRotation | undefined =
    camera && attr(camera, "prst") !== "orthographicFront"
      ? {
          camera: attr(camera, "prst") ?? "orthographicFront",
          ...(rot
            ? {
                latitude: (numAttr(rot, "lat") ?? 0) / ANGLE_UNIT,
                longitude: (numAttr(rot, "lon") ?? 0) / ANGLE_UNIT,
                revolution: (numAttr(rot, "rev") ?? 0) / ANGLE_UNIT,
              }
            : {}),
        }
      : undefined;
  return {
    ...(shadowEl ? { shadow: readShadow(shadowEl) } : {}),
    ...(refl
      ? {
          reflection: {
            blurEmu: numAttr(refl, "blurRad") ?? 0,
            startOpacity: (numAttr(refl, "stA") ?? 100 * PERCENT_UNIT) / PERCENT_UNIT,
            endOpacity: (numAttr(refl, "endA") ?? 0) / PERCENT_UNIT,
            endPosition: (numAttr(refl, "endPos") ?? 100 * PERCENT_UNIT) / PERCENT_UNIT,
            distanceEmu: numAttr(refl, "dist") ?? 0,
          },
        }
      : {}),
    ...(glow
      ? {
          glow: {
            radiusEmu: numAttr(glow, "rad") ?? 0,
            color: glowColor?.hex ?? "000000",
            opacity: glowColor?.alpha ?? 100,
          },
        }
      : {}),
    ...(soft ? { softEdgeEmu: numAttr(soft, "rad") ?? 0 } : {}),
    ...(bevel
      ? {
          bevel: {
            preset: (attr(bevel, "prst") ?? "circle") as BevelPreset,
            widthEmu: numAttr(bevel, "w") ?? 76200,
            heightEmu: numAttr(bevel, "h") ?? 76200,
          },
        }
      : {}),
    ...(rotation3d ? { rotation3d } : {}),
  };
}

const emu = (n: number): number => Math.max(0, Math.round(n));
const angle = (deg: number): number => Math.round(normalizeDegrees(deg) * ANGLE_UNIT);
const pct = (n: number): number => Math.round(n * PERCENT_UNIT);

function shadowMarkup(s: PictureShadow): string {
  const color = srgbMarkup(s.color, s.opacity);
  const common = `blurRad="${emu(s.blurEmu)}" dist="${emu(s.distanceEmu)}" dir="${angle(s.directionDeg)}"`;
  if (s.kind === "inner") return `<a:innerShdw ${common}>${color}</a:innerShdw>`;
  const extra = [
    s.scaleX !== undefined ? ` sx="${pct(s.scaleX)}"` : "",
    s.scaleY !== undefined ? ` sy="${pct(s.scaleY)}"` : "",
    s.skewX !== undefined ? ` kx="${Math.round(s.skewX * ANGLE_UNIT)}"` : "",
    ` algn="${s.align ?? "ctr"}" rotWithShape="0"`,
  ].join("");
  return `<a:outerShdw ${common}${extra}>${color}</a:outerShdw>`;
}

/**
 * Replace the picture's effects. The effect list is written in the
 * CT_EffectList order (glow, innerShdw, outerShdw, reflection, softEdge); an
 * empty `effects` removes `a:effectLst`, `a:scene3d` and `a:sp3d` entirely.
 */
export function applyEffects(drawing: XmlElement, pic: XmlElement, effects: PictureEffects): void {
  ensureDrawingNamespaces(drawing);
  const spPr = spPrOf(pic);
  const parts: string[] = [];
  if (effects.glow) {
    parts.push(
      `<a:glow rad="${emu(effects.glow.radiusEmu)}">${srgbMarkup(effects.glow.color, effects.glow.opacity)}</a:glow>`,
    );
  }
  if (effects.shadow?.kind === "inner") parts.push(shadowMarkup(effects.shadow));
  if (effects.shadow?.kind === "outer") parts.push(shadowMarkup(effects.shadow));
  if (effects.reflection) {
    const r = effects.reflection;
    // A reflection is the picture mirrored below itself: flipped vertically
    // (sy = -100 %), anchored at the bottom-left, fading from stA to endA.
    parts.push(
      `<a:reflection blurRad="${emu(r.blurEmu)}" stA="${pct(r.startOpacity)}" endA="${pct(r.endOpacity)}" endPos="${pct(r.endPosition)}" dist="${emu(r.distanceEmu)}" dir="5400000" sy="-100000" algn="bl" rotWithShape="0"/>`,
    );
  }
  if (effects.softEdgeEmu !== undefined)
    parts.push(`<a:softEdge rad="${emu(effects.softEdgeEmu)}"/>`);
  clearSlot(spPr, A_NS, "effectLst", SP_PR_ORDER);
  if (parts.length)
    placeChild(spPr, fragment(`<a:effectLst>${parts.join("")}</a:effectLst>`), SP_PR_ORDER);

  removeChildren(spPr, A_NS, new Set(["scene3d", "sp3d"]));
  const rot = effects.rotation3d;
  if (rot || effects.bevel) {
    const rotMarkup =
      rot && rot.latitude !== undefined
        ? `<a:rot lat="${angle(rot.latitude)}" lon="${angle(rot.longitude ?? 0)}" rev="${angle(rot.revolution ?? 0)}"/>`
        : "";
    placeChild(
      spPr,
      fragment(
        `<a:scene3d><a:camera prst="${rot?.camera ?? "orthographicFront"}">${rotMarkup}</a:camera><a:lightRig rig="threePt" dir="t"/></a:scene3d>`,
      ),
      SP_PR_ORDER,
    );
  }
  if (effects.bevel) {
    const b = effects.bevel;
    placeChild(
      spPr,
      fragment(
        `<a:sp3d><a:bevelT w="${emu(b.widthEmu)}" h="${emu(b.heightEmu)}" prst="${b.preset}"/></a:sp3d>`,
      ),
      SP_PR_ORDER,
    );
  }
}

// --- blip effects -------------------------------------------------------------

export function readAdjustments(pic: XmlElement): PictureColorAdjustments {
  const blip = blipOf(pic);
  if (!blip) return {};
  const out: { -readonly [K in keyof PictureColorAdjustments]: PictureColorAdjustments[K] } = {};
  for (const el of elementChildren(blip)) {
    if (el.name.uri !== A_NS) continue;
    switch (el.name.local) {
      case "lum": {
        const bright = numAttr(el, "bright");
        const contrast = numAttr(el, "contrast");
        if (bright !== undefined) out.brightness = bright / PERCENT_UNIT;
        if (contrast !== undefined) out.contrast = contrast / PERCENT_UNIT;
        break;
      }
      case "hsl": {
        const sat = numAttr(el, "sat");
        if (sat !== undefined) out.saturation = 100 + sat / PERCENT_UNIT;
        break;
      }
      case "grayscl":
        out.grayscale = true;
        break;
      case "biLevel":
        out.biLevelThreshold = (numAttr(el, "thresh") ?? 50 * PERCENT_UNIT) / PERCENT_UNIT;
        break;
      case "duotone": {
        const [dark, light] = colorElements(el).map(resolveColor);
        if (dark && light) out.duotone = [dark.hex, light.hex];
        break;
      }
      case "clrChange": {
        const from = readColor(child(el, A_NS, "clrFrom"));
        if (from) out.transparentColor = from.hex;
        break;
      }
      case "alphaModFix":
        out.transparency = 100 - (numAttr(el, "amt") ?? 100 * PERCENT_UNIT) / PERCENT_UNIT;
        break;
      default:
        break;
    }
  }
  return out;
}

function clampPercent(n: number, min: number, max: number, what: string): number {
  if (!Number.isFinite(n) || n < min || n > max) throw new Error(`Invalid ${what}: ${n}.`);
  return n;
}

/**
 * Replace the picture's color adjustments. Each present field becomes one
 * blip effect; absent fields remove theirs. Blip effects this module does not
 * model (and `a:extLst`) stay as they are.
 */
export function applyAdjustments(
  drawing: XmlElement,
  pic: XmlElement,
  adj: PictureColorAdjustments,
): void {
  ensureDrawingNamespaces(drawing);
  const blip = blipOf(pic);
  if (!blip) throw new Error("The picture has no a:blip.");
  const parts: string[] = [];
  if (adj.transparency !== undefined && adj.transparency > 0) {
    const t = clampPercent(adj.transparency, 0, 100, "transparency");
    parts.push(`<a:alphaModFix amt="${pct(100 - t)}"/>`);
  }
  if (adj.biLevelThreshold !== undefined) {
    parts.push(
      `<a:biLevel thresh="${pct(clampPercent(adj.biLevelThreshold, 0, 100, "threshold"))}"/>`,
    );
  }
  if (adj.transparentColor !== undefined) {
    parts.push(
      `<a:clrChange><a:clrFrom>${srgbMarkup(adj.transparentColor)}</a:clrFrom><a:clrTo>${srgbMarkup(adj.transparentColor, 0)}</a:clrTo></a:clrChange>`,
    );
  }
  if (adj.duotone) {
    parts.push(`<a:duotone>${srgbMarkup(adj.duotone[0])}${srgbMarkup(adj.duotone[1])}</a:duotone>`);
  }
  if (adj.grayscale) parts.push(`<a:grayscl/>`);
  if (adj.saturation !== undefined && adj.saturation !== 100) {
    const sat = clampPercent(adj.saturation, 0, 200, "saturation") - 100;
    parts.push(`<a:hsl hue="0" sat="${pct(sat)}" lum="0"/>`);
  }
  const bright = adj.brightness ?? 0;
  const contrast = adj.contrast ?? 0;
  if (bright !== 0 || contrast !== 0) {
    clampPercent(bright, -100, 100, "brightness");
    clampPercent(contrast, -100, 100, "contrast");
    parts.push(`<a:lum bright="${pct(bright)}" contrast="${pct(contrast)}"/>`);
  }
  const kids = blip.children as XmlNode[];
  const kept = kids.filter(
    (k) => !(k.kind === "element" && k.name.uri === A_NS && BLIP_EFFECTS.has(k.name.local)),
  );
  const extAt = kept.findIndex((k) => k.kind === "element" && k.name.local === "extLst");
  const added = parts.map((p) => fragment(p));
  kept.splice(extAt < 0 ? kept.length : extAt, 0, ...added);
  kids.splice(0, kids.length, ...kept);
}

/** Point the picture's blip at another image relationship. */
export function setBlipRelationship(pic: XmlElement, relId: string): void {
  const blip = blipOf(pic);
  if (!blip) throw new Error("The picture has no a:blip.");
  const attrs = blip.attrs as XmlAttr[];
  const i = attrs.findIndex((a) => a.name.uri === R_NS && a.name.local === "embed");
  const next: XmlAttr = {
    name: { uri: R_NS, local: "embed", prefix: "r" },
    value: relId,
    isNamespaceDecl: false,
  };
  if (i >= 0) attrs[i] = next;
  else attrs.push(next);
}
