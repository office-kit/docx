/**
 * DrawingML colors (ECMA-376 Part 1, §20.1.2.3): reading the EG_ColorChoice
 * group into a plain RGB hex + opacity, and writing one back as `a:srgbClr`.
 */

import type { XmlElement } from "../xml/index.js";
import { A_NS, attr, elementChildren, escapeXml, numAttr } from "./xml.js";

/** A resolved color: `RRGGBB` and opacity in percent (100 = opaque). */
export interface ResolvedColor {
  readonly hex: string;
  readonly alpha: number;
}

// The Office theme's color scheme, used for `a:schemeClr` when the document's
// theme part is not consulted (the editor renders against the default theme).
const DEFAULT_SCHEME: Readonly<Record<string, string>> = {
  dk1: "000000",
  lt1: "FFFFFF",
  dk2: "44546A",
  lt2: "E7E6E6",
  tx1: "000000",
  bg1: "FFFFFF",
  tx2: "44546A",
  bg2: "E7E6E6",
  accent1: "4472C4",
  accent2: "ED7D31",
  accent3: "A5A5A5",
  accent4: "FFC000",
  accent5: "5B9BD5",
  accent6: "70AD47",
  hlink: "0563C1",
  folHlink: "954F72",
  phClr: "4472C4",
};

// The ST_PresetColorVal names Word itself writes.
const PRESET_COLORS: Readonly<Record<string, string>> = {
  black: "000000",
  white: "FFFFFF",
  red: "FF0000",
  green: "008000",
  blue: "0000FF",
  yellow: "FFFF00",
  gray: "808080",
  grey: "808080",
  silver: "C0C0C0",
  orange: "FFA500",
  purple: "800080",
};

// Color transforms are in 1000ths of a percent (ST_Percentage / ST_PositivePercentage).
const PERCENT = 1000;
const HEX6 = /^[0-9A-Fa-f]{6}$/;

function rgbOf(hex: string): [number, number, number] {
  const n = Number.parseInt(hex, 16);
  return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
}

function clampByte(v: number): number {
  return Math.max(0, Math.min(255, Math.round(v)));
}

function hexOf([r, g, b]: readonly [number, number, number]): string {
  return [r, g, b]
    .map((v) => clampByte(v).toString(16).padStart(2, "0"))
    .join("")
    .toUpperCase();
}

function toHsl([r, g, b]: readonly [number, number, number]): [number, number, number] {
  const [rr, gg, bb] = [r / 255, g / 255, b / 255];
  const max = Math.max(rr, gg, bb);
  const min = Math.min(rr, gg, bb);
  const l = (max + min) / 2;
  if (max === min) return [0, 0, l];
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  let h: number;
  if (max === rr) h = (gg - bb) / d + (gg < bb ? 6 : 0);
  else if (max === gg) h = (bb - rr) / d + 2;
  else h = (rr - gg) / d + 4;
  return [h / 6, s, l];
}

function hue(p: number, q: number, t: number): number {
  const tt = t < 0 ? t + 1 : t > 1 ? t - 1 : t;
  if (tt < 1 / 6) return p + (q - p) * 6 * tt;
  if (tt < 1 / 2) return q;
  if (tt < 2 / 3) return p + (q - p) * (2 / 3 - tt) * 6;
  return p;
}

function fromHsl([h, s, l]: readonly [number, number, number]): [number, number, number] {
  if (s === 0) return [l * 255, l * 255, l * 255];
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  return [hue(p, q, h + 1 / 3) * 255, hue(p, q, h) * 255, hue(p, q, h - 1 / 3) * 255];
}

/**
 * Resolve the first color element under `container` (`a:srgbClr`,
 * `a:schemeClr`, `a:prstClr`, `a:sysClr`, `a:scrgbClr`, `a:hslClr`) with the
 * transforms Word writes (`lumMod`, `lumOff`, `tint`, `shade`, `alpha`).
 */
export function readColor(container: XmlElement | undefined): ResolvedColor | undefined {
  return container ? colorElements(container).map(resolveColor)[0] : undefined;
}

/** The color-choice children of an element, in order (`a:duotone` holds two). */
export function colorElements(container: XmlElement): XmlElement[] {
  return elementChildren(container).filter(
    (c) => c.name.uri === A_NS && c.name.local.endsWith("Clr"),
  );
}

export function resolveColor(el: XmlElement): ResolvedColor | undefined {
  let hex: string | undefined;
  switch (el.name.local) {
    case "srgbClr":
      hex = attr(el, "val");
      break;
    case "schemeClr":
      hex = DEFAULT_SCHEME[attr(el, "val") ?? ""];
      break;
    case "prstClr":
      hex = PRESET_COLORS[attr(el, "val") ?? ""];
      break;
    case "sysClr":
      hex = attr(el, "lastClr") ?? (attr(el, "val") === "window" ? "FFFFFF" : "000000");
      break;
    case "scrgbClr": {
      const c = (k: string): number => ((numAttr(el, k) ?? 0) / (100 * PERCENT)) * 255;
      hex = hexOf([c("r"), c("g"), c("b")]);
      break;
    }
    default:
      hex = undefined;
  }
  if (!hex || !HEX6.test(hex)) return undefined;
  let rgb = rgbOf(hex);
  let alpha = 100;
  for (const t of elementChildren(el)) {
    const v = (numAttr(t, "val") ?? 0) / (100 * PERCENT);
    switch (t.name.local) {
      case "alpha":
        alpha = v * 100;
        break;
      case "lumMod": {
        const [h, s, l] = toHsl(rgb);
        rgb = fromHsl([h, s, Math.min(1, l * v)]);
        break;
      }
      case "lumOff": {
        const [h, s, l] = toHsl(rgb);
        rgb = fromHsl([h, s, Math.max(0, Math.min(1, l + v))]);
        break;
      }
      case "tint":
        rgb = [
          rgb[0] + (255 - rgb[0]) * (1 - v),
          rgb[1] + (255 - rgb[1]) * (1 - v),
          rgb[2] + (255 - rgb[2]) * (1 - v),
        ];
        break;
      case "shade":
        rgb = [rgb[0] * v, rgb[1] * v, rgb[2] * v];
        break;
      default:
        break;
    }
  }
  return { hex: hexOf(rgb), alpha };
}

/** `a:srgbClr` markup for a hex color, with `a:alpha` when not opaque. */
export function srgbMarkup(hex: string, alpha = 100): string {
  if (!HEX6.test(hex)) throw new Error(`Invalid color ${JSON.stringify(hex)}: expected RRGGBB.`);
  const value = escapeXml(hex.toUpperCase());
  if (alpha >= 100) return `<a:srgbClr val="${value}"/>`;
  return `<a:srgbClr val="${value}"><a:alpha val="${Math.round(alpha * PERCENT)}"/></a:srgbClr>`;
}

export function isHexColor(value: string): boolean {
  return HEX6.test(value);
}
