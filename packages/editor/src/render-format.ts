/**
 * Canvas CSS for the character and paragraph formatting the Home tab edits
 * beyond the basics in render.ts: text effects (caps, outline, emboss …),
 * every underline style, character spacing and position, run and paragraph
 * shading and borders, hidden text, phonetic guides, and list labels.
 */

import { type Docx, readRuby, type WmlParagraph, type WmlRun } from "@office-kit/docx";
import { createListCounter, createNumberingResolver } from "./list-numbering.js";
import type {
  ResolvedBorder,
  ResolvedParagraphFormat,
  ResolvedRunFormat,
  ResolvedShading,
  StyleResolver,
} from "./resolve.js";

const HEX_COLOR = /^[0-9A-Fa-f]{6}$/;
const TWIPS_PER_POINT = 20;
const HALF_POINTS_PER_POINT = 2;
const EIGHTHS_PER_POINT = 8;

// ST_Underline → CSS text-decoration style / thickness.
const UNDERLINE_CSS: Readonly<Record<string, string>> = {
  double: "text-decoration-style:double",
  thick: "text-decoration-thickness:2px",
  dotted: "text-decoration-style:dotted",
  dottedHeavy: "text-decoration-style:dotted;text-decoration-thickness:2px",
  dash: "text-decoration-style:dashed",
  dashedHeavy: "text-decoration-style:dashed;text-decoration-thickness:2px",
  dashLong: "text-decoration-style:dashed",
  dashLongHeavy: "text-decoration-style:dashed;text-decoration-thickness:2px",
  dotDash: "text-decoration-style:dashed",
  dashDotHeavy: "text-decoration-style:dashed;text-decoration-thickness:2px",
  dotDotDash: "text-decoration-style:dotted",
  dashDotDotHeavy: "text-decoration-style:dotted;text-decoration-thickness:2px",
  wave: "text-decoration-style:wavy",
  wavyHeavy: "text-decoration-style:wavy;text-decoration-thickness:2px",
  wavyDouble: "text-decoration-style:wavy",
};

/** CSS for an underline style (the line itself is drawn by render.ts). */
export function underlineStyleCss(style: string | undefined): string | undefined {
  return style === undefined ? undefined : UNDERLINE_CSS[style];
}

function cssColor(hex: string | undefined, fallback: string): string {
  return hex !== undefined && HEX_COLOR.test(hex) ? `#${hex}` : fallback;
}

const PERCENT_PATTERN = /^pct(\d+)$/;
const SOLID_PERCENT = 100;

/**
 * The colour a shading shows: the fill, with a percentage pattern blended in
 * the pattern colour (Word draws `pct15` as 15 % of the pattern colour over
 * the fill). Other patterns (stripes, crosses) show as their fill.
 */
export function shadingCss(shading: ResolvedShading | undefined): string | undefined {
  if (!shading || shading.pattern === "nil") return undefined;
  const percent =
    shading.pattern === "solid"
      ? SOLID_PERCENT
      : Number(PERCENT_PATTERN.exec(shading.pattern)?.[1] ?? 0);
  const fill = HEX_COLOR.test(shading.fill) ? shading.fill : undefined;
  if (percent === 0) return fill === undefined ? undefined : `#${fill}`;
  const over = HEX_COLOR.test(shading.color) ? shading.color : "000000";
  const under = fill ?? "FFFFFF";
  const mix = [0, 2, 4]
    .map((i) => {
      const a = Number.parseInt(over.slice(i, i + 2), 16);
      const b = Number.parseInt(under.slice(i, i + 2), 16);
      return Math.round((a * percent + b * (SOLID_PERCENT - percent)) / SOLID_PERCENT)
        .toString(16)
        .padStart(2, "0");
    })
    .join("");
  return `#${mix}`;
}

const NO_BORDER: ReadonlySet<string> = new Set(["none", "nil"]);
const BORDER_STYLE: Readonly<Record<string, string>> = {
  double: "double",
  triple: "double",
  dotted: "dotted",
  dashed: "dashed",
  dashSmallGap: "dashed",
  dotDash: "dashed",
  dotDotDash: "dotted",
  threeDEmboss: "ridge",
  threeDEngrave: "groove",
  outset: "outset",
  inset: "inset",
};
// Word draws a zero-width line at its thinnest, ¼ pt.
const MIN_BORDER_EIGHTHS = 2;

/** A CSS border for a resolved border line, or `undefined` for none. */
export function borderCss(border: ResolvedBorder | undefined): string | undefined {
  if (!border || NO_BORDER.has(border.style)) return undefined;
  const width = Math.max(border.size, MIN_BORDER_EIGHTHS) / EIGHTHS_PER_POINT;
  const style = BORDER_STYLE[border.style] ?? "solid";
  // A double line needs room for two strokes and a gap.
  const shown = style === "double" ? Math.max(width * 3, 2) : width;
  return `${shown}pt ${style} ${cssColor(border.color, "#000")}`;
}

// The light / dark tones Word uses for Emboss and Engrave.
const RELIEF_LIGHT = "#fff";
const RELIEF_DARK = "#808080";

/** CSS for a run's effects beyond bold / italic / underline / strike / colour / size. */
export function runEffectsCss(fmt: ResolvedRunFormat): string[] {
  const css: string[] = [];
  const t = fmt.toggles;
  if (t.has("caps")) css.push("text-transform:uppercase");
  else if (t.has("smallCaps")) css.push("font-variant-caps:small-caps");
  if (t.has("outline"))
    css.push("-webkit-text-stroke:0.04em currentColor;-webkit-text-fill-color:#fff");
  if (t.has("emboss")) css.push(`text-shadow:-1px -1px 0 ${RELIEF_LIGHT},1px 1px 0 ${RELIEF_DARK}`);
  else if (t.has("imprint"))
    css.push(`text-shadow:1px 1px 0 ${RELIEF_LIGHT},-1px -1px 0 ${RELIEF_DARK}`);
  else if (t.has("shadow")) css.push("text-shadow:1px 1px 1px rgba(0,0,0,.45)");
  // Hidden text shows (dotted-underlined) only while Show ¶ is on: the page
  // sets --wk-hidden-display then.
  if (t.has("vanish"))
    css.push("display:var(--wk-hidden-display,none);text-decoration:underline dotted");
  if (fmt.underlineColor && HEX_COLOR.test(fmt.underlineColor))
    css.push(`text-decoration-color:#${fmt.underlineColor}`);
  if (fmt.spacing) css.push(`letter-spacing:${fmt.spacing / TWIPS_PER_POINT}pt`);
  if (fmt.position) css.push(`position:relative;top:${-fmt.position / HALF_POINTS_PER_POINT}pt`);
  if (!fmt.highlight || fmt.highlight === "none") {
    const shade = shadingCss(fmt.shading);
    if (shade) css.push(`background-color:${shade}`);
  }
  const border = borderCss(fmt.border);
  if (border)
    css.push(`border:${border};box-decoration-break:clone;-webkit-box-decoration-break:clone`);
  return css;
}

const PARAGRAPH_SIDES = ["top", "left", "bottom", "right"] as const;

/** CSS for a paragraph's shading and border box. */
export function paragraphBoxCss(fmt: ResolvedParagraphFormat): string[] {
  const css: string[] = [];
  const shade = shadingCss(fmt.shading);
  if (shade) css.push(`background-color:${shade}`);
  for (const side of PARAGRAPH_SIDES) {
    const border = fmt.borders[side];
    const line = borderCss(border);
    if (!line || !border) continue;
    css.push(`border-${side}:${line}`, `padding-${side}:${border.space}pt`);
  }
  return css;
}

function escapeHtml(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

/**
 * Escape run text, giving each tab its own element: a tab is as wide as the
 * gap to its tab stop, which only the canvas can measure (see tab-layout).
 * The element is not editable, so the caret steps over a tab as one
 * character and typing never lands inside it.
 */
export function textHtml(text: string): string {
  return text
    .split("\t")
    .map(escapeHtml)
    .join('<span class="wk-tab" contenteditable="false">\t</span>');
}

/** The tab stops a paragraph's tabs lay out against, read back by `layoutTabStops`. */
export interface TabStopSpec {
  /** Twips between implicit stops. */
  readonly interval: number;
  readonly stops: ResolvedParagraphFormat["tabs"];
  /** A hanging indent is a stop too: the first tab of the line jumps to it. */
  readonly hangingIndent?: number | undefined;
}

export function tabStopsAttr(fmt: ResolvedParagraphFormat, interval: number): string {
  const spec: TabStopSpec = {
    interval,
    stops: fmt.tabs,
    ...(fmt.hanging ? { hangingIndent: fmt.left ?? 0 } : {}),
  };
  return ` data-wk-tabs="${escapeHtml(JSON.stringify(spec))}"`;
}

/** Inner HTML for runs the plain text path cannot show (a phonetic guide), else `undefined`. */
export function specialRunHtml(run: WmlRun): string | undefined {
  if (!run.pieces.some((p) => p.kind === "raw")) return undefined;
  const ruby = readRuby(run);
  if (!ruby) return undefined;
  const rtSize = ruby.rubySizeHalfPoints / HALF_POINTS_PER_POINT;
  return `<ruby contenteditable="false">${escapeHtml(ruby.base)}<rt style="font-size:${rtSize}pt">${escapeHtml(ruby.ruby)}</rt></ruby>`;
}

/**
 * The label of a list paragraph, drawn where Word draws it: at the first-line
 * position, padded to the hanging indent so the text starts at the indent.
 */
export function listLabelHtml(label: string, fmt: ResolvedParagraphFormat, runCss: string): string {
  const hanging = (fmt.hanging ?? 0) / TWIPS_PER_POINT;
  const css = [runCss, `display:inline-block`, `min-width:${hanging}pt`, `text-indent:0`];
  return `<span class="wk-list-label" contenteditable="false" style="${escapeHtml(css.join(";"))}">${escapeHtml(label)}</span>`;
}

/** The style resolver a render uses, plus the list counter for that render. */
export interface RenderResolver extends StyleResolver {
  /** The number or bullet of a list paragraph; call once per paragraph, in document order. */
  listLabel(para: WmlParagraph): string | undefined;
}

/** Wrap a resolver with a fresh list counter for one render of `doc`. */
export function renderResolver(doc: Docx, styles: StyleResolver): RenderResolver {
  const count = createListCounter(createNumberingResolver(doc));
  return {
    ...styles,
    listLabel(para) {
      const numbering = styles.paragraph(para).numbering;
      if (!numbering) return undefined;
      const label = count(numbering.numId, numbering.ilvl);
      if (!label) return undefined;
      return label.text;
    },
  };
}
