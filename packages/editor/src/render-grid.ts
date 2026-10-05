/**
 * The document grid (`w:docGrid`, §17.6.5) and East Asian paragraph settings
 * as CSS: what Japanese and Chinese Word documents rely on for their look.
 *
 * - With a line grid, lines snap to whole multiples of the line pitch, so a
 *   10.5 pt line takes one 18 pt grid line and a heading two.
 * - With a character grid, East Asian characters sit on the character pitch.
 *   The canvas spaces every character by the difference, where Word compresses
 *   Latin text: a full-width line still holds the grid's characters.
 * - Indents in characters (`leftChars` …) and spacing in lines
 *   (`beforeLines` …) are converted with that pitch.
 */

import type { DocumentGrid } from "./layout/sections.js";
import type { ResolvedParagraphFormat, ResolvedRunFormat } from "./resolve.js";

const TWIPS_PER_POINT = 20;
const HUNDREDTHS = 100;
const SIXTEENTHS_OF_POINT = 4096;
// The line unit of `beforeLines` / `afterLines` without a line grid: Word's
// single line, 12 pt.
const DEFAULT_LINE_TWIPS = 240;
// Word's single line for a font, as a multiple of its size: the font's own
// ascent + descent. Measured in Word for Mac 16 (游明朝 10.5 pt: 15.2 pt).
// It decides when a size takes two grid lines (16 pt does, 10.5 pt does not).
const FONT_LINE_HEIGHT: Readonly<Record<string, number>> = {
  游明朝: 1.447,
  "游明朝 Demibold": 1.447,
  "Yu Mincho": 1.447,
  游ゴシック: 1.447,
  "游ゴシック Light": 1.447,
  "Yu Gothic": 1.447,
  "Yu Gothic Light": 1.447,
};
// Word's single spacing for its default Latin fonts (Calibri, Cambria …).
const DEFAULT_LINE_HEIGHT = 1.2;

/**
 * The single line height of a paragraph mark's fonts, as a multiple of the
 * size; `known` is false when the fonts are not in the table.
 */
export function fontLineHeight(mark: ResolvedRunFormat): { ratio: number; known: boolean } {
  const ratios = [mark.font, mark.eastAsiaFont]
    .map((f) => (f === undefined ? undefined : FONT_LINE_HEIGHT[f]))
    .filter((r): r is number => r !== undefined);
  return ratios.length > 0
    ? { ratio: Math.max(...ratios), known: true }
    : { ratio: DEFAULT_LINE_HEIGHT, known: false };
}
const AUTO_LINE_UNIT = 240;

/** The grid pitches that apply to a paragraph, in points. */
export interface GridPitch {
  /** Set when lines snap to the grid. */
  readonly line?: number | undefined;
  /** Set when characters sit on a character grid. */
  readonly char?: number | undefined;
}

/** A section's grid as pitches, given the document's base font size in points. */
export function gridPitch(grid: DocumentGrid | undefined, baseSize: number): GridPitch {
  if (!grid || grid.type === "default") return {};
  const line =
    (grid.type === "lines" || grid.type === "linesAndChars") && grid.linePitch
      ? grid.linePitch / TWIPS_PER_POINT
      : undefined;
  const char =
    grid.type === "linesAndChars" || grid.type === "snapToChars"
      ? baseSize + (grid.charSpace ?? 0) / SIXTEENTHS_OF_POINT
      : undefined;
  return { line, char };
}

/** Indents and spacing in points, with character / line units converted. */
export function paragraphMetrics(
  fmt: ResolvedParagraphFormat,
  pitch: GridPitch,
  fontSize: number,
): {
  left: number | undefined;
  right: number | undefined;
  firstLine: number | undefined;
  hanging: number | undefined;
  before: number;
  after: number;
} {
  const char = pitch.char ?? fontSize;
  const line = pitch.line ?? DEFAULT_LINE_TWIPS / TWIPS_PER_POINT;
  const pick = (twips: number | undefined, units: number | undefined, unit: number) =>
    units !== undefined && units !== 0
      ? (units / HUNDREDTHS) * unit
      : twips === undefined
        ? undefined
        : twips / TWIPS_PER_POINT;
  const hanging = pick(fmt.hanging, fmt.hangingChars, char);
  return {
    left: pick(fmt.left, fmt.leftChars, char),
    right: pick(fmt.right, fmt.rightChars, char),
    hanging,
    firstLine: hanging === undefined ? pick(fmt.firstLine, fmt.firstLineChars, char) : undefined,
    before: pick(fmt.before, fmt.beforeLines, line) ?? 0,
    after: pick(fmt.after, fmt.afterLines, line) ?? 0,
  };
}

/**
 * The `line-height` for text in a run's font within a paragraph, or
 * `undefined` to inherit. Set on every run (and the paragraph, for its mark),
 * so a line is as tall as its tallest text, as in Word:
 *
 * - on a line grid, the font's single line times an auto multiple, rounded up
 *   to whole grid lines (§17.6.5);
 * - otherwise, auto spacing as a multiple of the font's single line.
 *
 * Exact and at-least spacing is the paragraph's own fixed height.
 */
export function lineHeightCss(
  fmt: ResolvedParagraphFormat,
  pitch: GridPitch,
  run: ResolvedRunFormat,
): string | undefined {
  if (fmt.lineRule !== undefined && fmt.lineRule !== "auto") return undefined;
  const size = (run.sizeHalfPoints ?? DEFAULT_SIZE_HALF_POINTS) / 2;
  const single = fontLineHeight(run);
  const multiple = fmt.line ? fmt.line / AUTO_LINE_UNIT : 1;
  if (pitch.line !== undefined && fmt.toggles.snapToGrid) {
    const lines = Math.max(1, Math.ceil((size * single.ratio * multiple) / pitch.line));
    return `line-height:${lines * pitch.line}pt`;
  }
  if (!fmt.line && !single.known) return undefined;
  return `line-height:${single.ratio * multiple}`;
}

// w:sz when nothing sets it (§17.3.2.38).
const DEFAULT_SIZE_HALF_POINTS = 20;

/** Spacing that puts each character on the character grid. */
export function gridLetterSpacing(
  fmt: ResolvedParagraphFormat,
  pitch: GridPitch,
  fontSize: number,
): number | undefined {
  if (pitch.char === undefined || !fmt.toggles.snapToGrid) return undefined;
  const extra = pitch.char - fontSize;
  return extra === 0 ? undefined : extra;
}

/**
 * East Asian line breaking and spacing (§17.3.1.16 kinsoku, §17.3.1.45
 * wordWrap, §17.3.1.21 overflowPunct, §17.3.1.43 topLinePunct, §17.3.1.2–3
 * autoSpaceDE / DN) as their CSS Text 4 counterparts.
 */
export function eastAsianTypographyCss(fmt: ResolvedParagraphFormat): string[] {
  const { kinsoku, wordWrap, overflowPunct, topLinePunct, autoSpaceDE, autoSpaceDN } = fmt.toggles;
  // Chrome knows only `normal` (both kinds of spacing) and `no-autospace`;
  // a browser that has the finer values takes the later declaration.
  const autospace =
    autoSpaceDE && autoSpaceDN
      ? ["text-autospace:normal"]
      : autoSpaceDE || autoSpaceDN
        ? [
            "text-autospace:normal",
            `text-autospace:${autoSpaceDE ? "ideograph-alpha" : "ideograph-numeric"}`,
          ]
        : ["text-autospace:no-autospace"];
  return [
    // Word's standard kinsoku set keeps closing punctuation off a line start
    // but lets small kana begin one, which is CSS's `normal`.
    `line-break:${kinsoku ? "normal" : "anywhere"}`,
    ...(wordWrap ? [] : ["word-break:break-all"]),
    ...(overflowPunct ? ["hanging-punctuation:allow-end"] : []),
    `text-spacing-trim:${topLinePunct ? "trim-start" : "space-all"}`,
    ...autospace,
  ];
}
