/** Length conversions for the page canvas, which lays out in CSS px. */

const CSS_PX_PER_INCH = 96;
const TWIPS_PER_INCH = 1440;
const POINTS_PER_INCH = 72;

export const twipsToPx = (twips: number): number => (twips * CSS_PX_PER_INCH) / TWIPS_PER_INCH;
export const ptToPx = (pt: number): number => (pt * CSS_PX_PER_INCH) / POINTS_PER_INCH;
/** Border widths are in eighths of a point (ST_EighthPointMeasure). */
export const eighthPtToPx = (eighths: number): number => ptToPx(eighths / 8);
