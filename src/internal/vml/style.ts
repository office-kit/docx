/**
 * VML's `style` attribute (ECMA-376 Part 4, §19.1.2.19 and the CSS2 subset in
 * §14.1.2.19): a `;`-separated list of `name:value` declarations carrying the
 * shape's box (`width`, `margin-left` …), its anchoring
 * (`mso-position-horizontal-relative` …), z-order and rotation.
 */

/** Parse a VML style into an insertion-ordered map. Unknown names are kept. */
export function parseVmlStyle(style: string | undefined): Map<string, string> {
  const out = new Map<string, string>();
  if (!style) return out;
  for (const decl of style.split(";")) {
    const colon = decl.indexOf(":");
    if (colon < 0) continue;
    const name = decl.slice(0, colon).trim().toLowerCase();
    const value = decl.slice(colon + 1).trim();
    if (name) out.set(name, value);
  }
  return out;
}

export function serializeVmlStyle(style: ReadonlyMap<string, string>): string {
  return [...style].map(([name, value]) => `${name}:${value}`).join(";");
}

const POINTS_PER_UNIT: Readonly<Record<string, number>> = {
  pt: 1,
  in: 72,
  cm: 72 / 2.54,
  mm: 72 / 25.4,
  pc: 12,
  px: 0.75,
  // EMU are not a CSS unit, but Word occasionally writes them in VML.
  emu: 1 / 12700,
};
const LENGTH = /^(-?[0-9]*\.?[0-9]+)\s*([a-z]*)$/i;

/**
 * A CSS length in points. A unitless number is in pixels (CSS2), except in a
 * group's children, whose coordinates are unitless numbers in the group's
 * coordinate space: callers pass `unitless: "raw"` there.
 */
export function lengthToPoints(
  value: string | undefined,
  unitless: "px" | "raw" = "px",
): number | undefined {
  if (value === undefined) return undefined;
  const m = LENGTH.exec(value.trim());
  if (!m) return undefined;
  const n = Number(m[1]);
  const unit = (m[2] ?? "").toLowerCase();
  if (unit === "") return unitless === "raw" ? n : n * 0.75;
  const factor = POINTS_PER_UNIT[unit];
  return factor === undefined ? undefined : n * factor;
}

/** Format points as a compact VML length (`12.5pt`). */
export function points(value: number): string {
  return `${Math.round(value * 100) / 100}pt`;
}
