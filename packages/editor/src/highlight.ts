/**
 * Display colors for `<w:highlight>`. The value list itself is
 * `HIGHLIGHT_COLORS` from `@office-kit/docx`, whose writers enforce it.
 */

import type { HighlightColor } from "@office-kit/docx";

/**
 * Display RGB for each named highlight. ECMA-376 names the colors but gives no
 * RGB; these are the values Word renders.
 */
const HIGHLIGHT_RGB: Readonly<Record<Exclude<HighlightColor, "none">, string>> = {
  yellow: "FFFF00",
  green: "00FF00",
  cyan: "00FFFF",
  magenta: "FF00FF",
  blue: "0000FF",
  red: "FF0000",
  darkBlue: "000080",
  darkCyan: "008080",
  darkGreen: "008000",
  darkMagenta: "800080",
  darkRed: "800000",
  darkYellow: "808000",
  darkGray: "808080",
  lightGray: "C0C0C0",
  black: "000000",
  white: "FFFFFF",
};

const RGB_BY_NAME: ReadonlyMap<string, string> = new Map(Object.entries(HIGHLIGHT_RGB));

/**
 * CSS color for a stored `w:val`, or `undefined` for `"none"` and for values
 * outside ST_HighlightColor (found in foreign files): those are not guessed.
 */
export function highlightCss(value: string): string | undefined {
  const rgb = RGB_BY_NAME.get(value);
  return rgb === undefined ? undefined : `#${rgb}`;
}
