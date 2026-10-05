/** Outline levels as Word's Outline view and Navigation pane see them. */

import { getParagraphProp, getParagraphStyle, type WmlParagraph } from "@office-kit/docx";

/** Outline levels 1–9 are headings; Word calls the rest Body Text (level 10). */
export const BODY_TEXT_LEVEL = 10;
export const MAX_HEADING_LEVEL = 9;
const HEADING_STYLE = /^Heading([1-9])$/;

/**
 * A paragraph's outline level as the Outline view shows it: its Heading
 * style's level, else a direct `w:outlineLvl` (0-based in the file), else
 * Body Text.
 */
export function outlineLevelOf(para: WmlParagraph): number {
  const style = HEADING_STYLE.exec(getParagraphStyle(para) ?? "");
  if (style) return Number(style[1]);
  const direct = Number(getParagraphProp(para, "outlineLvl").val);
  return Number.isInteger(direct) && direct >= 0 && direct < MAX_HEADING_LEVEL
    ? direct + 1
    : BODY_TEXT_LEVEL;
}
