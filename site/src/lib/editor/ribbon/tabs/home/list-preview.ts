/** Text for the list library tiles. */

import type { ListLevel } from "@office-kit/docx";
import { bulletGlyph, formatListNumber } from "@office-kit/docx-editor";

const LEVEL_REF = /%([1-9])/g;
const PREVIEW_ITEMS = 3;

/** The first three items of a list level, as Word's library tiles show them. */
export function previewLabels(level: ListLevel | undefined): string[] {
  if (!level) return [];
  return Array.from({ length: PREVIEW_ITEMS }, (_, i) =>
    level.format === "bullet"
      ? bulletGlyph(level.text)
      : level.text.replace(LEVEL_REF, () => formatListNumber(i + 1, level.format)),
  );
}

/** The first item of each of the top three levels, for the Multilevel List gallery. */
export function multilevelPreview(levels: readonly ListLevel[]): string[] {
  return levels
    .slice(0, PREVIEW_ITEMS)
    .map((level) =>
      level.format === "bullet"
        ? bulletGlyph(level.text)
        : level.text.replace(LEVEL_REF, (_, k: string) =>
            formatListNumber(1, levels[Number(k) - 1]?.format ?? "decimal"),
          ),
    );
}
