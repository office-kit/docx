/** Choices shared by the References dialogs, shown the way Word's drop-downs show them. */

import type { NumberingFormat, TabLeader } from "@office-kit/docx";

export const LEADERS: ReadonlyArray<{ readonly value: TabLeader; readonly sample: string }> = [
  { value: "none", sample: "" },
  { value: "dot", sample: "........" },
  { value: "hyphen", sample: "--------" },
  { value: "underscore", sample: "________" },
];

/** Number formats offered for captions and notes, with Word's sample text. */
export const NUMBER_FORMATS: ReadonlyArray<{ readonly value: NumberingFormat; readonly sample: string }> = [
  { value: "decimal", sample: "1, 2, 3, …" },
  { value: "lowerLetter", sample: "a, b, c, …" },
  { value: "upperLetter", sample: "A, B, C, …" },
  { value: "lowerRoman", sample: "i, ii, iii, …" },
  { value: "upperRoman", sample: "I, II, III, …" },
  { value: "chicago", sample: "*, †, ‡, §, …" },
];
