/**
 * The Bullet, Numbering, and Multilevel List libraries of Word's Home tab, as
 * list definitions. Bullets use the fonts Word uses (Symbol / Wingdings /
 * Courier New), stored as the private-use characters Word writes.
 */

import type { ListLevel, ListNumberFormat } from "@office-kit/docx";

export interface ListPreset {
  readonly id: string;
  readonly kind: "bullet" | "numbered" | "multilevel";
  readonly levels: readonly ListLevel[];
}

const LEVELS = 9;
const STEP = 360;
const HANGING = 360;

interface Bullet {
  readonly text: string;
  readonly font: string;
}

// Word's default deeper levels of a bullet list: o, ▪, •, repeating.
const DEEPER_BULLETS: readonly Bullet[] = [
  { text: "o", font: "Courier New" },
  { text: "", font: "Wingdings" },
  { text: "", font: "Symbol" },
];

function bulletPreset(id: string, first: Bullet): ListPreset {
  return {
    id,
    kind: "bullet",
    levels: Array.from({ length: LEVELS }, (_, i) => {
      const bullet = i === 0 ? first : (DEEPER_BULLETS[(i - 1) % DEEPER_BULLETS.length] ?? first);
      return {
        format: "bullet" as const,
        text: bullet.text,
        font: bullet.font,
        indentLeft: STEP * 2 * (i + 1),
        hanging: HANGING,
      };
    }),
  };
}

/** The Bullet Library. */
export const BULLET_PRESETS: readonly ListPreset[] = [
  bulletPreset("disc", { text: "", font: "Symbol" }),
  bulletPreset("circle", { text: "o", font: "Courier New" }),
  bulletPreset("square", { text: "", font: "Wingdings" }),
  bulletPreset("diamonds", { text: "", font: "Wingdings" }),
  bulletPreset("arrow", { text: "", font: "Wingdings" }),
  bulletPreset("check", { text: "", font: "Wingdings" }),
  bulletPreset("dash", { text: "–", font: "Calibri" }),
];

// Word's default deeper levels of a numbered list: a. i. 1. a. i. …
const DEEPER_FORMATS: ReadonlyArray<readonly [ListNumberFormat, "left" | "right"]> = [
  ["lowerLetter", "left"],
  ["lowerRoman", "right"],
  ["decimal", "left"],
];

function numberedPreset(
  id: string,
  format: ListNumberFormat,
  suffix: string,
  alignment: "left" | "right" = "left",
  prefix = "",
): ListPreset {
  return {
    id,
    kind: "numbered",
    levels: Array.from({ length: LEVELS }, (_, i) => {
      const [f, a] =
        i === 0
          ? [format, alignment]
          : (DEEPER_FORMATS[(i - 1) % DEEPER_FORMATS.length] ?? [format, alignment]);
      return {
        format: f,
        text: i === 0 ? `${prefix}%1${suffix}` : `%${i + 1}.`,
        alignment: a,
        indentLeft: STEP * 2 * (i + 1),
        hanging: HANGING,
      };
    }),
  };
}

/** The Numbering Library (Word's, plus the Japanese UI's entries). */
export const NUMBERING_PRESETS: readonly ListPreset[] = [
  numberedPreset("decimal-period", "decimal", "."),
  numberedPreset("decimal-paren", "decimal", ")"),
  numberedPreset("upper-roman", "upperRoman", ".", "right"),
  numberedPreset("upper-letter", "upperLetter", "."),
  numberedPreset("lower-letter-paren", "lowerLetter", ")"),
  numberedPreset("lower-letter-period", "lowerLetter", "."),
  numberedPreset("lower-roman", "lowerRoman", ".", "right"),
  numberedPreset("decimal-enclosed-paren", "decimal", ")", "left", "("),
  numberedPreset("enclosed-circle", "decimalEnclosedCircle", ""),
  numberedPreset("japanese-counting", "japaneseCounting", "."),
  numberedPreset("aiueo", "aiueoFullWidth", "."),
  numberedPreset("iroha", "irohaFullWidth", "."),
  numberedPreset("decimal-full-width", "decimalFullWidth", "."),
];

// Word's outline-numbered indents: each level's text further right, the
// label hanging wide enough for "1.1.1.1.".
const OUTLINE_LEFT = [360, 792, 1224, 1728, 2232, 2736, 3240, 3744, 4320];
const OUTLINE_HANGING = [360, 432, 504, 648, 792, 936, 1080, 1224, 1440];
const HEADING_INDENT = 144;

/** The Multilevel List library. */
export const MULTILEVEL_PRESETS: readonly ListPreset[] = [
  {
    id: "outline-decimal",
    kind: "multilevel",
    levels: Array.from({ length: LEVELS }, (_, i) => ({
      format: "decimal" as const,
      text: `${Array.from({ length: i + 1 }, (_, k) => `%${k + 1}`).join(".")}.`,
      indentLeft: OUTLINE_LEFT[i] ?? 0,
      hanging: OUTLINE_HANGING[i] ?? HANGING,
    })),
  },
  {
    id: "outline-paren",
    kind: "multilevel",
    levels: Array.from({ length: LEVELS }, (_, i) => {
      const formats: ListNumberFormat[] = ["decimal", "lowerLetter", "lowerRoman"];
      const format = formats[i % formats.length] ?? "decimal";
      const text = i < 3 ? `%${i + 1})` : i < 6 ? `(%${i + 1})` : `%${i + 1}.`;
      return { format, text, indentLeft: STEP * (i + 1), hanging: HANGING };
    }),
  },
  {
    id: "outline-bullets",
    kind: "multilevel",
    levels: Array.from({ length: LEVELS }, (_, i) => {
      const bullets: Bullet[] = [
        { text: "", font: "Wingdings" },
        { text: "", font: "Wingdings" },
        { text: "", font: "Symbol" },
      ];
      const bullet = bullets[i % bullets.length] ?? { text: "", font: "Symbol" };
      return {
        format: "bullet" as const,
        text: bullet.text,
        font: bullet.font,
        indentLeft: STEP * (i + 1),
        hanging: HANGING,
      };
    }),
  },
  {
    id: "headings-decimal",
    kind: "multilevel",
    levels: Array.from({ length: LEVELS }, (_, i) => ({
      format: "decimal" as const,
      text: Array.from({ length: i + 1 }, (_, k) => `%${k + 1}`).join("."),
      indentLeft: HEADING_INDENT * (i + 3),
      hanging: HEADING_INDENT * (i + 3),
      style: `Heading${i + 1}`,
    })),
  },
  {
    id: "headings-article",
    kind: "multilevel",
    levels: Array.from({ length: LEVELS }, (_, i) => {
      if (i === 0)
        return {
          format: "upperRoman" as const,
          text: "Article %1.",
          indentLeft: 0,
          hanging: 0,
          style: "Heading1",
          suffix: "space" as const,
        };
      if (i === 1)
        return {
          format: "decimalZero" as const,
          text: "Section %1.%2",
          indentLeft: 0,
          hanging: 0,
          style: "Heading2",
          legal: true,
          suffix: "space" as const,
        };
      const formats: ListNumberFormat[] = ["lowerLetter", "lowerRoman", "decimal"];
      return {
        format: formats[(i - 2) % formats.length] ?? "decimal",
        text: `(%${i + 1})`,
        indentLeft: STEP * 2 * (i - 1),
        hanging: HANGING,
        style: `Heading${i + 1}`,
      };
    }),
  },
];

export const ALL_LIST_PRESETS: readonly ListPreset[] = [
  ...BULLET_PRESETS,
  ...NUMBERING_PRESETS,
  ...MULTILEVEL_PRESETS,
];
