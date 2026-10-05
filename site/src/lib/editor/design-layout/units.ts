/**
 * Measurement units for the Layout tab and its dialogs. Word shows lengths
 * in the user's measurement unit (inches in the US English UI, centimetres in
 * most others, millimetres in the Japanese one) and paragraph spacing in
 * points; the document stores twips (1/20 pt).
 */

import type { LocaleId } from "../i18n/locales.js";

export type LengthUnit = "in" | "cm" | "mm" | "pt";

export const TWIPS_PER_INCH = 1440;
export const TWIPS_PER_POINT = 20;
const TWIPS_PER_UNIT: Readonly<Record<LengthUnit, number>> = {
  in: TWIPS_PER_INCH,
  cm: TWIPS_PER_INCH / 2.54,
  mm: TWIPS_PER_INCH / 25.4,
  pt: TWIPS_PER_POINT,
};
const DECIMALS: Readonly<Record<LengthUnit, number>> = { in: 2, cm: 2, mm: 1, pt: 1 };
const SUFFIX: Readonly<Record<LengthUnit, string>> = { in: '"', cm: " cm", mm: " mm", pt: " pt" };
// Spellings Word accepts after a number, by unit.
const UNIT_SPELLINGS: ReadonlyArray<readonly [RegExp, LengthUnit]> = [
  [/^("|in|inch|inches|″)$/i, "in"],
  [/^(cm|センチ)$/i, "cm"],
  [/^(mm|ミリ)$/i, "mm"],
  [/^(pt|point|points|ポイント)$/i, "pt"],
];

/** Word's default length unit for a UI language. */
export function lengthUnitFor(locale: LocaleId): Exclude<LengthUnit, "pt"> {
  if (locale === "en") return "in";
  if (locale === "ja") return "mm";
  return "cm";
}

/** Twips → a number in `unit`, rounded as Word displays it. */
export function fromTwips(twips: number, unit: LengthUnit): number {
  const factor = 10 ** DECIMALS[unit];
  return Math.round((twips / TWIPS_PER_UNIT[unit]) * factor) / factor;
}

/** Twips as Word's text boxes show them, e.g. `1"`, `2.54 cm`, `12 pt`. */
export function formatLength(twips: number, unit: LengthUnit): string {
  return `${fromTwips(twips, unit)}${SUFFIX[unit]}`;
}

/**
 * Parse a length typed into a box: a number with an optional unit
 * (`1.5"`, `3 cm`, `12pt`); a bare number is in `defaultUnit`. Returns
 * twips, or `undefined` when the text is not a length.
 */
export function parseLength(text: string, defaultUnit: LengthUnit): number | undefined {
  const match = /^\s*(-?\d+(?:[.,]\d+)?)\s*([^\d\s]*)\s*$/.exec(text);
  if (!match?.[1]) return undefined;
  const value = Number(match[1].replace(",", "."));
  const suffix = match[2] ?? "";
  const unit = suffix === "" ? defaultUnit : UNIT_SPELLINGS.find(([re]) => re.test(suffix))?.[1];
  if (unit === undefined || !Number.isFinite(value)) return undefined;
  return Math.round(value * TWIPS_PER_UNIT[unit]);
}

/** The step Word's spinner arrows move by, in twips. */
export function spinStep(unit: LengthUnit): number {
  if (unit === "in") return TWIPS_PER_INCH / 10;
  if (unit === "pt") return 6 * TWIPS_PER_POINT;
  return Math.round(TWIPS_PER_UNIT.cm / 10);
}
