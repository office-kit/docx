/**
 * Measurements in the Home dialogs. Word shows indents in the UI's unit
 * (inches in US English, centimetres elsewhere) and spacing in points; the
 * document stores twips.
 */

import type { LocaleId } from "../../../i18n/index.svelte";

export type LengthUnit = "in" | "cm";

const TWIPS_PER_INCH = 1440;
const CM_PER_INCH = 2.54;
export const TWIPS_PER_POINT = 20;
const DECIMALS = 2;

export function unitFor(locale: LocaleId): LengthUnit {
  return locale === "en" ? "in" : "cm";
}

export function fromTwips(twips: number, unit: LengthUnit): number {
  const inches = twips / TWIPS_PER_INCH;
  return Number((unit === "in" ? inches : inches * CM_PER_INCH).toFixed(DECIMALS));
}

export function toTwips(value: number, unit: LengthUnit): number {
  return Math.round((unit === "in" ? value : value / CM_PER_INCH) * TWIPS_PER_INCH);
}
