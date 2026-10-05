/**
 * Proofing languages offered by the Language dialog and named in the status
 * bar. Names come from the browser (`Intl.DisplayNames`) in the UI locale, so
 * they read as Word's localized language list does.
 */

import type { LocaleId } from "./i18n/index.svelte";

/** Latin-script and complex-script languages (`w:lang/@w:val`, `@w:bidi`). */
export const LATIN_LANGUAGES: readonly string[] = [
  "en-US",
  "en-GB",
  "en-AU",
  "en-CA",
  "fr-FR",
  "fr-CA",
  "de-DE",
  "de-CH",
  "es-ES",
  "es-MX",
  "it-IT",
  "pt-BR",
  "pt-PT",
  "nl-NL",
  "sv-SE",
  "da-DK",
  "nb-NO",
  "fi-FI",
  "pl-PL",
  "cs-CZ",
  "hu-HU",
  "tr-TR",
  "el-GR",
  "ru-RU",
  "uk-UA",
  "ar-SA",
  "he-IL",
  "hi-IN",
  "th-TH",
  "vi-VN",
  "id-ID",
];

/** East Asian languages (`w:lang/@w:eastAsia`). */
export const EAST_ASIAN_LANGUAGES: readonly string[] = ["ja-JP", "zh-CN", "zh-TW", "ko-KR"];

/** The documented default when neither the run nor the dialog names one. */
export const DEFAULT_LANGUAGE = "en-US";

/** A language tag's name in the UI locale, or the tag itself if the browser has no name. */
export function languageName(tag: string, uiLocale: LocaleId): string {
  try {
    return new Intl.DisplayNames([uiLocale], { type: "language" }).of(tag) ?? tag;
  } catch {
    // A legacy hex LCID ("0411") is valid ST_Lang but not a BCP 47 tag.
    return tag;
  }
}
