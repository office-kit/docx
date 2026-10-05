import { LOCALES as EDITOR_LOCALES } from "$lib/editor/i18n/locales";

/**
 * The languages the docs site is written in: a subset of the editor's, named
 * the same way, so a page and the editor never disagree on a locale's id.
 */
export const LOCALES = {
  en: EDITOR_LOCALES.en,
  ja: EDITOR_LOCALES.ja,
} as const;

export type SiteLocale = keyof typeof LOCALES;

/** Served at the unprefixed URLs; every other locale lives under `/<id>`. */
export const DEFAULT_LOCALE = "en" satisfies SiteLocale;

export function isSiteLocale(value: string | undefined): value is SiteLocale {
  return value !== undefined && Object.hasOwn(LOCALES, value);
}
