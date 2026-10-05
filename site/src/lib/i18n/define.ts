import type { SiteLocale } from "./locales.js";

/** One page's (or feature's) text, in every locale. */
export type Messages<T> = Readonly<Record<SiteLocale, T>>;

/** Text kept beside English source data, for every locale other than English. */
export type Translations<T> = Readonly<Record<Exclude<SiteLocale, "en">, T>>;

/**
 * Declare an area's messages, the way the editor's `defineMessages` does:
 * English defines the shape, and every other locale must supply exactly that
 * shape, so a missing translation is a type error rather than English leaking
 * into a localized page. Values may be functions when text needs arguments.
 */
export function defineMessages<T>(
  messages: { en: T } & Record<Exclude<SiteLocale, "en">, NoInfer<T>>,
): Messages<T> {
  return messages;
}
