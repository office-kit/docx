/**
 * i18n for the docs site. Unlike the editor, whose locale is a UI setting, a
 * page's locale is part of its URL (`/ja/docs/...`), so every page can be
 * prerendered and linked in each language. `locale()` reads the route param;
 * because `page` is reactive, markup calling it follows client navigation.
 */

import { base } from "$app/paths";
import { page } from "$app/state";
import type { Messages, Translations } from "./define.js";
import { DEFAULT_LOCALE, isSiteLocale, LOCALES, type SiteLocale } from "./locales.js";

export { defineMessages, type Messages, type Translations } from "./define.js";
export { DEFAULT_LOCALE, isSiteLocale, LOCALES, type SiteLocale };

// Every translated page lives under this optional segment; `/editor` and the
// llms files sit outside it and exist in English only.
const LOCALIZED_ROUTE = "/[[lang=lang]]";

const STORAGE_KEY = "office-kit-lang";

/** The active page's locale (reactive). */
export function locale(): SiteLocale {
  const { lang } = page.params;
  return isSiteLocale(lang) ? lang : DEFAULT_LOCALE;
}

/** The active locale's entry of an area's messages. */
export function localized<T>(messages: Messages<T>): T {
  return messages[locale()];
}

/** Source data written in English, or its translation in the active locale. */
export function translate<T>(english: T, translations: Translations<T>): T {
  const id = locale();
  return id === DEFAULT_LOCALE ? english : translations[id];
}

/**
 * `path` (site-absolute, no base) as it is spelled in locale `id`. Paths with
 * a file extension, such as `/llms.txt`, are files rather than pages and are
 * not translated, so they keep one URL.
 */
export function localizePath(path: string, id: SiteLocale): string {
  if (id === DEFAULT_LOCALE || /\.\w+$/.test(path)) return path;
  return path === "/" ? `/${id}` : `/${id}${path}`;
}

/** A link to `path` in the active locale, base included. */
export function href(path: string): string {
  return `${base}${localizePath(path, locale())}`;
}

/**
 * The active page's path with no locale prefix, or `undefined` when the page
 * has no translations. Derived from the route id, not the URL, because during
 * prerendering `base` is relative and cannot be stripped from a pathname;
 * no localized route has a param besides the locale, so the id is the path.
 */
export function routePath(): string | undefined {
  const id = page.route.id;
  if (!id?.startsWith(LOCALIZED_ROUTE)) return undefined;
  return id.slice(LOCALIZED_ROUTE.length) || "/";
}

// Storage can be unavailable (private mode, blocked site data). The choice is
// a convenience, so the site then just behaves as if none was made.
export function storedLocale(): SiteLocale | undefined {
  try {
    const stored = localStorage.getItem(STORAGE_KEY) ?? undefined;
    return isSiteLocale(stored) ? stored : undefined;
  } catch {
    return undefined;
  }
}

export function storeLocale(id: SiteLocale): void {
  try {
    localStorage.setItem(STORAGE_KEY, id);
  } catch {
    // See `storedLocale`.
  }
}
