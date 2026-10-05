import type { ParamMatcher } from "@sveltejs/kit";
import { DEFAULT_LOCALE, isSiteLocale } from "$lib/i18n/locales";

// English is served without a prefix, so `/en/...` must not match: one page,
// one URL.
export const match = ((param: string) =>
  param !== DEFAULT_LOCALE && isSiteLocale(param)) satisfies ParamMatcher;
