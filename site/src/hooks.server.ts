import type { Handle } from "@sveltejs/kit";
import { DEFAULT_LOCALE, isSiteLocale } from "$lib/i18n/locales";

// `<html lang>` sits outside the Svelte tree, so prerendering fills it in here.
export const handle: Handle = ({ event, resolve }) => {
  const { lang } = event.params;
  const id = isSiteLocale(lang) ? lang : DEFAULT_LOCALE;
  return resolve(event, { transformPageChunk: ({ html }) => html.replace("%lang%", id) });
};
