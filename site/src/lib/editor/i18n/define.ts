import type { LocaleId } from "./locales.js";

/** One feature area's strings, in every locale. */
export type Messages<K extends string> = Readonly<Record<LocaleId, Readonly<Record<K, string>>>>;

/**
 * Declare a feature area's messages. English defines the key set; every other
 * locale must translate exactly those keys, so a missing translation is a type
 * error rather than an English string leaking into a localized ribbon.
 */
export function defineMessages<K extends string>(
  messages: { en: Record<K, string> } & Record<Exclude<LocaleId, "en">, Record<NoInfer<K>, string>>,
): Messages<K> {
  return messages;
}
