export const LOCALES = {
  en: "English",
  ja: "日本語",
  es: "Español",
  fr: "Français",
  de: "Deutsch",
  zh: "中文",
} as const;

export type LocaleId = keyof typeof LOCALES;
