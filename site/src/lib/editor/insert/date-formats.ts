/**
 * Word's Date and Time formats (`\@` pictures) per language, in the order its
 * dialog lists them. Literal text is single-quoted, as Word writes it, so
 * letters like the "d" and "e" of Spanish "de" are not read as date codes.
 */

import type { LocaleId } from "../i18n/locales";

/** The `w:lang` each UI language inserts fields in. */
export const DATE_LANGUAGE: Record<LocaleId, string> = {
  en: "en-US",
  ja: "ja-JP",
  es: "es-ES",
  fr: "fr-FR",
  de: "de-DE",
  zh: "zh-CN",
};

export const DATE_FORMATS: Record<LocaleId, readonly string[]> = {
  en: [
    "M/d/yyyy",
    "dddd, MMMM d, yyyy",
    "MMMM d, yyyy",
    "M/d/yy",
    "yyyy-MM-dd",
    "d-MMM-yy",
    "M.d.yyyy",
    "MMM. d, yy",
    "d MMMM yyyy",
    "MMMM yy",
    "MMM-yy",
    "M/d/yyyy h:mm am/pm",
    "M/d/yyyy h:mm:ss am/pm",
    "h:mm am/pm",
    "h:mm:ss am/pm",
    "HH:mm",
    "HH:mm:ss",
  ],
  ja: [
    "yyyy/MM/dd",
    "yyyy'年'M'月'd'日'",
    "yyyy'年'M'月'd'日' dddd",
    "yyyy/M/d",
    "yy/M/d",
    "M'月'd'日'",
    "yyyy-MM-dd",
    "H:mm",
    "H:mm:ss",
    "H'時'mm'分'",
    "H'時'mm'分'ss'秒'",
  ],
  es: [
    "dd/MM/yyyy",
    "dddd, d' de 'MMMM' de 'yyyy",
    "d' de 'MMMM' de 'yyyy",
    "dd/MM/yy",
    "yyyy-MM-dd",
    "dd.MM.yyyy",
    "MMMM' de 'yyyy",
    "dd/MM/yyyy H:mm",
    "H:mm",
    "H:mm:ss",
  ],
  fr: [
    "dd/MM/yyyy",
    "dddd d MMMM yyyy",
    "d MMMM yyyy",
    "dd/MM/yy",
    "yyyy-MM-dd",
    "dd.MM.yyyy",
    "MMMM yy",
    "dd/MM/yyyy HH:mm",
    "HH:mm",
    "HH:mm:ss",
  ],
  de: [
    "dd.MM.yyyy",
    "dddd, d. MMMM yyyy",
    "d. MMMM yyyy",
    "dd.MM.yy",
    "yyyy-MM-dd",
    "MMMM yy",
    "dd.MM.yyyy HH:mm",
    "HH:mm",
    "HH:mm:ss",
  ],
  zh: [
    "yyyy/M/d",
    "yyyy'年'M'月'd'日'",
    "yyyy'年'M'月'd'日'dddd",
    "yyyy-MM-dd",
    "yyyy.M.d",
    "M'月'd'日'",
    "H:mm",
    "H:mm:ss",
    "H'时'mm'分'",
  ],
};
