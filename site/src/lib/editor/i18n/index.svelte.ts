/**
 * Lightweight i18n for the editor UI. A single reactive `current` locale drives
 * `t(key)`; because `t` reads the `$state`, any markup calling it re-renders when
 * the locale changes.
 *
 * Ribbon labels and tooltips use the wording of Microsoft Word's own localized
 * UI, so a Word user finds each command under the name they already know.
 * Strings live in one module per feature area under `./messages/`.
 */

import type { Messages } from "./define.js";
import { LOCALES, type LocaleId } from "./locales.js";
import canvas from "./messages/canvas.js";
import core from "./messages/core.js";
import design from "./messages/design.js";
import draw from "./messages/draw.js";
import home from "./messages/home.js";
import insert from "./messages/insert.js";
import layout from "./messages/layout.js";
import mailings from "./messages/mailings.js";
import picture from "./messages/picture.js";
import references from "./messages/references.js";
import review from "./messages/review.js";
import table from "./messages/table.js";
import view from "./messages/view.js";

export { LOCALES, type LocaleId };

const MODULES = [
  core,
  home,
  insert,
  draw,
  design,
  layout,
  references,
  mailings,
  review,
  view,
  table,
  picture,
  canvas,
] as const;

type KeysOf<M> = M extends Messages<infer K> ? K : never;
export type MessageKey = KeysOf<(typeof MODULES)[number]>;

const DICTS = Object.fromEntries(
  Object.keys(LOCALES).map((id) => [
    id,
    Object.assign({}, ...MODULES.map((m) => m[id as LocaleId])),
  ]),
) as Record<LocaleId, Readonly<Record<MessageKey, string>>>;

let current = $state<LocaleId>("en");

/** The active locale (reactive). */
export function locale(): LocaleId {
  return current;
}

export function setLocale(id: LocaleId): void {
  current = id;
}

/** Translate a message key in the active locale. */
export function t(key: MessageKey): string {
  return DICTS[current][key];
}

/**
 * Word's names for the highlight palette (`w:highlight` values), in each
 * locale. `white` is valid OOXML but not in Word's menu, so it has no name.
 */
const HIGHLIGHT_NAMES: Record<LocaleId, Readonly<Record<string, string>>> = {
  en: {
    yellow: "Yellow",
    green: "Bright Green",
    cyan: "Turquoise",
    magenta: "Pink",
    blue: "Blue",
    red: "Red",
    darkBlue: "Dark Blue",
    darkCyan: "Teal",
    darkGreen: "Green",
    darkMagenta: "Violet",
    darkRed: "Dark Red",
    darkYellow: "Dark Yellow",
    darkGray: "Gray-50%",
    lightGray: "Gray-25%",
    black: "Black",
  },
  ja: {
    yellow: "黄",
    green: "明るい緑",
    cyan: "ターコイズ",
    magenta: "ピンク",
    blue: "青",
    red: "赤",
    darkBlue: "濃い青",
    darkCyan: "青緑",
    darkGreen: "緑",
    darkMagenta: "紫",
    darkRed: "濃い赤",
    darkYellow: "濃い黄",
    darkGray: "灰色 50%",
    lightGray: "灰色 25%",
    black: "黒",
  },
  es: {
    yellow: "Amarillo",
    green: "Verde brillante",
    cyan: "Turquesa",
    magenta: "Rosa",
    blue: "Azul",
    red: "Rojo",
    darkBlue: "Azul oscuro",
    darkCyan: "Verde azulado",
    darkGreen: "Verde",
    darkMagenta: "Violeta",
    darkRed: "Rojo oscuro",
    darkYellow: "Amarillo oscuro",
    darkGray: "Gris 50%",
    lightGray: "Gris 25%",
    black: "Negro",
  },
  fr: {
    yellow: "Jaune",
    green: "Vert vif",
    cyan: "Turquoise",
    magenta: "Rose",
    blue: "Bleu",
    red: "Rouge",
    darkBlue: "Bleu foncé",
    darkCyan: "Bleu-vert",
    darkGreen: "Vert",
    darkMagenta: "Violet",
    darkRed: "Rouge foncé",
    darkYellow: "Jaune foncé",
    darkGray: "Gris 50 %",
    lightGray: "Gris 25 %",
    black: "Noir",
  },
  de: {
    yellow: "Gelb",
    green: "Hellgrün",
    cyan: "Türkis",
    magenta: "Rosa",
    blue: "Blau",
    red: "Rot",
    darkBlue: "Dunkelblau",
    darkCyan: "Blaugrün",
    darkGreen: "Grün",
    darkMagenta: "Violett",
    darkRed: "Dunkelrot",
    darkYellow: "Dunkelgelb",
    darkGray: "Grau 50 %",
    lightGray: "Grau 25 %",
    black: "Schwarz",
  },
  zh: {
    yellow: "黄色",
    green: "鲜绿",
    cyan: "青绿",
    magenta: "粉红",
    blue: "蓝色",
    red: "红色",
    darkBlue: "深蓝",
    darkCyan: "青色",
    darkGreen: "绿色",
    darkMagenta: "紫罗兰",
    darkRed: "深红",
    darkYellow: "深黄",
    darkGray: "灰色-50%",
    lightGray: "灰色-25%",
    black: "黑色",
  },
};

/** Word's name for a highlight color in the active locale; the raw value otherwise. */
export function highlightName(color: string): string {
  return HIGHLIGHT_NAMES[current][color] ?? color;
}

/**
 * The names Japanese Word shows for its built-in styles; the document keeps
 * the English name (`w:name`), so the UI translates it. Keys are lowercase,
 * as built-in names are matched without case.
 */
const JA_STYLE_NAMES: Readonly<Record<string, string>> = {
  normal: "標準",
  "no spacing": "行間詰め",
  title: "表題",
  subtitle: "副題",
  "subtle emphasis": "斜体",
  emphasis: "強調斜体",
  "intense emphasis": "強調斜体 2",
  strong: "強調太字",
  quote: "引用文",
  "intense quote": "引用文 2",
  "subtle reference": "参照",
  "intense reference": "参照 2",
  "book title": "書名",
  "list paragraph": "リスト段落",
  caption: "図表番号",
  "toc heading": "目次の見出し",
  header: "ヘッダー",
  footer: "フッター",
  "footnote text": "脚注文字列",
  hyperlink: "ハイパーリンク",
};
const HEADING_NAME = /^heading ([1-9])$/i;

/** A style's name as Word's UI shows it in the active locale. */
export function styleName(name: string): string {
  if (current !== "ja") return name;
  const heading = HEADING_NAME.exec(name);
  if (heading) return `見出し ${heading[1]}`;
  return JA_STYLE_NAMES[name.toLowerCase()] ?? name;
}
