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
