/**
 * List labels: which `<w:lvl>` a numbered paragraph uses, and the label text
 * Word draws for it ("1.", "a)", "①", "•" …), counted in document order.
 */

import {
  childElementsOf,
  type Docx,
  getElementAttr,
  numberingPart,
  type XmlElement,
} from "@office-kit/docx";

function child(el: XmlElement | undefined, local: string): XmlElement | undefined {
  return el && childElementsOf(el).find((c) => c.name.local === local);
}

function val(el: XmlElement | undefined, local: string): string | undefined {
  const c = child(el, local);
  return c && getElementAttr(c, "val");
}

function intVal(el: XmlElement | undefined, local: string): number | undefined {
  const n = Number(val(el, local));
  return Number.isInteger(n) ? n : undefined;
}

/** A level of a list instance as it applies: the definition plus any override. */
export interface ListLevelInfo {
  readonly lvl: XmlElement;
  readonly format: string;
  readonly text: string;
  readonly start: number;
  /** The `abstractNumId` the instance points at. */
  readonly abstractNumId: number;
}

export interface NumberingResolver {
  /** The level `ilvl` of list instance `numId`, or `undefined` when either is missing. */
  level(numId: number, ilvl: number): ListLevelInfo | undefined;
}

const MAX_LEVELS = 9;

/** Index the numbering part once, for repeated lookups during one render. */
export function createNumberingResolver(doc: Docx): NumberingResolver {
  const part = numberingPart(doc);
  const abstracts = new Map<number, XmlElement>();
  for (const a of part?.abstractNums ?? []) {
    const id = Number(getElementAttr(a, "abstractNumId"));
    if (Number.isInteger(id)) abstracts.set(id, a);
  }
  const nums = new Map<number, XmlElement>();
  for (const n of part?.nums ?? []) {
    const id = Number(getElementAttr(n, "numId"));
    if (Number.isInteger(id)) nums.set(id, n);
  }
  const levelOf = (container: XmlElement | undefined, ilvl: number): XmlElement | undefined =>
    container &&
    childElementsOf(container).find(
      (c) => c.name.local === "lvl" && Number(getElementAttr(c, "ilvl") ?? "0") === ilvl,
    );
  return {
    level(numId, ilvl) {
      if (ilvl < 0 || ilvl >= MAX_LEVELS) return undefined;
      const num = nums.get(numId);
      const abstractNumId = intVal(num, "abstractNumId");
      if (abstractNumId === undefined) return undefined;
      const override = num
        ? childElementsOf(num).find(
            (c) => c.name.local === "lvlOverride" && Number(getElementAttr(c, "ilvl")) === ilvl,
          )
        : undefined;
      let abstract = abstracts.get(abstractNumId);
      // A list bound to a numbering style takes its levels from the style's
      // definition (numStyleLink → the abstractNum with that styleLink).
      const styleLink = val(abstract, "numStyleLink");
      if (styleLink !== undefined) {
        for (const a of abstracts.values()) if (val(a, "styleLink") === styleLink) abstract = a;
      }
      const lvl = levelOf(override, ilvl) ?? levelOf(abstract, ilvl);
      if (!lvl) return undefined;
      return {
        lvl,
        format: val(lvl, "numFmt") ?? "decimal",
        text: val(lvl, "lvlText") ?? "",
        start: intVal(override, "startOverride") ?? intVal(lvl, "start") ?? 1,
        abstractNumId,
      };
    },
  };
}

const ROMAN: ReadonlyArray<readonly [number, string]> = [
  [1000, "m"],
  [900, "cm"],
  [500, "d"],
  [400, "cd"],
  [100, "c"],
  [90, "xc"],
  [50, "l"],
  [40, "xl"],
  [10, "x"],
  [9, "ix"],
  [5, "v"],
  [4, "iv"],
  [1, "i"],
];

function roman(n: number): string {
  let rest = n;
  let out = "";
  for (const [value, digits] of ROMAN) {
    while (rest >= value) {
      out += digits;
      rest -= value;
    }
  }
  return out;
}

const ALPHABET_SIZE = 26;
const A_CODE = 97;

/** a … z, aa … zz, aaa …, as Word letters lists. */
function letters(n: number): string {
  const letter = String.fromCodePoint(A_CODE + ((n - 1) % ALPHABET_SIZE));
  return letter.repeat(Math.floor((n - 1) / ALPHABET_SIZE) + 1);
}

const KANJI_DIGITS = ["〇", "一", "二", "三", "四", "五", "六", "七", "八", "九"];

/** Japanese counting (一, 十, 二十一 …) for 1–99, Word's japaneseCounting. */
function japaneseCounting(n: number): string {
  if (n < 10) return KANJI_DIGITS[n] ?? String(n);
  if (n >= 100) return String(n);
  const tens = Math.floor(n / 10);
  const ones = n % 10;
  return `${tens === 1 ? "" : KANJI_DIGITS[tens]}十${ones === 0 ? "" : KANJI_DIGITS[ones]}`;
}

const AIUEO =
  "アイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワヲン";
const IROHA =
  "イロハニホヘトチリヌルヲワカヨタレソツネナラムウヰノオクヤマケフコエテアサキユメミシヱヒモセス";
const TRADITIONAL = "甲乙丙丁戊己庚辛壬癸";
const FULL_WIDTH_OFFSET = 0xfee0;
const CIRCLED_ONE = 0x2460;
const CIRCLED_TWENTY_ONE = 0x3251;
const CIRCLED_THIRTY_SIX = 0x32b1;
const MAX_CIRCLED = 50;
const CIRCLED_TWENTY = 20;
const CIRCLED_THIRTY_FIVE = 35;

function fullWidth(text: string): string {
  return [...text]
    .map((c) => String.fromCodePoint((c.codePointAt(0) ?? 0) + FULL_WIDTH_OFFSET))
    .join("");
}

function circled(n: number): string {
  if (n >= 1 && n <= CIRCLED_TWENTY) return String.fromCodePoint(CIRCLED_ONE + n - 1);
  if (n > CIRCLED_TWENTY && n <= CIRCLED_THIRTY_FIVE)
    return String.fromCodePoint(CIRCLED_TWENTY_ONE + n - CIRCLED_TWENTY - 1);
  if (n > CIRCLED_THIRTY_FIVE && n <= MAX_CIRCLED)
    return String.fromCodePoint(CIRCLED_THIRTY_SIX + n - CIRCLED_THIRTY_FIVE - 1);
  return String(n);
}

function cyclic(chars: string, n: number): string {
  const list = [...chars];
  return list[(n - 1) % list.length] ?? String(n);
}

const ORDINAL_SUFFIX = (n: number): string => {
  const mod100 = n % 100;
  if (mod100 >= 11 && mod100 <= 13) return "th";
  return ["th", "st", "nd", "rd"][n % 10] ?? "th";
};

/** A counter value in an ST_NumberFormat, as Word displays it. */
export function formatListNumber(n: number, format: string): string {
  switch (format) {
    case "upperRoman":
      return roman(n).toUpperCase();
    case "lowerRoman":
      return roman(n);
    case "upperLetter":
      return letters(n).toUpperCase();
    case "lowerLetter":
      return letters(n);
    case "decimalZero":
      return n < 10 ? `0${n}` : String(n);
    case "ordinal":
      return `${n}${ORDINAL_SUFFIX(n)}`;
    case "decimalFullWidth":
    case "decimalFullWidth2":
      return fullWidth(String(n));
    case "decimalEnclosedCircle":
    case "decimalEnclosedCircleChinese":
      return circled(n);
    case "decimalEnclosedParen":
      return `(${n})`;
    case "decimalEnclosedFullstop":
      return `${n}.`;
    case "japaneseCounting":
    case "chineseCounting":
    case "chineseCountingThousand":
    case "japaneseLegal":
    case "ideographDigital":
      return japaneseCounting(n);
    case "aiueo":
    case "aiueoFullWidth":
      return cyclic(AIUEO, n);
    case "iroha":
    case "irohaFullWidth":
      return cyclic(IROHA, n);
    case "ideographTraditional":
      return cyclic(TRADITIONAL, n);
    case "none":
      return "";
    default:
      return String(n);
  }
}

// Bullet characters stored in the Symbol / Wingdings private-use area, and
// what they look like, so a browser without those fonts draws the right mark.
const PUA_BULLETS: Readonly<Record<string, string>> = {
  "": "•",
  "": "▪",
  "": "➢",
  "": "✓",
  "": "❖",
  "": "◆",
  "": "■",
  "": "➔",
  "": "–",
};

/** A bullet's visible character. */
export function bulletGlyph(text: string): string {
  return [...text].map((c) => PUA_BULLETS[c] ?? c).join("");
}

/**
 * Counts list items in document order and produces each one's label. Feed
 * every paragraph in order; non-list paragraphs do not reset the count (Word
 * continues a list across interruptions).
 */
export function createListCounter(resolver: NumberingResolver) {
  const counters = new Map<number, number[]>();
  return (numId: number, ilvl: number): { text: string; level: ListLevelInfo } | undefined => {
    const level = resolver.level(numId, ilvl);
    if (!level) return undefined;
    const counts = counters.get(numId) ?? [];
    counters.set(numId, counts);
    counts[ilvl] = counts[ilvl] === undefined ? level.start : counts[ilvl] + 1;
    // A deeper level restarts after a higher one.
    counts.length = ilvl + 1;
    if (level.format === "bullet") return { text: bulletGlyph(level.text), level };
    const legal = childElementsOf(level.lvl).some((c) => c.name.local === "isLgl");
    const text = level.text.replace(/%([1-9])/g, (_, d: string) => {
      const index = Number(d) - 1;
      const other = index === ilvl ? level : resolver.level(numId, index);
      const value = counts[index] ?? other?.start ?? 1;
      return formatListNumber(
        value,
        legal && index !== ilvl ? "decimal" : (other?.format ?? "decimal"),
      );
    });
    return { text, level };
  };
}
