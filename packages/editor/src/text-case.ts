/**
 * Word's Change Case transforms (Home ▸ Change Case), including the Japanese
 * UI's half-width / full-width and hiragana / katakana entries.
 */

export type CaseMode =
  | "sentence"
  | "lower"
  | "upper"
  | "title"
  | "toggle"
  | "halfWidth"
  | "fullWidth"
  | "katakana"
  | "hiragana";

/** Modes that map each character to exactly one character. */
export const LENGTH_PRESERVING_MODES: ReadonlySet<CaseMode> = new Set([
  "sentence",
  "lower",
  "upper",
  "title",
  "toggle",
  "katakana",
  "hiragana",
]);

// Upper / lower case one character, keeping it when the mapping would change
// the length ("ß" → "SS"), so run boundaries stay where they were.
function upper(c: string): string {
  const u = c.toUpperCase();
  return u.length === c.length ? u : c;
}

function lower(c: string): string {
  const l = c.toLowerCase();
  return l.length === c.length ? l : c;
}

const LETTER = /\p{L}/u;
const SENTENCE_END = /[.!?。！？]/;
const WORD_BREAK = /[\s\p{P}]/u;

/**
 * Sentence / title / upper / lower / toggle case. `atSentenceStart` says
 * whether the text begins a sentence (the start of a paragraph does).
 */
function caseText(text: string, mode: CaseMode, atSentenceStart: boolean): string {
  const chars = [...text];
  switch (mode) {
    case "lower":
      return chars.map(lower).join("");
    case "upper":
      return chars.map(upper).join("");
    case "toggle":
      return chars.map((c) => (upper(c) === c ? lower(c) : upper(c))).join("");
    case "title": {
      let wordStart = true;
      return chars
        .map((c) => {
          const out = LETTER.test(c) ? (wordStart ? upper(c) : lower(c)) : c;
          wordStart = WORD_BREAK.test(c);
          return out;
        })
        .join("");
    }
    case "sentence": {
      let sentenceStart = atSentenceStart;
      return chars
        .map((c) => {
          if (LETTER.test(c)) {
            const out = sentenceStart ? upper(c) : lower(c);
            sentenceStart = false;
            return out;
          }
          if (SENTENCE_END.test(c)) sentenceStart = true;
          return c;
        })
        .join("");
    }
    default:
      return text;
  }
}

const HIRAGANA_START = 0x3041;
const HIRAGANA_END = 0x3096;
const KANA_OFFSET = 0x60;
const FULL_ASCII_START = 0xff01;
const FULL_ASCII_END = 0xff5e;
const HALF_ASCII_START = 0x21;
const HALF_ASCII_END = 0x7e;
const FULL_WIDTH_OFFSET = 0xfee0;
const IDEOGRAPHIC_SPACE = "　";

// Half-width katakana (U+FF61–FF9F) and the full-width forms they stand for,
// index for index.
const HALF_KANA = "｡｢｣､･ｦｧｨｩｪｫｬｭｮｯｰｱｲｳｴｵｶｷｸｹｺｻｼｽｾｿﾀﾁﾂﾃﾄﾅﾆﾇﾈﾉﾊﾋﾌﾍﾎﾏﾐﾑﾒﾓﾔﾕﾖﾗﾘﾙﾚﾛﾜﾝﾞﾟ";
const FULL_KANA =
  "。「」、・ヲァィゥェォャュョッーアイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワン゛゜";
const VOICED = "ﾞ";
const SEMI_VOICED = "ﾟ";
// Full-width kana with a (semi-)voiced mark, as base + mark.
const VOICED_BASES = "カキクケコサシスセソタチツテトハヒフヘホウ";
const SEMI_VOICED_BASES = "ハヒフヘホ";

const HALF_TO_FULL = new Map([...HALF_KANA].map((c, i) => [c, [...FULL_KANA][i] ?? c]));
const FULL_TO_HALF = new Map([...FULL_KANA].map((c, i) => [c, [...HALF_KANA][i] ?? c]));

function compose(base: string, mark: string): string | undefined {
  const code = base.codePointAt(0) ?? 0;
  if (mark === VOICED && VOICED_BASES.includes(base)) {
    // ウ + ゛ is ヴ; the others are the next code point.
    return base === "ウ" ? "ヴ" : String.fromCodePoint(code + 1);
  }
  if (mark === SEMI_VOICED && SEMI_VOICED_BASES.includes(base))
    return String.fromCodePoint(code + 2);
  return undefined;
}

function decompose(c: string): string | undefined {
  if (c === "ヴ") return `${FULL_TO_HALF.get("ウ")}${VOICED}`;
  const code = c.codePointAt(0) ?? 0;
  const prev = String.fromCodePoint(code - 1);
  if (VOICED_BASES.includes(prev)) return `${FULL_TO_HALF.get(prev)}${VOICED}`;
  const prev2 = String.fromCodePoint(code - 2);
  if (SEMI_VOICED_BASES.includes(prev2)) return `${FULL_TO_HALF.get(prev2)}${SEMI_VOICED}`;
  return undefined;
}

function toFullWidth(text: string): string {
  const chars = [...text];
  let out = "";
  for (let i = 0; i < chars.length; i++) {
    const c = chars[i] ?? "";
    const code = c.codePointAt(0) ?? 0;
    if (code >= HALF_ASCII_START && code <= HALF_ASCII_END) {
      out += String.fromCodePoint(code + FULL_WIDTH_OFFSET);
    } else if (c === " ") {
      out += IDEOGRAPHIC_SPACE;
    } else if (HALF_TO_FULL.has(c)) {
      const full = HALF_TO_FULL.get(c) ?? c;
      const next = chars[i + 1];
      const composed = next === undefined ? undefined : compose(full, next);
      if (composed) {
        out += composed;
        i++;
      } else {
        out += full;
      }
    } else {
      out += c;
    }
  }
  return out;
}

function toHalfWidth(text: string): string {
  return [...text]
    .map((c) => {
      const code = c.codePointAt(0) ?? 0;
      if (code >= FULL_ASCII_START && code <= FULL_ASCII_END)
        return String.fromCodePoint(code - FULL_WIDTH_OFFSET);
      if (c === IDEOGRAPHIC_SPACE) return " ";
      return FULL_TO_HALF.get(c) ?? decompose(c) ?? c;
    })
    .join("");
}

function shiftKana(text: string, toKatakana: boolean): string {
  return [...text]
    .map((c) => {
      const code = c.codePointAt(0) ?? 0;
      if (toKatakana && code >= HIRAGANA_START && code <= HIRAGANA_END)
        return String.fromCodePoint(code + KANA_OFFSET);
      if (!toKatakana && code >= HIRAGANA_START + KANA_OFFSET && code <= HIRAGANA_END + KANA_OFFSET)
        return String.fromCodePoint(code - KANA_OFFSET);
      return c;
    })
    .join("");
}

/** Apply a Change Case mode to `text`. */
export function changeCase(text: string, mode: CaseMode, atSentenceStart = true): string {
  switch (mode) {
    case "fullWidth":
      return toFullWidth(text);
    case "halfWidth":
      return toHalfWidth(text);
    case "katakana":
      return shiftKana(text, true);
    case "hiragana":
      return shiftKana(text, false);
    default:
      return caseText(text, mode, atSentenceStart);
  }
}

/** Whether `text` ends a sentence, so the next text starts one. */
export function endsSentence(text: string): boolean {
  const trimmed = text.trimEnd();
  return trimmed === "" ? false : SENTENCE_END.test(trimmed.at(-1) ?? "");
}
