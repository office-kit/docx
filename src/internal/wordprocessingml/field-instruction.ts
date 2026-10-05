/**
 * Field instructions (ECMA-376 Part 1 §17.16): tokenizing an instruction into
 * its field type, arguments and switches, and the number formats the general
 * `\*` switch and the note/caption numbering share.
 */

/** One `\x` switch and its argument, if it takes one. */
export interface FieldSwitch {
  /** The switch without its backslash, case preserved (`o`, `h`, `*`, `#`). */
  readonly name: string;
  readonly arg?: string;
}

/** A parsed field instruction. */
export interface FieldInstruction {
  /** The field type, upper-cased (`TOC`, `SEQ`, `MERGEFIELD`). */
  readonly type: string;
  /** Positional arguments (unquoted). */
  readonly args: readonly string[];
  readonly switches: readonly FieldSwitch[];
}

// Switches that never take an argument, per field type (§17.16.5). Any other
// switch followed by a non-switch token takes it as its argument, which is
// how Word reads instructions too.
const FLAG_SWITCHES: Readonly<Record<string, ReadonlySet<string>>> = {
  TOC: new Set(["h", "z", "u", "w", "x"]),
  SEQ: new Set(["c", "h", "n"]),
  XE: new Set(["b", "i"]),
  TA: new Set(["b", "i"]),
  TOA: new Set(["h", "f", "p"]),
  INDEX: new Set(["r"]),
  PAGEREF: new Set(["h", "p"]),
  REF: new Set(["f", "h", "n", "p", "r", "t", "w"]),
  NOTEREF: new Set(["f", "h", "p"]),
  CITATION: new Set(["n", "t", "y"]),
  MERGEFIELD: new Set(["m", "v"]),
  ADDRESSBLOCK: new Set(["d"]),
  ASK: new Set(["o"]),
  FILLIN: new Set(["o"]),
  STYLEREF: new Set(["l", "n", "p", "r", "t", "w"]),
};
// The general formatting switches (§17.16.4) always take an argument.
const GENERAL_SWITCHES: ReadonlySet<string> = new Set(["*", "#", "@"]);

interface Token {
  readonly value: string;
  readonly quoted: boolean;
}

function tokenize(instruction: string): Token[] {
  const tokens: Token[] = [];
  let i = 0;
  const n = instruction.length;
  while (i < n) {
    const ch = instruction[i] as string;
    if (/\s/.test(ch)) {
      i++;
      continue;
    }
    if (ch === '"') {
      // Inside quotes a backslash escapes `"` and `\` (§17.16.1).
      let value = "";
      i++;
      while (i < n && instruction[i] !== '"') {
        const c = instruction[i] as string;
        if (c === "\\" && (instruction[i + 1] === '"' || instruction[i + 1] === "\\")) {
          value += instruction[i + 1];
          i += 2;
        } else {
          value += c;
          i++;
        }
      }
      i++;
      tokens.push({ value, quoted: true });
      continue;
    }
    if (ch === "\\") {
      // `\*`, `\#`, `\@` or a letter switch.
      const name = instruction[i + 1] ?? "";
      tokens.push({ value: `\\${name}`, quoted: false });
      i += 2;
      continue;
    }
    let value = "";
    while (i < n && !/\s/.test(instruction[i] as string) && instruction[i] !== '"') {
      value += instruction[i];
      i++;
    }
    tokens.push({ value, quoted: false });
  }
  return tokens;
}

function isSwitch(token: Token | undefined): boolean {
  return !!token && !token.quoted && token.value.startsWith("\\") && token.value.length === 2;
}

/** Parse a field instruction (`TOC \o "1-3" \h`) into type, arguments and switches. */
export function parseFieldInstruction(instruction: string): FieldInstruction {
  const tokens = tokenize(instruction);
  const type = (tokens[0]?.value ?? "").toUpperCase();
  const flags = FLAG_SWITCHES[type] ?? new Set<string>();
  const args: string[] = [];
  const switches: FieldSwitch[] = [];
  for (let i = 1; i < tokens.length; i++) {
    const token = tokens[i] as Token;
    if (!isSwitch(token)) {
      args.push(token.value);
      continue;
    }
    const name = token.value.slice(1);
    const next = tokens[i + 1];
    const takesArg =
      GENERAL_SWITCHES.has(name) || (!flags.has(name.toLowerCase()) && next && !isSwitch(next));
    if (takesArg && next && !isSwitch(next)) {
      switches.push({ name, arg: next.value });
      i++;
    } else {
      switches.push({ name });
    }
  }
  return { type, args, switches };
}

/** The first switch named `name` (case-sensitive for `\*`, letters case-insensitive). */
export function fieldSwitch(field: FieldInstruction, name: string): FieldSwitch | undefined {
  const lower = name.toLowerCase();
  return field.switches.find((s) => s.name.toLowerCase() === lower);
}

/** Quote a field argument when it needs it (§17.16.1). */
export function quoteFieldArg(value: string): string {
  if (value !== "" && !/[\s"\\]/.test(value)) return value;
  return quotedFieldArg(value);
}

/** Always quote a field argument, as Word does for entry text (XE, TA, `\c "Figure"`). */
export function quotedFieldArg(value: string): string {
  return `"${value.replace(/\\/g, "\\\\").replace(/"/g, '\\"')}"`;
}

/**
 * Number formats used for note marks, caption and SEQ numbering: the
 * ST_NumberFormat values (§17.18.59) Word's dialogs offer.
 */
export type NumberingFormat =
  | "decimal"
  | "upperRoman"
  | "lowerRoman"
  | "upperLetter"
  | "lowerLetter"
  | "ordinal"
  | "cardinalText"
  | "ordinalText"
  | "hex"
  | "chicago"
  | "decimalZero"
  | "numberInDash";

const ROMAN: ReadonlyArray<readonly [number, string]> = [
  [1000, "M"],
  [900, "CM"],
  [500, "D"],
  [400, "CD"],
  [100, "C"],
  [90, "XC"],
  [50, "L"],
  [40, "XL"],
  [10, "X"],
  [9, "IX"],
  [5, "V"],
  [4, "IV"],
  [1, "I"],
];
// The Chicago Manual of Style note marks, repeated (doubled, tripled …) past four.
const CHICAGO_MARKS = ["*", "†", "‡", "§"];
const ALPHABET_SIZE = 26;
const ONES = [
  "zero",
  "one",
  "two",
  "three",
  "four",
  "five",
  "six",
  "seven",
  "eight",
  "nine",
  "ten",
  "eleven",
  "twelve",
  "thirteen",
  "fourteen",
  "fifteen",
  "sixteen",
  "seventeen",
  "eighteen",
  "nineteen",
];
const TENS = ["", "", "twenty", "thirty", "forty", "fifty", "sixty", "seventy", "eighty", "ninety"];
const ORDINAL_WORDS: Readonly<Record<string, string>> = {
  one: "first",
  two: "second",
  three: "third",
  five: "fifth",
  eight: "eighth",
  nine: "ninth",
  twelve: "twelfth",
};

function roman(n: number): string {
  let rest = n;
  let out = "";
  for (const [value, glyph] of ROMAN) {
    while (rest >= value) {
      out += glyph;
      rest -= value;
    }
  }
  return out;
}

/** Word's letter numbering: a…z, then aa, bb … (the letter repeated). */
function letters(n: number): string {
  const letter = String.fromCharCode("A".charCodeAt(0) + ((n - 1) % ALPHABET_SIZE));
  return letter.repeat(Math.floor((n - 1) / ALPHABET_SIZE) + 1);
}

function cardinalWords(n: number): string {
  if (n < 20) return ONES[n] as string;
  if (n < 100) {
    const tens = TENS[Math.floor(n / 10)] as string;
    return n % 10 === 0 ? tens : `${tens}-${ONES[n % 10]}`;
  }
  if (n < 1000) {
    const rest = n % 100;
    const head = `${ONES[Math.floor(n / 100)]} hundred`;
    return rest === 0 ? head : `${head} ${cardinalWords(rest)}`;
  }
  const rest = n % 1000;
  const head = `${cardinalWords(Math.floor(n / 1000))} thousand`;
  return rest === 0 ? head : `${head} ${cardinalWords(rest)}`;
}

function ordinalWords(n: number): string {
  const words = cardinalWords(n);
  const m = /([a-z]+)$/.exec(words);
  const last = m?.[1] ?? "";
  const head = words.slice(0, words.length - last.length);
  const ordinal =
    ORDINAL_WORDS[last] ?? (last.endsWith("y") ? `${last.slice(0, -1)}ieth` : `${last}th`);
  return head + ordinal;
}

function ordinalSuffix(n: number): string {
  const mod100 = n % 100;
  if (mod100 >= 11 && mod100 <= 13) return `${n}th`;
  switch (n % 10) {
    case 1:
      return `${n}st`;
    case 2:
      return `${n}nd`;
    case 3:
      return `${n}rd`;
    default:
      return `${n}th`;
  }
}

/** Format a positive number in one of the {@link NumberingFormat}s. */
export function formatNumber(n: number, format: NumberingFormat): string {
  if (n < 1 && format !== "decimal" && format !== "decimalZero") return String(n);
  switch (format) {
    case "decimal":
      return String(n);
    case "decimalZero":
      return n < 10 ? `0${n}` : String(n);
    case "numberInDash":
      return `- ${n} -`;
    case "upperRoman":
      return roman(n);
    case "lowerRoman":
      return roman(n).toLowerCase();
    case "upperLetter":
      return letters(n);
    case "lowerLetter":
      return letters(n).toLowerCase();
    case "ordinal":
      return ordinalSuffix(n);
    case "cardinalText":
      return cardinalWords(n);
    case "ordinalText":
      return ordinalWords(n);
    case "hex":
      return n.toString(16).toUpperCase();
    case "chicago":
      return (CHICAGO_MARKS[(n - 1) % CHICAGO_MARKS.length] as string).repeat(
        Math.floor((n - 1) / CHICAGO_MARKS.length) + 1,
      );
  }
}

// `\*` numeric picture names (§17.16.4.3) → the matching number format. The
// switch is case-sensitive for the letter and roman forms.
const GENERAL_FORMATS: Readonly<Record<string, NumberingFormat>> = {
  Arabic: "decimal",
  ARABIC: "decimal",
  arabic: "decimal",
  ArabicDash: "numberInDash",
  ROMAN: "upperRoman",
  Roman: "upperRoman",
  roman: "lowerRoman",
  ALPHABETIC: "upperLetter",
  Alphabetic: "upperLetter",
  alphabetic: "lowerLetter",
  Ordinal: "ordinal",
  ordinal: "ordinal",
  CardText: "cardinalText",
  cardtext: "cardinalText",
  OrdText: "ordinalText",
  ordtext: "ordinalText",
  Hex: "hex",
  hex: "hex",
};

/** The `\*` switch name Word writes for a number format. */
export const GENERAL_FORMAT_SWITCH: Readonly<Partial<Record<NumberingFormat, string>>> = {
  decimal: "ARABIC",
  upperRoman: "ROMAN",
  lowerRoman: "roman",
  upperLetter: "ALPHABETIC",
  lowerLetter: "alphabetic",
  ordinal: "Ordinal",
  cardinalText: "CardText",
  ordinalText: "OrdText",
  hex: "Hex",
  numberInDash: "ArabicDash",
};

/** Apply a field's `\*` number format switch (ARABIC when absent or not a number format). */
export function formatFieldNumber(n: number, field: FieldInstruction): string {
  for (const s of field.switches) {
    if (s.name !== "*" || s.arg === undefined) continue;
    const format = GENERAL_FORMATS[s.arg];
    if (format) return formatNumber(n, format);
  }
  return String(n);
}
