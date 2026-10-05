/**
 * Page and note numbers as text, for the `ST_NumberFormat` values (ECMA-376
 * Part 1 §17.18.59) Word offers for page numbers and notes, plus the field
 * format switches `\* roman`, `\* ALPHABETIC` … (§17.16.4.3). Formats not
 * listed here fall back to decimal, as Word does for formats it cannot show.
 */

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

// Word's footnote symbols, repeated (doubled, tripled …) after the fourth.
const CHICAGO = ["*", "†", "‡", "§"];
const ALPHABET_SIZE = 26;
const CHAR_CODE_A = 97;
// Roman numerals past this are written by Word as decimal.
const MAX_ROMAN = 3999;
const CIRCLED_ONE = 0x2460;
const MAX_CIRCLED = 20;

function roman(n: number): string {
  if (n <= 0 || n > MAX_ROMAN) return String(n);
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

/** a … z, then aa … zz, aaa … (Word repeats the letter, it does not carry). */
function letters(n: number): string {
  if (n <= 0) return String(n);
  const letter = String.fromCharCode(CHAR_CODE_A + ((n - 1) % ALPHABET_SIZE));
  return letter.repeat(Math.floor((n - 1) / ALPHABET_SIZE) + 1);
}

function ordinalSuffix(n: number): string {
  const tens = n % 100;
  if (tens >= 11 && tens <= 13) return "th";
  return ["th", "st", "nd", "rd"][n % 10] ?? "th";
}

/** `n` written in an `ST_NumberFormat` (or field `\*` switch). */
export function formatNumber(n: number, format: string | undefined): string {
  switch (format) {
    case "upperRoman":
    case "ROMAN":
      return roman(n).toUpperCase();
    case "lowerRoman":
    case "roman":
      return roman(n);
    case "upperLetter":
    case "ALPHABETIC":
      return letters(n).toUpperCase();
    case "lowerLetter":
    case "alphabetic":
      return letters(n);
    case "ordinal":
    case "Ordinal":
      return `${n}${ordinalSuffix(n)}`;
    case "decimalZero":
      return n < 10 && n >= 0 ? `0${n}` : String(n);
    case "numberInDash":
      return `- ${n} -`;
    case "chicago": {
      if (n <= 0) return String(n);
      const symbol = CHICAGO[(n - 1) % CHICAGO.length] ?? "*";
      return symbol.repeat(Math.floor((n - 1) / CHICAGO.length) + 1);
    }
    case "decimalEnclosedCircle":
      return n >= 1 && n <= MAX_CIRCLED ? String.fromCodePoint(CIRCLED_ONE + n - 1) : String(n);
    case "decimalEnclosedParen":
      return `(${n})`;
    case "hex":
    case "Hex":
      return n.toString(16).toUpperCase();
    default:
      return String(n);
  }
}
