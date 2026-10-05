/**
 * Field instruction parsing and result formatting (ECMA-376 Part 1 §17.16).
 *
 * A field instruction is `TYPE arg… \switch arg…`. Arguments are bare words or
 * double-quoted strings, in which `\"` and `\\` escape a quote and a backslash
 * (§17.16.1). The general formatting switches (`\*`, `\#`, `\@`) are shared
 * by every field (§17.16.4); this module implements them so computed results
 * read the way Word writes them.
 */

export interface FieldSwitch {
  /** The switch letter(s) without the backslash, e.g. `"*"`, `"@"`, `"h"`. */
  readonly name: string;
  readonly arg?: string;
}

export interface ParsedFieldInstruction {
  /** The field type in upper case (`"DATE"`, `"="`, `"REF"` …). */
  readonly type: string;
  /** Positional arguments in order, quotes removed. */
  readonly args: readonly string[];
  readonly switches: readonly FieldSwitch[];
}

// Switches that take an argument, per field (§17.16.5). The general switches
// `\*`, `\#` and `\@` always take one. Anything else is a flag.
const SWITCHES_WITH_ARG: Readonly<Record<string, ReadonlySet<string>>> = {
  HYPERLINK: new Set(["l", "o", "t"]),
  REF: new Set(["d"]),
  SEQ: new Set(["r", "s"]),
  SYMBOL: new Set(["f", "s"]),
  INCLUDETEXT: new Set(["c"]),
  INCLUDEPICTURE: new Set(["c"]),
  LISTNUM: new Set(["l", "s"]),
  AUTONUM: new Set(["s"]),
  AUTONUMLGL: new Set(["s"]),
  AUTONUMOUT: new Set(["s"]),
  ASK: new Set(["d"]),
  FILLIN: new Set(["d"]),
  BARCODE: new Set(["b", "t"]),
  ADVANCE: new Set(["d", "l", "r", "u", "x", "y"]),
  FORMDROPDOWN: new Set(),
};
const GENERAL_SWITCHES: ReadonlySet<string> = new Set(["*", "#", "@"]);

interface Token {
  readonly text: string;
  /** True for a backslash switch token. */
  readonly isSwitch: boolean;
}

function tokenize(instruction: string): Token[] {
  const tokens: Token[] = [];
  let i = 0;
  const n = instruction.length;
  while (i < n) {
    const ch = instruction[i] ?? "";
    if (/\s/.test(ch)) {
      i++;
      continue;
    }
    if (ch === '"') {
      let value = "";
      i++;
      while (i < n && instruction[i] !== '"') {
        const c = instruction[i] ?? "";
        if (c === "\\" && (instruction[i + 1] === '"' || instruction[i + 1] === "\\")) {
          value += instruction[i + 1];
          i += 2;
          continue;
        }
        value += c;
        i++;
      }
      i++;
      tokens.push({ text: value, isSwitch: false });
      continue;
    }
    if (ch === "\\") {
      // `\*`, `\#`, `\@`, `\!` are one-character switches; letters may run
      // together (`\h`, `\MERGEFORMAT` is not a switch but `\* MERGEFORMAT`).
      const next = instruction[i + 1] ?? "";
      if (/[*#@!]/.test(next)) {
        tokens.push({ text: next, isSwitch: true });
        i += 2;
        continue;
      }
      let j = i + 1;
      while (j < n && /[A-Za-z]/.test(instruction[j] ?? "")) j++;
      tokens.push({ text: instruction.slice(i + 1, j).toLowerCase(), isSwitch: true });
      i = j;
      continue;
    }
    let j = i;
    while (j < n && !/\s/.test(instruction[j] ?? "")) j++;
    tokens.push({ text: instruction.slice(i, j), isSwitch: false });
    i = j;
  }
  return tokens;
}

/** Parse a field instruction (the text of its `<w:instrText>` runs). */
export function parseFieldInstruction(instruction: string): ParsedFieldInstruction {
  const trimmed = instruction.trim();
  // `=` formulas keep the rest of the text as one expression argument.
  if (trimmed.startsWith("=")) {
    const body = trimmed.slice(1);
    const sw = body.search(/\\[*#@!]/);
    const expr = sw >= 0 ? body.slice(0, sw) : body;
    const rest = sw >= 0 ? parseSwitchesOnly(body.slice(sw), "=") : [];
    return { type: "=", args: expr.trim() ? [expr.trim()] : [], switches: rest };
  }
  const tokens = tokenize(trimmed);
  const first = tokens.shift();
  const type = (first?.text ?? "").toUpperCase();
  return { type, ...collect(tokens, type) };
}

function parseSwitchesOnly(text: string, type: string): FieldSwitch[] {
  return [...collect(tokenize(text), type).switches];
}

function collect(
  tokens: Token[],
  type: string,
): { args: string[]; switches: FieldSwitch[] } {
  const args: string[] = [];
  const switches: FieldSwitch[] = [];
  const withArg = SWITCHES_WITH_ARG[type];
  for (let i = 0; i < tokens.length; i++) {
    const tok = tokens[i];
    if (!tok) continue;
    if (!tok.isSwitch) {
      args.push(tok.text);
      continue;
    }
    const takesArg = GENERAL_SWITCHES.has(tok.text) || (withArg?.has(tok.text) ?? false);
    const next = tokens[i + 1];
    if (takesArg && next && !next.isSwitch) {
      switches.push({ name: tok.text, arg: next.text });
      i++;
    } else {
      switches.push({ name: tok.text });
    }
  }
  return { args, switches };
}

/** Quote a field argument when it needs it (spaces, quotes, backslashes). */
export function quoteFieldArgument(value: string): string {
  if (value !== "" && !/[\s"\\]/.test(value)) return value;
  return `"${value.replace(/\\/g, "\\\\").replace(/"/g, '\\"')}"`;
}

// --- \@ date-time pictures (§17.16.4.1) ------------------------------------

const MONTHS = [
  "January",
  "February",
  "March",
  "April",
  "May",
  "June",
  "July",
  "August",
  "September",
  "October",
  "November",
  "December",
];
const WEEKDAYS = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
const ABBREVIATION_LENGTH = 3;
const HOURS_PER_HALF_DAY = 12;

function pad(n: number, width: number): string {
  return String(n).padStart(width, "0");
}

/**
 * Format a date with a Word date-time picture such as `"dddd, MMMM d, yyyy"`
 * or `"h:mm am/pm"`. Text in single quotes is literal; month and day names
 * are English (Word uses the document language; the caller can pre-localize
 * by quoting literal names).
 */
export function formatDatePicture(date: Date, picture: string): string {
  let out = "";
  let i = 0;
  const p = picture;
  const run = (ch: string): number => {
    let j = i;
    while (p[j] === ch) j++;
    return j - i;
  };
  while (i < p.length) {
    const ch = p[i] ?? "";
    if (ch === "'") {
      const end = p.indexOf("'", i + 1);
      out += end < 0 ? p.slice(i + 1) : p.slice(i + 1, end);
      i = end < 0 ? p.length : end + 1;
      continue;
    }
    if (/^am\/pm/i.test(p.slice(i))) {
      const pm = date.getHours() >= HOURS_PER_HALF_DAY;
      const upper = p[i] === "A";
      out += upper ? (pm ? "PM" : "AM") : pm ? "pm" : "am";
      i += "am/pm".length;
      continue;
    }
    const len = run(ch);
    switch (ch) {
      case "d":
      case "D":
        if (len === 1) out += date.getDate();
        else if (len === 2) out += pad(date.getDate(), 2);
        else if (len === 3) out += (WEEKDAYS[date.getDay()] ?? "").slice(0, ABBREVIATION_LENGTH);
        else out += WEEKDAYS[date.getDay()] ?? "";
        break;
      case "M":
        if (len === 1) out += date.getMonth() + 1;
        else if (len === 2) out += pad(date.getMonth() + 1, 2);
        else if (len === 3) out += (MONTHS[date.getMonth()] ?? "").slice(0, ABBREVIATION_LENGTH);
        else out += MONTHS[date.getMonth()] ?? "";
        break;
      case "y":
      case "Y":
        out += len <= 2 ? pad(date.getFullYear() % 100, 2) : String(date.getFullYear());
        break;
      case "h": {
        // Lower-case h is the 12-hour clock, upper-case H the 24-hour one.
        const h = date.getHours() % HOURS_PER_HALF_DAY || HOURS_PER_HALF_DAY;
        out += len === 1 ? h : pad(h, 2);
        break;
      }
      case "H":
        out += len === 1 ? date.getHours() : pad(date.getHours(), 2);
        break;
      case "m":
        out += len === 1 ? date.getMinutes() : pad(date.getMinutes(), 2);
        break;
      case "s":
      case "S":
        out += len === 1 ? date.getSeconds() : pad(date.getSeconds(), 2);
        break;
      default:
        out += p.slice(i, i + len);
    }
    i += len;
  }
  return out;
}

// --- \* format switches (§17.16.4.3) ----------------------------------------

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

function toRoman(n: number): string {
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

const ALPHABET_SIZE = 26;
const CHAR_CODE_A = 65;

/** Word's alphabetic numbering: A…Z, AA…ZZ, AAA… (the letter repeats). */
function toAlphabetic(n: number): string {
  const letter = String.fromCharCode(CHAR_CODE_A + ((n - 1) % ALPHABET_SIZE));
  return letter.repeat(Math.floor((n - 1) / ALPHABET_SIZE) + 1);
}

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
const TEEN_LIMIT = 20;
const HUNDRED = 100;
const THOUSAND = 1000;
const MILLION = 1_000_000;

function cardinal(n: number): string {
  if (n < TEEN_LIMIT) return ONES[n] ?? "";
  if (n < HUNDRED) {
    const t = TENS[Math.floor(n / 10)] ?? "";
    return n % 10 ? `${t}-${ONES[n % 10]}` : t;
  }
  if (n < THOUSAND) {
    const rest = n % HUNDRED;
    return `${ONES[Math.floor(n / HUNDRED)]} hundred${rest ? ` ${cardinal(rest)}` : ""}`;
  }
  if (n < MILLION) {
    const rest = n % THOUSAND;
    return `${cardinal(Math.floor(n / THOUSAND))} thousand${rest ? ` ${cardinal(rest)}` : ""}`;
  }
  const rest = n % MILLION;
  return `${cardinal(Math.floor(n / MILLION))} million${rest ? ` ${cardinal(rest)}` : ""}`;
}

const ORDINAL_WORDS: Readonly<Record<string, string>> = {
  one: "first",
  two: "second",
  three: "third",
  five: "fifth",
  eight: "eighth",
  nine: "ninth",
  twelve: "twelfth",
};

function ordinalText(n: number): string {
  const words = cardinal(n);
  const m = /([a-z]+)$/.exec(words);
  const last = m?.[1] ?? "";
  const head = words.slice(0, words.length - last.length);
  const irregular = ORDINAL_WORDS[last];
  if (irregular) return head + irregular;
  if (last.endsWith("y")) return `${head}${last.slice(0, -1)}ieth`;
  return `${head}${last}th`;
}

function ordinalSuffix(n: number): string {
  const mod100 = n % HUNDRED;
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

const CENTS = 100;

/** Apply one `\*` format switch argument to a result string. */
export function applyFormatSwitch(value: string, format: string): string {
  const key = format.toUpperCase();
  const n = Number(value.replace(/,/g, ""));
  const isInt = value.trim() !== "" && Number.isInteger(n) && n >= 0;
  switch (key) {
    case "UPPER":
      return value.toUpperCase();
    case "LOWER":
      return value.toLowerCase();
    case "FIRSTCAP":
      return value.charAt(0).toUpperCase() + value.slice(1);
    case "CAPS":
      return value.replace(/\b\p{L}/gu, (c) => c.toUpperCase());
    case "MERGEFORMAT":
    case "CHARFORMAT":
      return value;
    default:
      break;
  }
  if (!isInt) return value;
  switch (key) {
    case "ARABIC":
      return String(n);
    case "ROMAN":
      return format === "roman" ? toRoman(n).toLowerCase() : toRoman(n);
    case "ALPHABETIC":
      return format === "alphabetic" ? toAlphabetic(n).toLowerCase() : toAlphabetic(n);
    case "ORDINAL":
      return ordinalSuffix(n);
    case "CARDTEXT":
      return cardinal(n);
    case "ORDTEXT":
      return ordinalText(n);
    case "HEX":
      return n.toString(16).toUpperCase();
    case "DOLLARTEXT": {
      const whole = Math.floor(n);
      const cents = Math.round((n - whole) * CENTS);
      return `${cardinal(whole)} and ${pad(cents, 2)}/100`;
    }
    default:
      return value;
  }
}

/**
 * Format a number with a `\#` numeric picture (§17.16.4.2): `0` a required
 * digit, `#` an optional digit, `,` grouping, `.` the decimal point; any other
 * character is literal. A `;` splits positive and negative pictures.
 */
export function formatNumericPicture(value: number, picture: string): string {
  const [positive = "", negative] = picture.replace(/'/g, "").split(";");
  const pic = value < 0 && negative !== undefined ? negative : positive;
  const abs = value < 0 && negative !== undefined ? -value : value;
  const dot = pic.indexOf(".");
  const intPic = dot >= 0 ? pic.slice(0, dot) : pic;
  const fracPic = dot >= 0 ? pic.slice(dot + 1) : "";
  const fracDigits = (fracPic.match(/[0#]/g) ?? []).length;
  const fixed = Math.abs(abs).toFixed(fracDigits);
  const [intDigits = "0", fracRaw = ""] = fixed.split(".");
  const grouping = intPic.includes(",");
  const minInt = (intPic.match(/0/g) ?? []).length;
  let intStr = intDigits === "0" && minInt === 0 ? "" : intDigits.padStart(minInt, "0");
  if (grouping) intStr = intStr.replace(/\B(?=(\d{3})+(?!\d))/g, ",");
  const prefix = intPic.slice(0, intPic.search(/[0#,]|$/));
  const suffix = intPic.slice(intPic.search(/[^0#,]*$/));
  let frac = fracRaw;
  const requiredFrac = (fracPic.match(/0/g) ?? []).length;
  while (frac.length > requiredFrac && frac.endsWith("0")) frac = frac.slice(0, -1);
  const fracSuffix = fracPic.replace(/[0#]/g, "");
  const sign = abs < 0 ? "-" : "";
  return `${prefix.replace(/[-]/, "")}${sign}${intStr}${frac ? `.${frac}` : ""}${suffix}${fracSuffix}`;
}

/** Apply every general formatting switch of a parsed instruction to a result. */
export function applyGeneralSwitches(
  value: string,
  switches: readonly FieldSwitch[],
): string {
  let out = value;
  for (const sw of switches) {
    if (sw.arg === undefined) continue;
    if (sw.name === "#") {
      const n = Number(out.replace(/,/g, ""));
      if (out.trim() !== "" && Number.isFinite(n)) out = formatNumericPicture(n, sw.arg);
    } else if (sw.name === "*") {
      out = applyFormatSwitch(out, sw.arg);
    }
  }
  return out;
}

// --- = formula fields (§17.16.3) --------------------------------------------

type FormulaToken =
  | { kind: "num"; value: number }
  | { kind: "op"; value: string }
  | { kind: "name"; value: string };

function tokenizeFormula(expr: string): FormulaToken[] {
  const out: FormulaToken[] = [];
  let i = 0;
  while (i < expr.length) {
    const ch = expr[i] ?? "";
    if (/\s/.test(ch)) {
      i++;
      continue;
    }
    // A comma separates function arguments, so it never groups digits here.
    const num = /^(\d+(\.\d+)?|\.\d+)/.exec(expr.slice(i));
    if (num) {
      out.push({ kind: "num", value: Number(num[0].replace(/,/g, "")) });
      i += num[0].length;
      continue;
    }
    const two = expr.slice(i, i + 2);
    if (two === "<=" || two === ">=" || two === "<>") {
      out.push({ kind: "op", value: two });
      i += 2;
      continue;
    }
    if ("+-*/^%()=<>,;".includes(ch)) {
      out.push({ kind: "op", value: ch });
      i++;
      continue;
    }
    const name = /^[A-Za-z_][\w]*/.exec(expr.slice(i));
    if (name) {
      out.push({ kind: "name", value: name[0].toUpperCase() });
      i += name[0].length;
      continue;
    }
    throw new Error(`Unexpected character ${JSON.stringify(ch)} in formula.`);
  }
  return out;
}

const FORMULA_FUNCTIONS: Readonly<Record<string, (args: number[]) => number>> = {
  ABS: (a) => Math.abs(a[0] ?? 0),
  AND: (a) => (a.every((x) => x !== 0) ? 1 : 0),
  AVERAGE: (a) => (a.length ? a.reduce((s, x) => s + x, 0) / a.length : 0),
  COUNT: (a) => a.length,
  DEFINED: (a) => (a.every((x) => Number.isFinite(x)) ? 1 : 0),
  FALSE: () => 0,
  INT: (a) => Math.trunc(a[0] ?? 0),
  MAX: (a) => Math.max(...a),
  MIN: (a) => Math.min(...a),
  MOD: (a) => (a[0] ?? 0) % (a[1] ?? 1),
  NOT: (a) => ((a[0] ?? 0) === 0 ? 1 : 0),
  OR: (a) => (a.some((x) => x !== 0) ? 1 : 0),
  PRODUCT: (a) => a.reduce((s, x) => s * x, 1),
  ROUND: (a) => {
    const f = 10 ** (a[1] ?? 0);
    return Math.round((a[0] ?? 0) * f) / f;
  },
  SIGN: (a) => Math.sign(a[0] ?? 0),
  SUM: (a) => a.reduce((s, x) => s + x, 0),
  TRUE: () => 1,
};

/**
 * Evaluate a `=` formula's arithmetic (numbers, `+ - * / ^ %`, comparisons
 * and Word's functions). Table-cell references and bookmarks are resolved by
 * `lookup`, which returns undefined for unknown names.
 */
export function evaluateFormula(
  expr: string,
  lookup: (name: string) => number | undefined = () => undefined,
): number {
  const tokens = tokenizeFormula(expr);
  let pos = 0;
  const peek = (): FormulaToken | undefined => tokens[pos];
  const isOp = (v: string): boolean => {
    const t = peek();
    return t?.kind === "op" && t.value === v;
  };
  /** Consume the current operator token and return its text. */
  const takeOp = (): string => {
    const t = tokens[pos++];
    return t?.kind === "op" ? t.value : "";
  };
  const expect = (v: string): void => {
    if (!isOp(v)) throw new Error(`Expected ${JSON.stringify(v)} in formula.`);
    pos++;
  };
  const comparison = (): number => {
    let left = additive();
    for (;;) {
      const t = peek();
      if (t?.kind !== "op" || !["=", "<", ">", "<=", ">=", "<>"].includes(t.value)) return left;
      pos++;
      const right = additive();
      left = compare(left, t.value, right) ? 1 : 0;
    }
  };
  const additive = (): number => {
    let left = term();
    while (isOp("+") || isOp("-")) {
      const op = takeOp();
      const right = term();
      left = op === "+" ? left + right : left - right;
    }
    return left;
  };
  const term = (): number => {
    let left = power();
    while (isOp("*") || isOp("/")) {
      const op = takeOp();
      const right = power();
      if (op === "/" && right === 0) throw new Error("!Zero Divide");
      left = op === "*" ? left * right : left / right;
    }
    return left;
  };
  const power = (): number => {
    const base = unary();
    if (isOp("^")) {
      pos++;
      return base ** power();
    }
    return base;
  };
  const unary = (): number => {
    if (isOp("-")) {
      pos++;
      return -unary();
    }
    if (isOp("+")) {
      pos++;
      return unary();
    }
    let value = primary();
    if (isOp("%")) {
      pos++;
      value /= HUNDRED;
    }
    return value;
  };
  const primary = (): number => {
    const t = tokens[pos++];
    if (!t) throw new Error("Unexpected end of formula.");
    if (t.kind === "num") return t.value;
    if (t.kind === "op" && t.value === "(") {
      const v = comparison();
      expect(")");
      return v;
    }
    if (t.kind === "name") {
      const fn = FORMULA_FUNCTIONS[t.value];
      if (fn && isOp("(")) {
        pos++;
        const args: number[] = [];
        if (!isOp(")")) {
          args.push(comparison());
          while (isOp(",") || isOp(";")) {
            pos++;
            args.push(comparison());
          }
        }
        expect(")");
        return fn(args);
      }
      if (fn) return fn([]);
      const v = lookup(t.value);
      if (v === undefined) throw new Error(`!Undefined Bookmark, ${t.value}`);
      return v;
    }
    throw new Error("!Syntax Error");
  };
  const result = comparison();
  if (pos < tokens.length) throw new Error("!Syntax Error");
  return result;
}

function compare(a: number | string, op: string, b: number | string): boolean {
  switch (op) {
    case "=":
      return a === b;
    case "<>":
      return a !== b;
    case "<":
      return a < b;
    case ">":
      return a > b;
    case "<=":
      return a <= b;
    case ">=":
      return a >= b;
    default:
      return false;
  }
}

/**
 * Evaluate an IF / COMPARE condition `left op right` (§17.16.5.31): numbers
 * compare numerically, everything else as text; `=`/`<>` on text accept the
 * `?` and `*` wildcards in the right operand.
 */
export function evaluateComparison(left: string, op: string, right: string): boolean {
  const a = Number(left);
  const b = Number(right);
  if (left.trim() !== "" && right.trim() !== "" && Number.isFinite(a) && Number.isFinite(b)) {
    return compare(a, op, b);
  }
  if ((op === "=" || op === "<>") && /[?*]/.test(right)) {
    const re = new RegExp(
      `^${right.replace(/[.+^${}()|[\]\\]/g, "\\$&").replace(/\?/g, ".").replace(/\*/g, ".*")}$`,
    );
    return op === "=" ? re.test(left) : !re.test(left);
  }
  return compare(left, op, right);
}
