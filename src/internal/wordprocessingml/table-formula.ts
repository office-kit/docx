/**
 * The `=` (Formula) field's expression language (ECMA-376 §17.16.3) as Word
 * evaluates it inside a table: arithmetic and comparison operators, the
 * functions Word's Formula dialog offers, A1-style cell references and ranges,
 * and the positional ranges ABOVE / BELOW / LEFT / RIGHT. Bookmark references
 * are not supported (they need the bookmark's text, which a table alone does
 * not hold).
 */

/** Reads the table for the evaluator: a cell's number, by row and grid column. */
export interface FormulaTable {
  readonly rows: number;
  readonly columns: number;
  /** The number a cell holds, or undefined when its text is not a number. */
  value(row: number, column: number): number | undefined;
  /** The formula cell. */
  readonly at: { readonly row: number; readonly column: number };
}

type Token =
  | { kind: "num"; value: number }
  | { kind: "name"; value: string }
  | { kind: "op"; value: string };

// A number, an identifier / cell reference (letters, digits, ':' for ranges),
// or an operator. Comparison operators are matched before their prefixes.
const TOKEN =
  /\s*(?:(\d+(?:\.\d+)?|\.\d+)|([A-Za-z][A-Za-z0-9]*(?::[A-Za-z]+\d+)?)|(<=|>=|<>|[-+*/^%=<>(),]))/y;

function tokenize(source: string): Token[] {
  const out: Token[] = [];
  let pos = 0;
  while (pos < source.length) {
    if (/^\s*$/.test(source.slice(pos))) break;
    TOKEN.lastIndex = pos;
    const m = TOKEN.exec(source);
    if (!m) throw new SyntaxError(`Syntax error at "${source.slice(pos)}"`);
    pos = TOKEN.lastIndex;
    if (m[1] !== undefined) out.push({ kind: "num", value: Number(m[1]) });
    else if (m[2] !== undefined) out.push({ kind: "name", value: m[2].toUpperCase() });
    else if (m[3] !== undefined) out.push({ kind: "op", value: m[3] });
  }
  return out;
}

/** A function argument: one number, or the numbers of a range. */
type Arg = number | readonly number[];

const CELL_REF = /^([A-Z]+)(\d+)$/;
const POSITIONAL = new Set(["ABOVE", "BELOW", "LEFT", "RIGHT"]);
const LETTER_BASE = 26;
const A_CODE = 65;

function columnIndex(letters: string): number {
  let n = 0;
  for (const ch of letters) n = n * LETTER_BASE + (ch.charCodeAt(0) - A_CODE + 1);
  return n - 1;
}

function sum(values: readonly number[]): number {
  return values.reduce((a, b) => a + b, 0);
}

function flat(args: readonly Arg[]): number[] {
  return args.flatMap((a) => (typeof a === "number" ? [a] : [...a]));
}

function scalar(arg: Arg | undefined): number {
  if (typeof arg === "number") return arg;
  if (arg === undefined) throw new SyntaxError("A function argument is missing.");
  throw new SyntaxError("This function takes a number, not a range.");
}

const FUNCTIONS: Readonly<Record<string, (args: readonly Arg[]) => number>> = {
  SUM: (a) => sum(flat(a)),
  AVERAGE: (a) => {
    const v = flat(a);
    return v.length ? sum(v) / v.length : 0;
  },
  COUNT: (a) => flat(a).length,
  MAX: (a) => Math.max(...flat(a)),
  MIN: (a) => Math.min(...flat(a)),
  PRODUCT: (a) => flat(a).reduce((x, y) => x * y, 1),
  ABS: (a) => Math.abs(scalar(a[0])),
  INT: (a) => Math.trunc(scalar(a[0])),
  MOD: (a) => scalar(a[0]) % scalar(a[1]),
  ROUND: (a) => {
    const factor = 10 ** scalar(a[1]);
    return Math.round(scalar(a[0]) * factor) / factor;
  },
  SIGN: (a) => Math.sign(scalar(a[0])),
  AND: (a) => (scalar(a[0]) !== 0 && scalar(a[1]) !== 0 ? 1 : 0),
  OR: (a) => (scalar(a[0]) !== 0 || scalar(a[1]) !== 0 ? 1 : 0),
  NOT: (a) => (scalar(a[0]) === 0 ? 1 : 0),
  IF: (a) => (scalar(a[0]) !== 0 ? scalar(a[1]) : scalar(a[2])),
  DEFINED: (a) => (a.length > 0 ? 1 : 0),
};

const CONSTANTS: Readonly<Record<string, number>> = { TRUE: 1, FALSE: 0 };

class Parser {
  private pos = 0;
  constructor(
    private readonly tokens: Token[],
    private readonly table: FormulaTable,
  ) {}

  parse(): number {
    const value = this.comparison();
    if (this.pos < this.tokens.length) throw new SyntaxError("Unexpected end of formula.");
    return value;
  }

  private peekOp(...ops: string[]): string | undefined {
    const t = this.tokens[this.pos];
    return t?.kind === "op" && ops.includes(t.value) ? t.value : undefined;
  }

  private expectOp(op: string): void {
    if (!this.peekOp(op)) throw new SyntaxError(`Missing "${op}".`);
    this.pos++;
  }

  private comparison(): number {
    let left = this.additive();
    for (let op = this.peekOp("=", "<", ">", "<=", ">=", "<>"); op; ) {
      this.pos++;
      const right = this.additive();
      left = compare(op, left, right) ? 1 : 0;
      op = this.peekOp("=", "<", ">", "<=", ">=", "<>");
    }
    return left;
  }

  private additive(): number {
    let left = this.term();
    for (let op = this.peekOp("+", "-"); op; op = this.peekOp("+", "-")) {
      this.pos++;
      const right = this.term();
      left = op === "+" ? left + right : left - right;
    }
    return left;
  }

  private term(): number {
    let left = this.power();
    for (let op = this.peekOp("*", "/"); op; op = this.peekOp("*", "/")) {
      this.pos++;
      const right = this.power();
      if (op === "/" && right === 0) throw new RangeError("Divide by zero.");
      left = op === "*" ? left * right : left / right;
    }
    return left;
  }

  private power(): number {
    const base = this.unary();
    if (this.peekOp("^")) {
      this.pos++;
      return base ** this.power();
    }
    return base;
  }

  private unary(): number {
    if (this.peekOp("-")) {
      this.pos++;
      return -this.unary();
    }
    if (this.peekOp("+")) {
      this.pos++;
      return this.unary();
    }
    let value = this.primary();
    // A trailing % divides by 100, as in Word's formulas.
    while (this.peekOp("%")) {
      this.pos++;
      value /= 100;
    }
    return value;
  }

  private primary(): number {
    const t = this.tokens[this.pos];
    if (!t) throw new SyntaxError("Unexpected end of formula.");
    if (t.kind === "num") {
      this.pos++;
      return t.value;
    }
    if (t.kind === "op" && t.value === "(") {
      this.pos++;
      const value = this.comparison();
      this.expectOp(")");
      return value;
    }
    if (t.kind === "name") {
      this.pos++;
      const fn = FUNCTIONS[t.value];
      if (fn && this.peekOp("(")) {
        this.pos++;
        const args: Arg[] = [];
        if (!this.peekOp(")")) {
          args.push(this.argument());
          while (this.peekOp(",")) {
            this.pos++;
            args.push(this.argument());
          }
        }
        this.expectOp(")");
        return fn(args);
      }
      const constant = CONSTANTS[t.value];
      if (constant !== undefined) return constant;
      const range = this.reference(t.value);
      // A single cell reference is a number; an empty or text cell counts as 0.
      if (range) {
        if (t.value.includes(":") || POSITIONAL.has(t.value)) {
          throw new SyntaxError("A range can only be used inside a function.");
        }
        return range[0] ?? 0;
      }
    }
    throw new SyntaxError(`Syntax error at "${t.value}".`);
  }

  private argument(): Arg {
    const t = this.tokens[this.pos];
    if (t?.kind === "name" && !FUNCTIONS[t.value] && CONSTANTS[t.value] === undefined) {
      const next = this.tokens[this.pos + 1];
      const standsAlone =
        !next || (next.kind === "op" && (next.value === "," || next.value === ")"));
      const range = standsAlone ? this.reference(t.value) : undefined;
      if (range) {
        this.pos++;
        return range;
      }
    }
    return this.comparison();
  }

  /** The numbers a reference names, or undefined when the name is not a reference. */
  private reference(name: string): number[] | undefined {
    const { table } = this;
    const { row, column } = table.at;
    const collect = (cells: Array<[number, number]>): number[] =>
      cells.flatMap(([r, c]) => {
        const v = table.value(r, c);
        return v === undefined ? [] : [v];
      });
    if (POSITIONAL.has(name)) {
      const cells: Array<[number, number]> = [];
      if (name === "ABOVE") for (let r = 0; r < row; r++) cells.push([r, column]);
      if (name === "BELOW") for (let r = row + 1; r < table.rows; r++) cells.push([r, column]);
      if (name === "LEFT") for (let c = 0; c < column; c++) cells.push([row, c]);
      if (name === "RIGHT") for (let c = column + 1; c < table.columns; c++) cells.push([row, c]);
      return collect(cells);
    }
    const [from, to = from] = name.split(":");
    const a = from && CELL_REF.exec(from);
    const b = to && CELL_REF.exec(to);
    if (!a || !b) return undefined;
    const c1 = columnIndex(a[1]!);
    const r1 = Number(a[2]) - 1;
    const c2 = columnIndex(b[1]!);
    const r2 = Number(b[2]) - 1;
    const cells: Array<[number, number]> = [];
    for (let r = Math.min(r1, r2); r <= Math.max(r1, r2); r++) {
      for (let c = Math.min(c1, c2); c <= Math.max(c1, c2); c++) cells.push([r, c]);
    }
    return collect(cells);
  }
}

function compare(op: string, a: number, b: number): boolean {
  switch (op) {
    case "=":
      return a === b;
    case "<":
      return a < b;
    case ">":
      return a > b;
    case "<=":
      return a <= b;
    case ">=":
      return a >= b;
    default:
      return a !== b;
  }
}

/** Evaluate a formula (without the leading `=`). Throws on a syntax error. */
export function evaluateFormula(expression: string, table: FormulaTable): number {
  const source = expression.trim().replace(/^=/, "");
  return new Parser(tokenize(source), table).parse();
}

// Currency signs and grouping separators Word ignores when it reads a cell's
// number; a parenthesized amount is negative, as in accounting formats.
const CELL_NUMBER = /^\(?[-+]?[$€£¥]?\s*[-+]?(\d{1,3}(?:,\d{3})+|\d*)(\.\d+)?\s*%?\)?$/;

/** The number a cell's text holds, as Word reads it for a formula; undefined otherwise. */
export function cellNumber(text: string): number | undefined {
  const trimmed = text.trim();
  if (!trimmed || !CELL_NUMBER.test(trimmed)) return undefined;
  const digits = trimmed.replace(/[^\d.]/g, "");
  if (!/\d/.test(digits)) return undefined;
  const value = Number(digits);
  const negative = trimmed.includes("-") || (trimmed.startsWith("(") && trimmed.endsWith(")"));
  return negative ? -value : value;
}

// Up to ten significant digits, as Word shows a formula result without a picture.
const GENERAL_PRECISION = 10;

/**
 * Format a number with a numeric picture (`\#` switch, §17.16.4.2): `0` a
 * digit, `#` a digit or nothing, `.` the decimal point, `,` digit grouping,
 * text in quotes (or any other character) literal, and up to three
 * `;`-separated sections for positive, negative and zero values.
 */
export function formatFieldNumber(value: number, picture?: string): string {
  if (picture === undefined || picture === "") {
    return String(Number(value.toPrecision(GENERAL_PRECISION)));
  }
  const sections = picture.split(";");
  let section = sections[0] ?? "";
  let magnitude = value;
  let sign = value < 0 ? "-" : "";
  if (value < 0 && sections[1] !== undefined) {
    section = sections[1];
    magnitude = -value;
    sign = "";
  } else if (value === 0 && sections[2] !== undefined) {
    section = sections[2];
  } else {
    magnitude = Math.abs(value);
  }
  const first = section.search(/[0#]/);
  if (first < 0) return section.replace(/"/g, "");
  let last = first;
  for (let i = first; i < section.length; i++) if (/[0#.,]/.test(section[i]!)) last = i;
  while (last > first && !/[0#]/.test(section[last]!)) last--;
  const core = section.slice(first, last + 1);
  const prefix = section.slice(0, first).replace(/"/g, "");
  const suffix = section.slice(last + 1).replace(/"/g, "");
  const [intPart = "", fracPart = ""] = core.split(".");
  const decimals = fracPart.replace(/[^0#]/g, "").length;
  const minDecimals = fracPart.replace(/[^0]/g, "").length;
  const minInt = intPart.replace(/[^0]/g, "").length;
  let [whole = "0", frac = ""] = magnitude.toFixed(decimals).split(".");
  frac = frac.replace(/0+$/, "").padEnd(minDecimals, "0");
  whole = whole === "0" && minInt === 0 ? "" : whole.padStart(minInt, "0");
  if (intPart.includes(",")) whole = whole.replace(/\B(?=(\d{3})+(?!\d))/g, ",");
  const number = frac ? `${whole}.${frac}` : whole || "0";
  return `${sign}${prefix}${number}${suffix}`;
}
