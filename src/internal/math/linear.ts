/**
 * Office Math (OMML, ECMA-376 Part 1 §22.1) ⇄ linear format.
 *
 * The linear format is the UnicodeMath notation Word's equation editor
 * accepts (`a/b`, `x^2`, `√(x+1)`, `∑_(i=1)^n▒i`, `■(a&b@c&d)` …). This module
 * implements the subset that maps onto OMML's structures: fractions, scripts,
 * radicals, n-ary operators, delimiters, functions, accents, bars, limits,
 * matrices, equation arrays and boxes. Building is a two-step affair — parse to
 * a small tree, then emit `m:` elements — so the reverse direction (OMML back
 * to linear text, for re-editing) can share the operand rules.
 */

import type { XmlAttr, XmlElement, XmlNode } from "../xml/index.js";
import { XML_NAMESPACE } from "../xml/index.js";

export const OMML_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math";
const WML_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const XMLNS_URI = "http://www.w3.org/2000/xmlns/";
// Word writes every math run in Cambria Math, its OpenType math font.
const MATH_FONT = "Cambria Math";

type MathNode =
  | { readonly t: "row"; readonly items: readonly MathNode[] }
  | { readonly t: "text"; readonly s: string; readonly style?: "normal" | "plain" }
  | { readonly t: "frac"; readonly num: MathNode; readonly den: MathNode; readonly lin?: boolean }
  | {
      readonly t: "script";
      readonly base: MathNode;
      readonly sub?: MathNode;
      readonly sup?: MathNode;
    }
  | { readonly t: "rad"; readonly deg?: MathNode; readonly e: MathNode }
  | {
      readonly t: "nary";
      readonly chr: string;
      readonly sub?: MathNode;
      readonly sup?: MathNode;
      readonly e: MathNode;
    }
  | {
      readonly t: "delim";
      readonly open: string;
      readonly close: string;
      readonly items: readonly MathNode[];
    }
  | { readonly t: "func"; readonly name: MathNode; readonly e: MathNode }
  | { readonly t: "acc"; readonly chr: string; readonly e: MathNode }
  | { readonly t: "bar"; readonly top: boolean; readonly e: MathNode }
  | { readonly t: "lim"; readonly upper: boolean; readonly e: MathNode; readonly lim: MathNode }
  | { readonly t: "matrix"; readonly rows: readonly (readonly MathNode[])[] }
  | { readonly t: "eqArr"; readonly rows: readonly MathNode[] }
  | { readonly t: "box"; readonly e: MathNode };

/** Control words Word's math AutoCorrect expands (`\alpha` → α …). */
const CONTROL_WORDS: Readonly<Record<string, string>> = {
  alpha: "α",
  beta: "β",
  gamma: "γ",
  delta: "δ",
  epsilon: "ϵ",
  varepsilon: "ε",
  zeta: "ζ",
  eta: "η",
  theta: "θ",
  vartheta: "ϑ",
  iota: "ι",
  kappa: "κ",
  lambda: "λ",
  mu: "μ",
  nu: "ν",
  xi: "ξ",
  pi: "π",
  varpi: "ϖ",
  rho: "ρ",
  sigma: "σ",
  tau: "τ",
  upsilon: "υ",
  phi: "ϕ",
  varphi: "φ",
  chi: "χ",
  psi: "ψ",
  omega: "ω",
  Gamma: "Γ",
  Delta: "Δ",
  Theta: "Θ",
  Lambda: "Λ",
  Xi: "Ξ",
  Pi: "Π",
  Sigma: "Σ",
  Phi: "Φ",
  Psi: "Ψ",
  Omega: "Ω",
  times: "×",
  div: "÷",
  pm: "±",
  mp: "∓",
  cdot: "⋅",
  le: "≤",
  leq: "≤",
  ge: "≥",
  geq: "≥",
  ne: "≠",
  neq: "≠",
  approx: "≈",
  equiv: "≡",
  sim: "∼",
  cong: "≅",
  propto: "∝",
  infty: "∞",
  partial: "∂",
  nabla: "∇",
  in: "∈",
  notin: "∉",
  subset: "⊂",
  supset: "⊃",
  subseteq: "⊆",
  supseteq: "⊇",
  cup: "∪",
  cap: "∩",
  forall: "∀",
  exists: "∃",
  neg: "¬",
  wedge: "∧",
  vee: "∨",
  to: "→",
  rightarrow: "→",
  leftarrow: "←",
  leftrightarrow: "↔",
  Rightarrow: "⇒",
  Leftarrow: "⇐",
  Leftrightarrow: "⇔",
  ldots: "…",
  cdots: "⋯",
  hbar: "ℏ",
  ell: "ℓ",
  degree: "°",
  sum: "∑",
  prod: "∏",
  coprod: "∐",
  int: "∫",
  iint: "∬",
  iiint: "∭",
  oint: "∮",
  bigcup: "⋃",
  bigcap: "⋂",
  sqrt: "√",
  cbrt: "∛",
  qdrt: "∜",
  matrix: "■",
  eqarray: "█",
  box: "□",
  rect: "▭",
  overbar: "¯",
  underbar: "▁",
  below: "┬",
  above: "┴",
  naryand: "▒",
  hat: "̂",
  tilde: "̃",
  bar: "̅",
  dot: "̇",
  ddot: "̈",
  vec: "⃗",
  check: "̌",
  breve: "̆",
  acute: "́",
  grave: "̀",
  langle: "⟨",
  rangle: "⟩",
  lfloor: "⌊",
  rfloor: "⌋",
  lceil: "⌈",
  rceil: "⌉",
  norm: "‖",
  funcapply: "⁡",
};

const NARY_CHARS = "∑∏∐∫∬∭∮∯∰⋃⋂⋁⋀⨁⨂⨀";
const INTEGRALS = "∫∬∭∮∯∰";
const OPEN_DELIMS: Readonly<Record<string, string>> = {
  "(": ")",
  "[": "]",
  "{": "}",
  "⟨": "⟩",
  "⌊": "⌋",
  "⌈": "⌉",
  "|": "|",
  "‖": "‖",
};
const CLOSE_DELIMS = new Set(Object.values(OPEN_DELIMS));
const GROUP_OPEN = "〖";
const GROUP_CLOSE = "〗";
const DELIM_SEPARATOR = "│";
const OPERATORS = new Set("+-−=≠<>≤≥±∓×÷⋅·∙,;:!→←↔⇒⇐⇔∈∉⊂⊃⊆⊇∪∩≈≡∼≅∝∧∨¬∀∃…⋯");
const COMBINING_ACCENT = /[̀-ͯ⃐-⃿]/;
const FUNCTION_NAMES = [
  "arcsin",
  "arccos",
  "arctan",
  "sinh",
  "cosh",
  "tanh",
  "coth",
  "sin",
  "cos",
  "tan",
  "cot",
  "sec",
  "csc",
  "log",
  "ln",
  "lg",
  "exp",
  "lim",
  "max",
  "min",
  "sup",
  "inf",
  "det",
  "dim",
  "ker",
  "gcd",
  "arg",
  "deg",
  "Pr",
];
// Function names whose `_` / `^` scripts go below / above (limits).
const LIMIT_FUNCTIONS = new Set(["lim", "max", "min", "sup", "inf"]);
const FUNCTION_APPLY = "⁡";

/** Expand `\word` control words (Word's math AutoCorrect) into characters. */
function expandControlWords(src: string): string {
  return src.replace(/\\([A-Za-z]+) ?/g, (whole, word: string) => CONTROL_WORDS[word] ?? whole);
}

class LinearParser {
  private pos = 0;
  private readonly chars: string[];

  constructor(src: string) {
    this.chars = [...expandControlWords(src)];
  }

  parse(): MathNode {
    const row = this.expr(new Set());
    return row;
  }

  private peek(offset = 0): string | undefined {
    return this.chars[this.pos + offset];
  }

  private next(): string {
    return this.chars[this.pos++] ?? "";
  }

  private atEnd(): boolean {
    return this.pos >= this.chars.length;
  }

  /** A sequence of elements up to one of `stops` (not consumed) or the end. */
  private expr(stops: ReadonlySet<string>): MathNode {
    const items: MathNode[] = [];
    while (!this.atEnd()) {
      const ch = this.peek() ?? "";
      if (stops.has(ch)) break;
      if (ch === " ") {
        this.pos++;
        continue;
      }
      if (OPERATORS.has(ch)) {
        this.pos++;
        items.push({ t: "text", s: ch === "-" ? "−" : ch });
        continue;
      }
      if (CLOSE_DELIMS.has(ch) && !OPEN_DELIMS[ch]) {
        // A stray closing bracket is literal text.
        this.pos++;
        items.push({ t: "text", s: ch });
        continue;
      }
      let operand = this.operand(stops);
      while (this.peek() === "/") {
        this.pos++;
        const den = this.operand(stops);
        operand = { t: "frac", num: unwrap(operand), den: unwrap(den) };
      }
      items.push(operand);
    }
    return items.length === 1 && items[0] ? items[0] : { t: "row", items };
  }

  /** Consecutive scripted atoms with no operator or space between them. */
  private operand(stops: ReadonlySet<string>): MathNode {
    const items: MathNode[] = [];
    while (!this.atEnd()) {
      const ch = this.peek() ?? "";
      if (stops.has(ch) || ch === " " || ch === "/" || OPERATORS.has(ch)) break;
      if (CLOSE_DELIMS.has(ch) && !OPEN_DELIMS[ch]) break;
      if (ch === "▒" || ch === "&" || ch === "@") break;
      // `|` closes an enclosing |…| rather than opening a new one.
      if ((ch === "|" || ch === "‖") && stops.has(ch)) break;
      items.push(this.scripted(stops));
    }
    return items.length === 1 && items[0] ? items[0] : { t: "row", items };
  }

  /** One atom with its trailing scripts, accents and limits. */
  private scripted(stops: ReadonlySet<string>): MathNode {
    let base = this.atom(stops);
    for (;;) {
      const ch = this.peek();
      if (ch !== undefined && COMBINING_ACCENT.test(ch)) {
        this.pos++;
        base = { t: "acc", chr: ch, e: base };
        continue;
      }
      if (ch === "′" || ch === "″") {
        this.pos++;
        base = attachScript(base, "sup", { t: "text", s: ch });
        continue;
      }
      if (ch === "_" || ch === "^") {
        this.pos++;
        const script = unwrap(this.scriptOperand(stops));
        base = attachScript(base, ch === "_" ? "sub" : "sup", script);
        continue;
      }
      if (ch === "┬" || ch === "┴") {
        this.pos++;
        const lim = unwrap(this.scriptOperand(stops));
        base = { t: "lim", upper: ch === "┴", e: base, lim };
        continue;
      }
      return base;
    }
  }

  /**
   * The operand of `_`, `^`, `┬` …: one atom with its accents, so `x_i^2`
   * gives `x` both scripts rather than nesting `i^2`.
   */
  private scriptOperand(stops: ReadonlySet<string>): MathNode {
    if (this.peek() === "-" || this.peek() === "−" || this.peek() === "+") {
      // x^-1: a leading sign belongs to the script.
      const sign = this.next();
      const rest = this.scriptOperand(stops);
      return { t: "row", items: [{ t: "text", s: sign === "-" ? "−" : sign }, rest] };
    }
    let atom = this.atom(stops);
    for (;;) {
      const ch = this.peek();
      if (ch === undefined || !COMBINING_ACCENT.test(ch)) return atom;
      this.pos++;
      atom = { t: "acc", chr: ch, e: atom };
    }
  }

  private atom(stops: ReadonlySet<string>): MathNode {
    const ch = this.peek() ?? "";
    if (ch === GROUP_OPEN) {
      this.pos++;
      const inner = this.expr(new Set([GROUP_CLOSE]));
      if (this.peek() === GROUP_CLOSE) this.pos++;
      return { t: "row", items: inner.t === "row" ? inner.items : [inner] };
    }
    const close = OPEN_DELIMS[ch];
    if (close !== undefined) return this.delimited(ch, close);
    if (ch === '"') {
      this.pos++;
      let s = "";
      while (!this.atEnd() && this.peek() !== '"') s += this.next();
      this.pos++;
      return { t: "text", s, style: "normal" };
    }
    if (ch === "√" || ch === "∛" || ch === "∜") {
      this.pos++;
      return this.radical(ch, stops);
    }
    if (NARY_CHARS.includes(ch)) {
      this.pos++;
      return this.nary(ch, stops);
    }
    if (ch === "■" || ch === "█") {
      this.pos++;
      return this.array(ch === "■");
    }
    if (ch === "□" || ch === "▭") {
      this.pos++;
      return { t: "box", e: unwrap(this.scripted(stops)) };
    }
    if (ch === "¯" || ch === "▁") {
      this.pos++;
      return { t: "bar", top: ch === "¯", e: unwrap(this.scripted(stops)) };
    }
    if (/[0-9]/.test(ch)) {
      let s = "";
      while (/[0-9.]/.test(this.peek() ?? "")) {
        if (this.peek() === "." && !/[0-9]/.test(this.peek(1) ?? "")) break;
        s += this.next();
      }
      return { t: "text", s };
    }
    const fn = this.functionName();
    if (fn) return this.func(fn, stops);
    if (COMBINING_ACCENT.test(ch)) {
      // An accent typed before its base (`\hat x`): put it on the next atom.
      this.pos++;
      while (this.peek() === " ") this.pos++;
      return { t: "acc", chr: ch, e: unwrap(this.scripted(stops)) };
    }
    this.pos++;
    return { t: "text", s: ch };
  }

  private functionName(): string | undefined {
    for (const name of FUNCTION_NAMES) {
      const slice = this.chars.slice(this.pos, this.pos + name.length).join("");
      const after = this.chars[this.pos + name.length] ?? "";
      if (slice === name && !/[A-Za-z]/.test(after)) return name;
    }
    return undefined;
  }

  private func(name: string, stops: ReadonlySet<string>): MathNode {
    this.pos += name.length;
    let fName: MathNode = { t: "text", s: name, style: "plain" };
    // Scripts on the name: limits under/over for lim/max/min, else sub/sup.
    for (;;) {
      const ch = this.peek();
      if (ch !== "_" && ch !== "^" && ch !== "┬" && ch !== "┴") break;
      this.pos++;
      const script = unwrap(this.scriptOperand(stops));
      const asLimit = ch === "┬" || ch === "┴" || LIMIT_FUNCTIONS.has(name);
      fName = asLimit
        ? { t: "lim", upper: ch === "^" || ch === "┴", e: fName, lim: script }
        : attachScript(fName, ch === "_" ? "sub" : "sup", script);
    }
    while (this.peek() === " " || this.peek() === FUNCTION_APPLY) this.pos++;
    const arg = this.atEnd() || stops.has(this.peek() ?? "") ? emptyRow() : this.operand(stops);
    return { t: "func", name: fName, e: arg };
  }

  private delimited(open: string, close: string): MathNode {
    this.pos++;
    const items: MathNode[] = [];
    const stops = new Set([close, DELIM_SEPARATOR, "∣"]);
    for (;;) {
      items.push(this.expr(stops));
      const ch = this.peek();
      if (ch === DELIM_SEPARATOR || ch === "∣") {
        this.pos++;
        continue;
      }
      if (ch === close) this.pos++;
      break;
    }
    return { t: "delim", open, close, items };
  }

  private radical(ch: string, stops: ReadonlySet<string>): MathNode {
    const implicitDegree = ch === "∛" ? "3" : ch === "∜" ? "4" : undefined;
    if (this.peek() === "(") {
      this.pos++;
      const first = this.expr(new Set([")", "&"]));
      if (this.peek() === "&") {
        this.pos++;
        const e = this.expr(new Set([")"]));
        if (this.peek() === ")") this.pos++;
        return { t: "rad", deg: first, e };
      }
      if (this.peek() === ")") this.pos++;
      return implicitDegree
        ? { t: "rad", deg: { t: "text", s: implicitDegree }, e: first }
        : { t: "rad", e: first };
    }
    const e = unwrap(this.scripted(stops));
    return implicitDegree
      ? { t: "rad", deg: { t: "text", s: implicitDegree }, e }
      : { t: "rad", e };
  }

  private nary(chr: string, stops: ReadonlySet<string>): MathNode {
    let sub: MathNode | undefined;
    let sup: MathNode | undefined;
    for (;;) {
      const ch = this.peek();
      if (ch === "_") {
        this.pos++;
        sub = unwrap(this.scriptOperand(stops));
      } else if (ch === "^") {
        this.pos++;
        sup = unwrap(this.scriptOperand(stops));
      } else break;
    }
    if (this.peek() === "▒") this.pos++;
    while (this.peek() === " ") this.pos++;
    const e =
      this.atEnd() || stops.has(this.peek() ?? "") ? emptyRow() : unwrapGroup(this.operand(stops));
    return {
      t: "nary",
      chr,
      ...(sub ? { sub } : {}),
      ...(sup ? { sup } : {}),
      e,
    };
  }

  private array(isMatrix: boolean): MathNode {
    if (this.peek() !== "(") return { t: "text", s: isMatrix ? "■" : "█" };
    this.pos++;
    const rows: MathNode[][] = [[]];
    const stops = new Set([")", "&", "@"]);
    for (;;) {
      const cell = this.expr(stops);
      rows[rows.length - 1]?.push(cell);
      const ch = this.next();
      if (ch === "&") continue;
      if (ch === "@") {
        rows.push([]);
        continue;
      }
      break;
    }
    if (isMatrix) return { t: "matrix", rows };
    return {
      t: "eqArr",
      rows: rows.map((r) => (r.length === 1 && r[0] ? r[0] : { t: "row", items: r })),
    };
  }
}

function emptyRow(): MathNode {
  return { t: "row", items: [] };
}

/** Parentheses around a script / fraction operand only group it (UnicodeMath). */
function unwrap(node: MathNode): MathNode {
  if (node.t === "delim" && node.open === "(" && node.close === ")" && node.items.length === 1) {
    return node.items[0] ?? emptyRow();
  }
  return node;
}

function unwrapGroup(node: MathNode): MathNode {
  return node;
}

function attachScript(base: MathNode, kind: "sub" | "sup", script: MathNode): MathNode {
  if (base.t === "script" && base[kind] === undefined) return { ...base, [kind]: script };
  return { t: "script", base, [kind]: script };
}

// --- OMML emission ------------------------------------------------------------

function mAttr(local: string, value: string): XmlAttr {
  return { name: { uri: OMML_NS, local, prefix: "m" }, value, isNamespaceDecl: false };
}

function mEl(local: string, children: readonly XmlNode[] = [], attrs: XmlAttr[] = []): XmlElement {
  return {
    kind: "element",
    name: { uri: OMML_NS, local, prefix: "m" },
    attrs,
    children,
    xmlSpace: "default",
    selfClosing: children.length === 0,
  };
}

function mVal(local: string, value: string): XmlElement {
  return mEl(local, [], [mAttr("val", value)]);
}

function wEl(local: string, attrs: XmlAttr[], children: readonly XmlNode[] = []): XmlElement {
  return {
    kind: "element",
    name: { uri: WML_NS, local, prefix: "w" },
    attrs,
    children,
    xmlSpace: "default",
    selfClosing: children.length === 0,
  };
}

function wAttr(local: string, value: string): XmlAttr {
  return { name: { uri: WML_NS, local, prefix: "w" }, value, isNamespaceDecl: false };
}

function mathRun(text: string, style?: "normal" | "plain"): XmlElement {
  const preserve = /^\s|\s$/.test(text);
  const t: XmlElement = {
    kind: "element",
    name: { uri: OMML_NS, local: "t", prefix: "m" },
    attrs: preserve
      ? [
          {
            name: { uri: XML_NAMESPACE, local: "space", prefix: "xml" },
            value: "preserve",
            isNamespaceDecl: false,
          },
        ]
      : [],
    children: [{ kind: "text", value: text }],
    xmlSpace: preserve ? "preserve" : "default",
    selfClosing: false,
  };
  const children: XmlElement[] = [];
  if (style === "normal") children.push(mEl("rPr", [mEl("nor")]));
  else if (style === "plain") children.push(mEl("rPr", [mVal("sty", "p")]));
  children.push(
    wEl("rPr", [], [wEl("rFonts", [wAttr("ascii", MATH_FONT), wAttr("hAnsi", MATH_FONT)])]),
  );
  children.push(t);
  return mEl("r", children);
}

/** The children an argument element (`m:e`, `m:num` …) holds for a node. */
function emitArg(node: MathNode): XmlElement[] {
  if (node.t !== "row") return emit(node);
  // Adjacent plain text items share one run, as Word writes them.
  const out: XmlElement[] = [];
  let pending = "";
  const flush = (): void => {
    if (pending) out.push(mathRun(pending));
    pending = "";
  };
  for (const item of node.items) {
    if (item.t === "text" && item.style === undefined) {
      pending += item.s;
      continue;
    }
    flush();
    out.push(...emitArg(item));
  }
  flush();
  return out;
}

function arg(local: string, node: MathNode | undefined): XmlElement {
  return mEl(local, node ? emitArg(node) : []);
}

function emit(node: MathNode): XmlElement[] {
  switch (node.t) {
    case "row":
      return emitArg(node);
    case "text":
      return [mathRun(node.s, node.style)];
    case "frac":
      return [
        mEl("f", [
          ...(node.lin ? [mEl("fPr", [mVal("type", "lin")])] : []),
          arg("num", node.num),
          arg("den", node.den),
        ]),
      ];
    case "script": {
      if (node.sub && node.sup) {
        return [mEl("sSubSup", [arg("e", node.base), arg("sub", node.sub), arg("sup", node.sup)])];
      }
      if (node.sub) return [mEl("sSub", [arg("e", node.base), arg("sub", node.sub)])];
      return [mEl("sSup", [arg("e", node.base), arg("sup", node.sup)])];
    }
    case "rad":
      return [
        mEl("rad", [
          ...(node.deg ? [] : [mEl("radPr", [mVal("degHide", "1")])]),
          arg("deg", node.deg),
          arg("e", node.e),
        ]),
      ];
    case "nary": {
      const pr: XmlElement[] = [];
      // ∫ is the schema default for m:chr; Word omits it for integrals.
      if (node.chr !== "∫") pr.push(mVal("chr", node.chr));
      pr.push(mVal("limLoc", INTEGRALS.includes(node.chr) ? "subSup" : "undOvr"));
      if (!node.sub) pr.push(mVal("subHide", "1"));
      if (!node.sup) pr.push(mVal("supHide", "1"));
      return [
        mEl("nary", [
          mEl("naryPr", pr),
          arg("sub", node.sub),
          arg("sup", node.sup),
          arg("e", node.e),
        ]),
      ];
    }
    case "delim": {
      // dPr children in schema order: begChr, sepChr, endChr.
      const pr: XmlElement[] = [];
      if (node.open !== "(") pr.push(mVal("begChr", node.open));
      if (node.items.length > 1) pr.push(mVal("sepChr", DELIM_SEPARATOR));
      if (node.close !== ")") pr.push(mVal("endChr", node.close));
      return [
        mEl("d", [...(pr.length ? [mEl("dPr", pr)] : []), ...node.items.map((i) => arg("e", i))]),
      ];
    }
    case "func":
      return [mEl("func", [arg("fName", node.name), arg("e", node.e)])];
    case "acc":
      return [mEl("acc", [mEl("accPr", [mVal("chr", node.chr)]), arg("e", node.e)])];
    case "bar":
      return [
        mEl("bar", [mEl("barPr", [mVal("pos", node.top ? "top" : "bot")]), arg("e", node.e)]),
      ];
    case "lim":
      return [mEl(node.upper ? "limUpp" : "limLow", [arg("e", node.e), arg("lim", node.lim)])];
    case "matrix":
      return [
        mEl(
          "m",
          node.rows.map((row) =>
            mEl(
              "mr",
              row.map((cell) => arg("e", cell)),
            ),
          ),
        ),
      ];
    case "eqArr":
      return [
        mEl(
          "eqArr",
          node.rows.map((row) => arg("e", row)),
        ),
      ];
    case "box":
      return [mEl("borderBox", [arg("e", node.e)])];
    default: {
      const unhandled: never = node;
      return unhandled;
    }
  }
}

function withMathNamespace(el: XmlElement): XmlElement {
  return {
    ...el,
    attrs: [
      {
        name: { uri: XMLNS_URI, local: "m", prefix: "xmlns" },
        value: OMML_NS,
        isNamespaceDecl: true,
      },
      ...el.attrs,
    ],
  };
}

/**
 * Build an `<m:oMath>` (inline) or `<m:oMathPara>` (display, on its own line)
 * element from linear-format text. The element declares the `m:` namespace
 * itself, so it is valid wherever it is spliced (document, header, footnote).
 */
export function buildMathFromLinear(linear: string, display = false): XmlElement {
  const tree = new LinearParser(linear).parse();
  const oMath = mEl("oMath", emitArg(tree));
  if (!display) return withMathNamespace(oMath);
  return withMathNamespace(mEl("oMathPara", [oMath]));
}

// --- OMML → linear --------------------------------------------------------------

declare const MATH_ELEMENT: unique symbol;
// Branded so the guard's false branch keeps other elements (see isW in
// src/api/insert-xml.ts for the same pattern).
type MathElement = XmlElement & { readonly [MATH_ELEMENT]: true };

function isM(node: XmlNode, local: string): node is MathElement {
  return node.kind === "element" && node.name.uri === OMML_NS && node.name.local === local;
}

function child(el: XmlElement, local: string): XmlElement | undefined {
  return el.children.find((c): c is XmlElement => isM(c, local));
}

function propVal(el: XmlElement, prLocal: string, local: string): string | undefined {
  const pr = child(el, prLocal);
  const prop = pr ? child(pr, local) : undefined;
  if (!prop) return undefined;
  return prop.attrs.find((a) => a.name.local === "val")?.value ?? "1";
}

function textOf(el: XmlElement): string {
  let out = "";
  for (const c of el.children) {
    if (c.kind === "text") out += c.value;
    else if (c.kind === "element") out += textOf(c);
  }
  return out;
}

/** Linear text of an argument element's content. */
function linearArg(el: XmlElement | undefined): string {
  if (!el) return "";
  return joinLinear(el.children.filter((c): c is XmlElement => c.kind === "element"));
}

// After these, a space ends their operand (an n-ary body, a function
// argument, a denominator), so following content is not read back into it.
const OPERAND_TAKING = new Set(["nary", "func", "f"]);

function joinLinear(children: readonly XmlElement[]): string {
  let out = "";
  for (const [i, c] of children.entries()) {
    out += linearOf(c);
    if (i < children.length - 1 && c.name.uri === OMML_NS && OPERAND_TAKING.has(c.name.local)) {
      out += " ";
    }
  }
  return out;
}

/** Wrap an operand in parentheses when it would not parse back as one atom. */
function operand(el: XmlElement | undefined): string {
  const s = linearArg(el);
  if ([...s].length <= 1) return s;
  if (/^[0-9.]+$/.test(s) || /^[A-Za-z]$/.test(s)) return s;
  const only = el?.children.filter((c): c is XmlElement => c.kind === "element");
  if (only?.length === 1 && only[0] && !isM(only[0], "r") && !isM(only[0], "f")) {
    const k = only[0].name.local;
    // A delimiter already brackets itself; scripts/radicals bind tightly.
    if (k === "rad" || k === "acc" || k === "m") return s;
    if (k === "d") return `(${s})`;
  }
  return `(${s})`;
}

/**
 * The base of a script or accent. The parser binds a script to the single
 * atom before it and keeps a bracketed base as a delimiter (`(x+a)^n`), so
 * a delimiter, a function name or one character is written as-is and
 * anything longer is grouped with 〖…〗 (which adds no brackets).
 */
function baseOperand(el: XmlElement | undefined): string {
  const s = linearArg(el);
  if ([...s].length <= 1 || FUNCTION_NAMES.includes(s)) return s;
  const only = el?.children.filter((c): c is XmlElement => c.kind === "element");
  if (only?.length === 1 && only[0] && !isM(only[0], "r") && !isM(only[0], "f")) return s;
  return `${GROUP_OPEN}${s}${GROUP_CLOSE}`;
}

function linearOf(el: XmlElement): string {
  if (el.name.uri !== OMML_NS) return "";
  switch (el.name.local) {
    case "r": {
      const t = child(el, "t");
      const text = t ? textOf(t) : "";
      const rPr = child(el, "rPr");
      return rPr && child(rPr, "nor") ? `"${text}"` : text;
    }
    case "f": {
      const lin = propVal(el, "fPr", "type");
      const sep = lin === "lin" ? "/" : "/";
      return `${operand(child(el, "num"))}${sep}${operand(child(el, "den"))}`;
    }
    case "sSup":
      return `${baseOperand(child(el, "e"))}^${operand(child(el, "sup"))}`;
    case "sSub":
      return `${baseOperand(child(el, "e"))}_${operand(child(el, "sub"))}`;
    case "sSubSup":
      return `${baseOperand(child(el, "e"))}_${operand(child(el, "sub"))}^${operand(child(el, "sup"))}`;
    case "sPre":
      return `〖_${operand(child(el, "sub"))}^${operand(child(el, "sup"))}〗${baseOperand(child(el, "e"))}`;
    case "rad": {
      const degHidden = propVal(el, "radPr", "degHide");
      const deg = linearArg(child(el, "deg"));
      const e = linearArg(child(el, "e"));
      return degHidden === "1" || degHidden === "on" || deg === "" ? `√(${e})` : `√(${deg}&${e})`;
    }
    case "nary": {
      const chr = propVal(el, "naryPr", "chr") ?? "∫";
      const sub = linearArg(child(el, "sub"));
      const sup = linearArg(child(el, "sup"));
      const subPart = sub ? `_${operand(child(el, "sub"))}` : "";
      const supPart = sup ? `^${operand(child(el, "sup"))}` : "";
      // The parser keeps a bracketed body as a delimiter, like a script base.
      return `${chr}${subPart}${supPart}▒${baseOperand(child(el, "e"))}`;
    }
    case "d": {
      const open = propVal(el, "dPr", "begChr") ?? "(";
      const close = propVal(el, "dPr", "endChr") ?? ")";
      const items = el.children.filter((c): c is XmlElement => isM(c, "e")).map(linearArg);
      return `${open}${items.join(DELIM_SEPARATOR)}${close}`;
    }
    case "func": {
      const name = linearArg(child(el, "fName"));
      const e = linearArg(child(el, "e"));
      const single = [...e].length <= 1 || /^\(.*\)$/.test(e);
      return single ? `${name} ${e}` : `${name}${GROUP_OPEN}${e}${GROUP_CLOSE}`;
    }
    case "acc": {
      const chr = propVal(el, "accPr", "chr") ?? "̂";
      return `${baseOperand(child(el, "e"))}${chr}`;
    }
    case "bar": {
      const pos = propVal(el, "barPr", "pos") ?? "bot";
      return `${pos === "top" ? "¯" : "▁"}(${linearArg(child(el, "e"))})`;
    }
    case "limLow":
      return `${linearArg(child(el, "e"))}┬${operand(child(el, "lim"))}`;
    case "limUpp":
      return `${linearArg(child(el, "e"))}┴${operand(child(el, "lim"))}`;
    case "m": {
      const rows = el.children
        .filter((c): c is XmlElement => isM(c, "mr"))
        .map((mr) =>
          mr.children
            .filter((c): c is XmlElement => isM(c, "e"))
            .map(linearArg)
            .join("&"),
        );
      return `■(${rows.join("@")})`;
    }
    case "eqArr":
      return `█(${el.children
        .filter((c): c is XmlElement => isM(c, "e"))
        .map(linearArg)
        .join("@")})`;
    case "borderBox":
      return `▭(${linearArg(child(el, "e"))})`;
    case "box":
      return `□(${linearArg(child(el, "e"))})`;
    case "oMathPara":
      return el.children
        .filter((c): c is XmlElement => isM(c, "oMath"))
        .map(linearOf)
        .join("@");
    default:
      // oMath, groupChr, phant, e … : their argument content.
      return joinLinear(
        el.children
          .filter((c): c is XmlElement => c.kind === "element")
          .filter((c) => !c.name.local.endsWith("Pr")),
      );
  }
}

/** The linear-format text of an `<m:oMath>` / `<m:oMathPara>` element. */
export function mathToLinear(el: XmlElement): string {
  return linearOf(el);
}
