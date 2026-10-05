/**
 * OMML (Office Math, ECMA-376 Part 1 §22.1) → MathML for the canvas.
 *
 * Browsers render MathML Core natively, so each OMML structure maps onto its
 * MathML counterpart: fractions → `mfrac`, scripts → `msub`/`msup`, n-ary
 * operators → `munderover` / `msubsup`, delimiters → fenced `mrow`s, matrices
 * → `mtable` and so on. The output is read-only on the canvas; equations are
 * edited through the equation dialog (linear format).
 */

import type { XmlElement, XmlNode } from "@office-kit/docx";

export const OMML_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math";

function escapeXml(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function isM(node: XmlNode, local: string): boolean {
  return node.kind === "element" && node.name.uri === OMML_NS && node.name.local === local;
}

function child(el: XmlElement | undefined, local: string): XmlElement | undefined {
  return el?.children.find((c): c is XmlElement => c.kind === "element" && isM(c, local));
}

function children(el: XmlElement, local: string): XmlElement[] {
  return el.children.filter((c): c is XmlElement => c.kind === "element" && isM(c, local));
}

/** A property's `m:val` (on/off properties without a value count as on). */
function prop(el: XmlElement, prLocal: string, local: string): string | undefined {
  const p = child(child(el, prLocal), local);
  if (!p) return undefined;
  return p.attrs.find((a) => a.name.local === "val")?.value ?? "1";
}

function on(value: string | undefined): boolean {
  return value === "1" || value === "on" || value === "true";
}

// Characters MathML should treat as operators (stretchy fences, relations …).
const OPERATOR_CHARS = /[+\-−=≠<>≤≥±∓×÷⋅·∙,;:!→←↔⇒⇐⇔∈∉⊂⊃⊆⊇∪∩≈≡∼≅∝∧∨¬∀∃()[\]{}|‖⟨⟩⌊⌋⌈⌉∑∏∫∮/…⋯]/;

/** Split a math run's text into `mi` / `mn` / `mo` tokens. */
function tokens(text: string, plain: boolean): string {
  if (plain) return `<mtext>${escapeXml(text)}</mtext>`;
  let out = "";
  const re = /(\d+(?:\.\d+)?)|(\s+)|(.)/gsu;
  for (const m of text.matchAll(re)) {
    const [whole, num, space, ch] = m;
    if (num) out += `<mn>${num}</mn>`;
    else if (space) out += `<mspace width="${0.25 * space.length}em"/>`;
    else if (ch && OPERATOR_CHARS.test(ch)) out += `<mo>${escapeXml(ch)}</mo>`;
    else if (ch) out += `<mi>${escapeXml(ch)}</mi>`;
    else out += escapeXml(whole);
  }
  return out;
}

function run(el: XmlElement): string {
  const t = child(el, "t");
  const text = t ? t.children.map((c) => (c.kind === "text" ? c.value : "")).join("") : "";
  const rPr = child(el, "rPr");
  const normal = !!child(rPr, "nor");
  const sty = child(rPr, "sty")?.attrs.find((a) => a.name.local === "val")?.value;
  // Multi-letter upright text (function names) is one identifier.
  if (sty === "p" && /^\p{L}+$/u.test(text))
    return `<mi mathvariant="normal">${escapeXml(text)}</mi>`;
  return tokens(text, normal);
}

/** The MathML for an argument element's content, as one `mrow`. */
function arg(el: XmlElement | undefined): string {
  if (!el) return "<mrow></mrow>";
  return `<mrow>${el.children
    .filter((c): c is XmlElement => c.kind === "element")
    .map(convert)
    .join("")}</mrow>`;
}

function fence(chr: string | undefined, fallback: string): string {
  const c = chr ?? fallback;
  return c === "" ? "" : `<mo fence="true" stretchy="true">${escapeXml(c)}</mo>`;
}

function convert(el: XmlElement): string {
  if (el.name.uri !== OMML_NS) return "";
  switch (el.name.local) {
    case "r":
      return run(el);
    case "f": {
      const type = prop(el, "fPr", "type");
      const num = arg(child(el, "num"));
      const den = arg(child(el, "den"));
      if (type === "lin") return `<mrow>${num}<mo>/</mo>${den}</mrow>`;
      if (type === "noBar") return `<mfrac linethickness="0">${num}${den}</mfrac>`;
      return `<mfrac>${num}${den}</mfrac>`;
    }
    case "sSup":
      return `<msup>${arg(child(el, "e"))}${arg(child(el, "sup"))}</msup>`;
    case "sSub":
      return `<msub>${arg(child(el, "e"))}${arg(child(el, "sub"))}</msub>`;
    case "sSubSup":
      return `<msubsup>${arg(child(el, "e"))}${arg(child(el, "sub"))}${arg(child(el, "sup"))}</msubsup>`;
    case "sPre":
      return `<mmultiscripts>${arg(child(el, "e"))}<mprescripts/>${arg(child(el, "sub"))}${arg(child(el, "sup"))}</mmultiscripts>`;
    case "rad": {
      const e = arg(child(el, "e"));
      const deg = child(el, "deg");
      const hidden = on(prop(el, "radPr", "degHide")) || !deg || deg.children.length === 0;
      return hidden ? `<msqrt>${e}</msqrt>` : `<mroot>${e}${arg(deg)}</mroot>`;
    }
    case "nary": {
      const chr = prop(el, "naryPr", "chr") ?? "∫";
      const op = `<mo largeop="true" movablelimits="false">${escapeXml(chr)}</mo>`;
      const sub = on(prop(el, "naryPr", "subHide")) ? undefined : arg(child(el, "sub"));
      const sup = on(prop(el, "naryPr", "supHide")) ? undefined : arg(child(el, "sup"));
      const integral = /[∫∬∭∮∯∰]/.test(chr);
      const under = (prop(el, "naryPr", "limLoc") ?? (integral ? "subSup" : "undOvr")) === "undOvr";
      let base = op;
      if (sub && sup)
        base = under
          ? `<munderover>${op}${sub}${sup}</munderover>`
          : `<msubsup>${op}${sub}${sup}</msubsup>`;
      else if (sub) base = under ? `<munder>${op}${sub}</munder>` : `<msub>${op}${sub}</msub>`;
      else if (sup) base = under ? `<mover>${op}${sup}</mover>` : `<msup>${op}${sup}</msup>`;
      return `<mrow>${base}${arg(child(el, "e"))}</mrow>`;
    }
    case "d": {
      const open = fence(prop(el, "dPr", "begChr"), "(");
      const close = fence(prop(el, "dPr", "endChr"), ")");
      const sep = prop(el, "dPr", "sepChr") ?? "|";
      const items = children(el, "e").map(arg);
      return `<mrow>${open}${items.join(`<mo separator="true">${escapeXml(sep)}</mo>`)}${close}</mrow>`;
    }
    case "func":
      return `<mrow>${arg(child(el, "fName"))}<mo>&#x2061;</mo>${arg(child(el, "e"))}</mrow>`;
    case "acc": {
      const chr = prop(el, "accPr", "chr") ?? "̂";
      // MathML wants the spacing form of a combining accent.
      const spacing = ACCENT_SPACING[chr] ?? chr;
      return `<mover accent="true">${arg(child(el, "e"))}<mo>${escapeXml(spacing)}</mo></mover>`;
    }
    case "bar": {
      const top = (prop(el, "barPr", "pos") ?? "bot") === "top";
      return top
        ? `<mover accent="true">${arg(child(el, "e"))}<mo>‾</mo></mover>`
        : `<munder accentunder="true">${arg(child(el, "e"))}<mo>_</mo></munder>`;
    }
    case "groupChr": {
      const chr = prop(el, "groupChrPr", "chr") ?? "⏟";
      const top = prop(el, "groupChrPr", "pos") === "top";
      const mo = `<mo stretchy="true">${escapeXml(chr)}</mo>`;
      return top
        ? `<mover>${arg(child(el, "e"))}${mo}</mover>`
        : `<munder>${arg(child(el, "e"))}${mo}</munder>`;
    }
    case "limLow":
      return `<munder>${arg(child(el, "e"))}${arg(child(el, "lim"))}</munder>`;
    case "limUpp":
      return `<mover>${arg(child(el, "e"))}${arg(child(el, "lim"))}</mover>`;
    case "m":
      return `<mtable>${children(el, "mr")
        .map(
          (mr) =>
            `<mtr>${children(mr, "e")
              .map((e) => `<mtd>${arg(e)}</mtd>`)
              .join("")}</mtr>`,
        )
        .join("")}</mtable>`;
    case "eqArr":
      return `<mtable columnalign="left">${children(el, "e")
        .map((e) => `<mtr><mtd>${arg(e)}</mtd></mtr>`)
        .join("")}</mtable>`;
    case "borderBox":
      return `<menclose notation="box">${arg(child(el, "e"))}</menclose>`;
    case "phant":
      return on(prop(el, "phantPr", "show"))
        ? arg(child(el, "e"))
        : `<mphantom>${arg(child(el, "e"))}</mphantom>`;
    case "box":
    case "e":
    case "oMath":
      return arg(el);
    default:
      // Property elements and anything unmodelled: render its arguments.
      return el.name.local.endsWith("Pr") ? "" : arg(el);
  }
}

const ACCENT_SPACING: Readonly<Record<string, string>> = {
  "̀": "`",
  "́": "´",
  "̂": "^",
  "̃": "~",
  "̄": "¯",
  "̅": "‾",
  "̆": "˘",
  "̇": "˙",
  "̈": "¨",
  "̌": "ˇ",
  "⃗": "→",
};

/**
 * Render an `<m:oMath>` (inline) or `<m:oMathPara>` (display) element as a
 * MathML `<math>` element. `attrs` is extra markup for the wrapper (anchors).
 */
export function renderMath(el: XmlElement, attrs = ""): string {
  if (isM(el, "oMathPara")) {
    const maths = children(el, "oMath")
      .map((m) => `<math display="block">${arg(m)}</math>`)
      .join("");
    return `<span class="wk-math wk-math-display" contenteditable="false" ${attrs}>${maths}</span>`;
  }
  return `<span class="wk-math" contenteditable="false" ${attrs}><math>${arg(el)}</math></span>`;
}

/** Whether a raw inline is an equation the canvas renders as math. */
export function isMathElement(el: XmlElement): boolean {
  return isM(el, "oMath") || isM(el, "oMathPara");
}
