/**
 * Field results on the canvas. A complex field is a run sequence
 * `fldChar begin · instrText… · fldChar separate · result runs · fldChar end`
 * (ECMA-376 Part 1 §17.16.18); a simple field is one `<w:fldSimple w:instr>`.
 * The renderer tags every result run with the field's type, so the page
 * layout can fill in PAGE / NUMPAGES … per page and the UI can style fields
 * (Mailings ▸ Highlight Merge Fields).
 *
 * The field characters and code render as hidden, non-editable spans and the
 * result as ordinary editable text (`wk-fresult`), so Toggle Field Codes is
 * pure CSS: the editor root gets a class that hides results and shows
 * `{ CODE }` instead, without re-rendering or moving the caret.
 */

import type { WmlParagraph, WmlRun, XmlElement } from "@office-kit/docx";
import { WML_NS } from "./wml-ns.js";

/** Fields whose result depends on the page the text lands on. */
export const PAGE_FIELDS: ReadonlySet<string> = new Set([
  "PAGE",
  "NUMPAGES",
  "SECTIONPAGES",
  "SECTION",
]);

/** The field type of an instruction: its first word, upper-cased (`PAGE \* roman` → `PAGE`). */
export function fieldType(instruction: string): string {
  return instruction.trim().split(/\s+/)[0]?.toUpperCase() ?? "";
}

/** The `\*` general format switch of an instruction (`PAGE \* ROMAN` → `ROMAN`). */
export function fieldFormatSwitch(instruction: string): string | undefined {
  return /\\\*\s*(\w+)/.exec(instruction)?.[1];
}

export interface FieldResult {
  readonly type: string;
  readonly instruction: string;
}

/**
 * For each child index of the paragraph that is a run inside a field result,
 * the innermost field it belongs to. Fields are tracked within the paragraph
 * only: a field spanning paragraphs (a TOC) is tagged in its first paragraph.
 */
export function fieldResultRuns(para: WmlParagraph): Map<number, FieldResult> {
  const out = new Map<number, FieldResult>();
  const stack: Array<{ instruction: string; inResult: boolean }> = [];
  para.children.forEach((child, index) => {
    if (child.kind !== "run") return;
    const top = stack.at(-1);
    if (top?.inResult && child.pieces.some((p) => p.kind === "text")) {
      out.set(index, { type: fieldType(top.instruction), instruction: top.instruction });
    }
    for (const piece of child.pieces) {
      const current = stack.at(-1);
      if (piece.kind === "fieldChar") {
        if (piece.charType === "begin") stack.push({ instruction: "", inResult: false });
        else if (piece.charType === "separate" && current) current.inResult = true;
        else if (piece.charType === "end") stack.pop();
      } else if (piece.kind === "instrText" && current && !current.inResult) {
        current.instruction += piece.value;
      }
    }
  });
  return out;
}

/** The instruction of a `<w:fldSimple>`, or `undefined` for any other element. */
export function simpleFieldInstruction(el: XmlElement): string | undefined {
  if (el.name.uri !== WML_NS || el.name.local !== "fldSimple") return undefined;
  return el.attrs.find((a) => a.name.local === "instr")?.value;
}

function escapeAttr(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;");
}

/**
 * Attributes for a field result element. Page fields are read-only: the
 * layout rewrites their text on every page, so typing into one would be lost.
 */
export function fieldAttrs(field: FieldResult): string {
  const attrs = `data-wk-field="${escapeAttr(field.type)}" data-wk-instr="${escapeAttr(field.instruction.trim())}"`;
  return PAGE_FIELDS.has(field.type) ? `${attrs} contenteditable="false"` : attrs;
}

export type FieldRunRole = "begin" | "code" | "separate" | "result" | "end";

/** Classify each run of a paragraph by its part in a complex field. */
export function fieldRunRoles(para: WmlParagraph): Map<WmlRun, FieldRunRole> {
  const roles = new Map<WmlRun, FieldRunRole>();
  // Per open field: whether its separator has been seen.
  const stack: boolean[] = [];
  for (const child of para.children) {
    if (child.kind !== "run") continue;
    let role: FieldRunRole | undefined;
    for (const piece of child.pieces) {
      if (piece.kind === "fieldChar") {
        if (piece.charType === "begin") {
          stack.push(false);
          role = "begin";
        } else if (piece.charType === "separate") {
          stack[stack.length - 1] = true;
          role = "separate";
        } else {
          stack.pop();
          role = "end";
        }
      } else if (piece.kind === "instrText" && stack.length) {
        role ??= "code";
      }
    }
    if (role === undefined && stack.length) {
      // Text inside a field: its result once separated, else part of the code.
      role = stack[stack.length - 1] ? "result" : "code";
    }
    if (role) roles.set(child, role);
  }
  return roles;
}

function escapeHtml(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

/**
 * HTML for a field's structural run (begin / code / separate / end), or
 * undefined for runs the normal run renderer draws (results, plain text).
 */
export function renderFieldStructureRun(
  run: WmlRun,
  role: FieldRunRole | undefined,
): string | undefined {
  switch (role) {
    case "begin":
      return `<span class="wk-fchar" contenteditable="false">{</span>`;
    case "end":
      return `<span class="wk-fchar" contenteditable="false">}</span>`;
    case "separate":
      return "";
    case "code": {
      const code = run.pieces
        .map((p) => (p.kind === "instrText" || p.kind === "text" ? p.value : ""))
        .join("");
      return `<span class="wk-fcode" contenteditable="false">${escapeHtml(code)}</span>`;
    }
    default:
      return undefined;
  }
}

// The Symbol font's letters are Greek (U+F041 'A' → Α …): Word stores the
// font's private-use code; browsers without the font need the Unicode letter.
const SYMBOL_FONT_GREEK = "ΑΒΧΔΕΦΓΗΙϑΚΛΜΝΟΠΘΡΣΤΥςΩΞΨΖ[∴]⊥_‾αβχδεφγηιϕκλμνοπθρστυϖωξψζ";
const SYMBOL_FONT_FIRST = 0x41;
const PRIVATE_USE_SYMBOL_BASE = 0xf000;

/**
 * The character a symbol-font code shows as in a browser: `char` is the
 * `w:sym/@w:char` hex code (with or without the U+F000 private-use offset).
 */
export function symbolGlyph(font: string, char: string): string {
  let code = Number.parseInt(char, 16);
  if (!Number.isFinite(code)) return "";
  if (code >= PRIVATE_USE_SYMBOL_BASE) code -= PRIVATE_USE_SYMBOL_BASE;
  const greek =
    font.toLowerCase() === "symbol" ? SYMBOL_FONT_GREEK[code - SYMBOL_FONT_FIRST] : undefined;
  return greek ?? String.fromCodePoint(code);
}

/**
 * HTML for a run holding `<w:sym>` characters (Insert ▸ Symbol from a symbol
 * font), shown non-editable in that font; undefined for other runs.
 */
export function renderSymbolRun(run: WmlRun, css: string): string | undefined {
  const syms = run.pieces.filter((p) => p.kind === "symbol");
  if (syms.length === 0) return undefined;
  const html = syms
    .map((p) => {
      if (p.kind !== "symbol") return "";
      const ch = symbolGlyph(p.font, p.char);
      if (!ch) return "";
      const font = p.font.replace(/['"\\\n\r;]/g, "");
      return `<span style="font-family:'${escapeHtml(font)}'">${escapeHtml(ch)}</span>`;
    })
    .join("");
  return `<span class="wk-sym" contenteditable="false" style="${escapeHtml(css)}">${html}</span>`;
}
