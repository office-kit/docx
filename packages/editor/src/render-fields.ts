/**
 * Canvas rendering for fields and symbols.
 *
 * A complex field is a run sequence begin / code / separate / result / end
 * (§17.16.18). The canvas shows the result runs as ordinary editable text
 * (marked `wk-fresult`) and renders the field characters and code as hidden,
 * non-editable spans. Toggle Field Codes is then pure CSS: the editor root gets
 * a class that hides results and shows `{ CODE }` instead, so switching views
 * never re-renders or moves the caret.
 */

import type { WmlParagraph, WmlRun } from "@office-kit/docx";

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
