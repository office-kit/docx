/**
 * Field results on the canvas. A complex field is a run sequence
 * `fldChar begin · instrText… · fldChar separate · result runs · fldChar end`
 * (ECMA-376 Part 1 §17.16.18); a simple field is one `<w:fldSimple w:instr>`.
 * The renderer tags every result run with the field's type, so the page
 * layout can fill in PAGE / NUMPAGES … per page and the UI can style fields
 * (Mailings ▸ Highlight Merge Fields).
 */

import type { WmlParagraph, XmlElement } from "@office-kit/docx";
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
