/**
 * Footnote and endnote marks on the canvas. A reference
 * (`<w:footnoteReference w:id>`, §17.11.14) and the mark at the start of the
 * note's own text (`<w:footnoteRef/>`, §17.11.13) show a number that depends
 * on numbering settings and, for per-page restarts, on pagination. The
 * renderer therefore emits an empty placeholder after the run, and the page
 * layout fills in the numbers (see `layout/notes.ts`).
 */

import type { WmlInline, WmlRun, XmlElement } from "@office-kit/docx";
import { WML_NS } from "./wml-ns.js";

export type NoteKind = "footnote" | "endnote";

export interface NoteReference {
  readonly kind: NoteKind;
  readonly id: number;
  /** `w:customMarkFollows`: the run after the reference holds the mark, no number. */
  readonly customMark: boolean;
}

const ON_VALUES: ReadonlySet<string> = new Set(["1", "true", "on"]);

function referenceOf(el: XmlElement): NoteReference | undefined {
  if (el.name.uri !== WML_NS) return undefined;
  const local = el.name.local;
  if (local !== "footnoteReference" && local !== "endnoteReference") return undefined;
  const id = Number(el.attrs.find((a) => a.name.local === "id")?.value);
  if (!Number.isInteger(id)) return undefined;
  const custom = el.attrs.find((a) => a.name.local === "customMarkFollows")?.value;
  return {
    kind: local === "footnoteReference" ? "footnote" : "endnote",
    id,
    customMark: custom !== undefined && ON_VALUES.has(custom),
  };
}

function referencesUnder(el: XmlElement, out: NoteReference[]): void {
  const own = referenceOf(el);
  if (own) {
    out.push(own);
    return;
  }
  for (const c of el.children) if (c.kind === "element") referencesUnder(c, out);
}

/**
 * The note references an inline holds: a run's own (normally zero or one),
 * or any inside an unmodelled inline — a reference run built in memory, or
 * one inside a hyperlink or content control.
 */
export function noteReferences(inline: WmlInline): NoteReference[] {
  const out: NoteReference[] = [];
  if (inline.kind === "raw") {
    referencesUnder(inline.node, out);
    return out;
  }
  for (const piece of inline.pieces) {
    if (piece.kind !== "raw") continue;
    const ref = referenceOf(piece.node);
    if (ref) out.push(ref);
  }
  return out;
}

/** Whether the run holds a note's own mark (`<w:footnoteRef/>` / `<w:endnoteRef/>`). */
export function holdsOwnNoteMark(run: WmlRun): boolean {
  return run.pieces.some(
    (p) =>
      p.kind === "raw" &&
      p.node.name.uri === WML_NS &&
      (p.node.name.local === "footnoteRef" || p.node.name.local === "endnoteRef"),
  );
}

/**
 * Placeholders for the references an inline holds, rendered after it. They
 * sit outside the run span so the number never counts as the run's text when
 * DOM offsets are mapped back.
 */
export function noteMarksHtml(inline: WmlInline): string {
  return noteReferences(inline)
    .filter((ref) => !ref.customMark)
    .map(
      (ref) =>
        `<sup class="wk-noteref" data-wk-note="${ref.kind}:${ref.id}" contenteditable="false"></sup>`,
    )
    .join("");
}

/**
 * Placeholder for a note's own mark, rendered before its run (the mark leads
 * the note text). It takes its number from the story it is rendered in.
 */
export function ownNoteMarkHtml(run: WmlRun): string {
  return holdsOwnNoteMark(run)
    ? `<sup class="wk-noteref wk-noteref-own" contenteditable="false"></sup>`
    : "";
}
