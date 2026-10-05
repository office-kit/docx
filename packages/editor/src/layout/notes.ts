/**
 * Footnote and endnote numbering (ECMA-376 Part 1 §17.11.11 – §17.11.19):
 * references are numbered in document order in the `w:numFmt` of the
 * section's note properties, starting at `w:numStart` and restarting per
 * `w:numRestart` — never, at each section, or on each page.
 *
 * Word writes the document-wide note properties of settings.xml into every
 * section as well, so the section's are the ones read here.
 */

import type { Docx, WmlBlock, WmlParagraph } from "@office-kit/docx";
import { type NoteKind, noteReferences } from "../render-notes.js";
import { formatNumber } from "./number-format.js";
import type { NoteProperties, SectionModel } from "./sections.js";

export interface NoteOccurrence {
  /** `footnote:3`, as in `data-wk-note`. */
  readonly key: string;
  readonly kind: NoteKind;
  readonly id: number;
  /** The body block holding the reference. */
  readonly block: number;
}

function paragraphsOf(block: WmlBlock): WmlParagraph[] {
  if (block.kind === "paragraph") return [block];
  if (block.kind === "table")
    return block.rows.flatMap((r) => r.cells.flatMap((c) => c.paragraphs));
  return [];
}

/** Every note reference in the body, in document order. */
export function collectNoteReferences(doc: Docx): NoteOccurrence[] {
  const out: NoteOccurrence[] = [];
  doc.document.body.blocks.forEach((block, index) => {
    for (const para of paragraphsOf(block)) {
      for (const child of para.children) {
        for (const ref of noteReferences(child)) {
          if (ref.customMark) continue;
          out.push({ key: `${ref.kind}:${ref.id}`, kind: ref.kind, id: ref.id, block: index });
        }
      }
    }
  });
  return out;
}

// The spec's defaults (§17.11.11, §17.11.17): footnotes 1, 2, 3 …; endnotes i, ii, iii …
const DEFAULT_FORMAT: Readonly<Record<NoteKind, string>> = {
  footnote: "decimal",
  endnote: "lowerRoman",
};

function effective(kind: NoteKind, section: SectionModel | undefined): NoteProperties {
  const own = kind === "footnote" ? section?.footnotePr : section?.endnotePr;
  return {
    numFmt: own?.numFmt ?? DEFAULT_FORMAT[kind],
    numStart: own?.numStart ?? 1,
    numRestart: own?.numRestart ?? "continuous",
    pos: own?.pos,
  };
}

/**
 * The mark text of every reference. `sectionOf` and `pageOf` map a body
 * block to its section and page; the page only matters for `eachPage`.
 */
export function numberNotes(
  refs: readonly NoteOccurrence[],
  sections: readonly SectionModel[],
  sectionOf: (block: number) => number,
  pageOf: (block: number) => number,
): Map<string, string> {
  const out = new Map<string, string>();
  const state: Record<NoteKind, { count: number; section: number; page: number }> = {
    footnote: { count: 0, section: -1, page: -1 },
    endnote: { count: 0, section: -1, page: -1 },
  };
  for (const ref of refs) {
    const sectionIndex = sectionOf(ref.block);
    const props = effective(ref.kind, sections[sectionIndex]);
    const s = state[ref.kind];
    const page = props.numRestart === "eachPage" ? pageOf(ref.block) : -1;
    const restart =
      s.section === -1 ||
      (props.numRestart === "eachSect" && sectionIndex !== s.section) ||
      (props.numRestart === "eachPage" && page !== s.page);
    s.count = restart ? (props.numStart ?? 1) : s.count + 1;
    s.section = sectionIndex;
    s.page = page;
    out.set(ref.key, formatNumber(s.count, props.numFmt));
  }
  return out;
}

/** Where endnotes go: after their section (`sectEnd`) or at the document end. */
export function endnotesAtSectionEnd(section: SectionModel | undefined): boolean {
  return effective("endnote", section).pos === "sectEnd";
}
