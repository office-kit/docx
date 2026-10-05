import type { XmlElement } from "../xml/index.js";
import { WML_NS } from "./namespaces.js";
import { buildPgSz, PAGE_SIZE_LETTER } from "./section-builders.js";
import type { WmlDocument, WmlParagraph } from "./types.js";

/**
 * Which sections a section-level setter changes: `"all"`, one section index,
 * or several. Sections are numbered from 0 in document order; the last one is
 * the body's trailing `<w:sectPr>`.
 */
export type SectionScope = "all" | number | readonly number[];

/** The `<w:sectPr>` inside a paragraph's `<w:pPr>` (a section break), if any. */
export function paragraphSectPr(p: WmlParagraph): XmlElement | undefined {
  return p.pPr?.children.find(
    (c): c is XmlElement =>
      c.kind === "element" && c.name.uri === WML_NS && c.name.local === "sectPr",
  );
}

/**
 * The body's trailing `<w:sectPr>`, created when missing. A new one carries a
 * Letter `<w:pgSz>` like Word's default document, since a section without
 * one has an application-defined page size.
 */
export function ensureBodySectionProperties(document: WmlDocument): XmlElement {
  if (document.body.sectPr) return document.body.sectPr;
  const sectPr: XmlElement = {
    kind: "element",
    name: { uri: WML_NS, local: "sectPr", prefix: "w" },
    attrs: [],
    children: [buildPgSz(PAGE_SIZE_LETTER)],
    xmlSpace: "default",
    selfClosing: false,
  };
  document.body.sectPr = sectPr;
  return sectPr;
}

/**
 * Every section's `<w:sectPr>` in document order: one per section-break
 * paragraph, then the body's own (ECMA-376 §17.6.17: a paragraph's sectPr
 * ends the section that contains it).
 */
export function sectionPropertiesList(document: WmlDocument): XmlElement[] {
  const out: XmlElement[] = [];
  for (const block of document.body.blocks) {
    if (block.kind !== "paragraph") continue;
    const sectPr = paragraphSectPr(block);
    if (sectPr) out.push(sectPr);
  }
  out.push(ensureBodySectionProperties(document));
  return out;
}

/** The section index (0-based) holding body block `blockIndex`. */
export function sectionIndexOfBlock(document: WmlDocument, blockIndex: number): number {
  let section = 0;
  const blocks = document.body.blocks;
  for (let i = 0; i < blocks.length && i < blockIndex; i++) {
    const block = blocks[i];
    if (block?.kind === "paragraph" && paragraphSectPr(block)) section++;
  }
  return section;
}

/** The `<w:sectPr>` elements a {@link SectionScope} names; the default is the last section. */
export function resolveSectionScope(
  document: WmlDocument,
  scope: SectionScope | undefined,
): XmlElement[] {
  const all = sectionPropertiesList(document);
  if (scope === "all") return all;
  const indices = scope === undefined ? [all.length - 1] : typeof scope === "number" ? [scope] : scope;
  return indices.map((i) => {
    const sectPr = all[i];
    if (!sectPr) throw new RangeError(`Section ${i} does not exist (the document has ${all.length}).`);
    return sectPr;
  });
}
