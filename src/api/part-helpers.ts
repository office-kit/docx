/**
 * Package-part plumbing shared by the References and Mailings APIs: settings
 * edits, built-in styles, note parts, and inserting content at a character
 * offset. Not re-exported from the package entry point.
 */

import type { XmlElement } from "../internal/xml/index.js";
import {
  addPart,
  addRelationship,
  partRelationships,
  relationshipsByType,
} from "../internal/opc/index.js";
import {
  SEED_ENDNOTES_XML,
  SEED_FOOTNOTES_XML,
  WML_CONTENT_TYPES,
  WML_RELATIONSHIPS,
  type WmlFootnotesPart,
  type WmlInline,
  type WmlParagraph,
} from "../internal/wordprocessingml/index.js";
import { childElement, elementFromXml, wAttr } from "../internal/wordprocessingml/field-flow.js";
import {
  addStyle,
  type Docx,
  endnotesPart,
  footnotesPart,
  isolateParagraphRunRange,
  runTextLength,
  stylesPart,
} from "./docx.js";

export { editSettings, readSettings } from "./settings-part.js";

const HEADING_STYLE = /^Heading[1-9]$/;

/**
 * Add Word's built-in styles a feature relies on (TOC 1, Caption, Index 1 …)
 * unless the document already defines them. `styles` is keyed by style id;
 * each value is the full `<w:style>` XML.
 */
export function ensureBuiltInStyles(doc: Docx, styles: Readonly<Record<string, string>>): void {
  const missing = Object.keys(styles).filter((id) => !hasStyle(doc, id));
  if (missing.length === 0) return;
  // addStyle creates styles.xml when the document has none; the placeholder
  // it adds is replaced by the full definition right after.
  const first = missing[0] as string;
  const created = !stylesPart(doc);
  if (created) addStyle(doc, { type: "paragraph", styleId: first });
  const part = stylesPart(doc);
  if (!part) return;
  if (created) part.styles = part.styles.filter((s) => wAttr(s, "styleId") !== first);
  for (const id of missing) part.styles.push(elementFromXml(styles[id] as string));
  doc.stylesDirty = true;
  doc.dirty = true;
}

export function hasStyle(doc: Docx, styleId: string): boolean {
  return !!stylesPart(doc)?.styles.some((s) => wAttr(s, "styleId") === styleId);
}

/** The `<w:style>` element with this id. */
export function findStyleElement(doc: Docx, styleId: string): XmlElement | undefined {
  return stylesPart(doc)?.styles.find((s) => wAttr(s, "styleId") === styleId);
}

/**
 * The outline level (1-9) a paragraph style carries: its own or an ancestor's
 * `w:outlineLvl`, or the built-in heading styles' level. `undefined` for body
 * text.
 */
export function styleOutlineLevel(doc: Docx, styleId: string | undefined): number | undefined {
  const seen = new Set<string>();
  let id = styleId;
  while (id && !seen.has(id)) {
    seen.add(id);
    const style = findStyleElement(doc, id);
    const level = wAttr(childElement(childElement(style, "pPr"), "outlineLvl"), "val");
    if (level !== undefined) {
      const n = Number(level);
      return n >= 0 && n < 9 ? n + 1 : undefined;
    }
    if (HEADING_STYLE.test(id)) return Number(id.slice("Heading".length));
    const name = wAttr(childElement(style, "name"), "val")?.toLowerCase();
    const m = name ? /^heading ([1-9])$/.exec(name) : null;
    if (m) return Number(m[1]);
    id = wAttr(childElement(style, "basedOn"), "val");
  }
  return undefined;
}

/** A paragraph's style id. */
export function paragraphStyleId(paragraph: WmlParagraph): string | undefined {
  return wAttr(childElement(paragraph.pPr, "pStyle"), "val");
}

const DEFAULT_TEXT_WIDTH_TWIPS = 9360;

/**
 * The body's final `<w:sectPr>`, created with Word's US Letter defaults when
 * the document has none (section-level settings need a section to live in).
 */
export function ensureBodySectPr(doc: Docx): XmlElement {
  doc.document.body.sectPr ??= elementFromXml(
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>',
  );
  doc.dirty = true;
  return doc.document.body.sectPr;
}

/** The body text width of the last section (page width minus side margins), in twips. */
export function textWidthTwips(doc: Docx): number {
  const sectPr = doc.document.body.sectPr;
  const width = Number(wAttr(childElement(sectPr, "pgSz"), "w"));
  const mar = childElement(sectPr, "pgMar");
  const left = Number(wAttr(mar, "left") ?? 0);
  const right = Number(wAttr(mar, "right") ?? 0);
  if (!Number.isFinite(width) || width <= 0) return DEFAULT_TEXT_WIDTH_TWIPS;
  return Math.max(width - left - right, 0) || DEFAULT_TEXT_WIDTH_TWIPS;
}

export type NoteKind = "footnote" | "endnote";

const NOTE_PARTS = {
  footnote: {
    name: "/word/footnotes.xml",
    target: "footnotes.xml",
    seed: SEED_FOOTNOTES_XML,
    contentType: WML_CONTENT_TYPES.footnotes,
    rel: WML_RELATIONSHIPS.footnotes,
  },
  endnote: {
    name: "/word/endnotes.xml",
    target: "endnotes.xml",
    seed: SEED_ENDNOTES_XML,
    contentType: WML_CONTENT_TYPES.endnotes,
    rel: WML_RELATIONSHIPS.endnotes,
  },
} as const;

/** The footnotes or endnotes part, created with its separators when missing. */
export function ensureNotesPart(doc: Docx, kind: NoteKind): WmlFootnotesPart {
  const read = (): WmlFootnotesPart | undefined =>
    kind === "footnote" ? footnotesPart(doc) : endnotesPart(doc);
  const existing = read();
  if (existing) return existing;
  const spec = NOTE_PARTS[kind];
  addPart(doc.opc, {
    name: spec.name,
    contentType: spec.contentType,
    data: new TextEncoder().encode(spec.seed),
  });
  const rels = partRelationships(doc.opc, doc.partName);
  if (relationshipsByType(rels, spec.rel).length === 0) {
    addRelationship(rels, { type: spec.rel, target: spec.target });
  }
  const part = read();
  if (!part) throw new Error(`Failed to create the ${kind}s part.`);
  markNotesDirty(doc, kind);
  return part;
}

export function notesPartOf(doc: Docx, kind: NoteKind): WmlFootnotesPart | undefined {
  return kind === "footnote" ? footnotesPart(doc) : endnotesPart(doc);
}

export function markNotesDirty(doc: Docx, kind: NoteKind): void {
  if (kind === "footnote") doc.footnotesDirty = true;
  else doc.endnotesDirty = true;
  doc.dirty = true;
}

/**
 * Insert inlines at a character offset of a paragraph (the offset counts the
 * text of its runs, as the editor's caret does), splitting the run there.
 * Returns the index of the first inserted inline.
 */
export function insertInlinesAt(
  paragraph: WmlParagraph,
  offset: number,
  inlines: readonly WmlInline[],
): number {
  let total = 0;
  for (const child of paragraph.children) if (child.kind === "run") total += runTextLength(child);
  const at = Math.max(0, Math.min(offset, total));
  if (at > 0 && at < total) isolateParagraphRunRange(paragraph, at, total);
  let index = paragraph.children.length;
  if (at === 0) {
    index = 0;
  } else {
    let cursor = 0;
    for (let i = 0; i < paragraph.children.length; i++) {
      const child = paragraph.children[i] as WmlInline;
      if (child.kind === "run") cursor += runTextLength(child);
      if (cursor >= at) {
        index = i + 1;
        break;
      }
    }
  }
  paragraph.children.splice(index, 0, ...inlines);
  return index;
}
