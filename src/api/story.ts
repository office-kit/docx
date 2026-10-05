/**
 * Stories: the text flows of a document other than the main body — headers,
 * footers, footnotes, endnotes and comments (ECMA-376 Part 1 §17.10, §17.11,
 * §17.13.4). Each one holds the same block-level content as the body, so it
 * is exposed as the same typed {@link WmlBody}, and {@link storyView} lets the
 * body-oriented functions of this library (`splitParagraphAt`,
 * `insertParagraphAt`, `setParagraphAlignment` …) edit a story unchanged.
 *
 * Also here: the section-level wiring of headers and footers — which part a
 * section shows (with the inheritance of §17.10.5), creating one, and Word's
 * "Link to Previous".
 */

import {
  addPart,
  addRelationship,
  getPart,
  hasPart,
  partRelationships,
  relationshipById,
  relativizeTarget,
  removePart,
  removeRelationship,
  resolveInternalTarget,
} from "../internal/opc/index.js";
import {
  addSectPrFooterRef,
  addSectPrHeaderRef,
  buildFooterXml,
  buildHeaderXml,
  type HeaderFooterType,
  makePropsElement,
  parseBody,
  WML_CONTENT_TYPES,
  WML_NS,
  WML_RELATIONSHIPS,
  type WmlBlock,
  type WmlBody,
  type WmlDocument,
  writeBodyChildren,
} from "../internal/wordprocessingml/index.js";
import { serializeXml, type XmlElement, type XmlNode } from "../internal/xml/index.js";
import { commentsPart, type Docx, endnotesPart, footnotesPart, getRawPartRoot } from "./docx.js";
import {
  hasStory,
  registerStory,
  registerStoryView,
  runStoryFlushes,
  storyViewOwner,
} from "./story-registry.js";

/** Addresses one story of a document. */
export type StoryRef =
  | { readonly kind: "header"; readonly partName: string }
  | { readonly kind: "footer"; readonly partName: string }
  | { readonly kind: "footnote"; readonly id: number }
  | { readonly kind: "endnote"; readonly id: number }
  | { readonly kind: "comment"; readonly id: number };

/** A stable string for a story, e.g. `header:/word/header1.xml` or `footnote:3`. */
export function storyKey(ref: StoryRef): string {
  return ref.kind === "header" || ref.kind === "footer"
    ? `${ref.kind}:${ref.partName}`
    : `${ref.kind}:${ref.id}`;
}

/** Inverse of {@link storyKey}; `undefined` for a malformed key. */
export function parseStoryKey(key: string): StoryRef | undefined {
  const colon = key.indexOf(":");
  if (colon < 0) return undefined;
  const kind = key.slice(0, colon);
  const rest = key.slice(colon + 1);
  if (!rest) return undefined;
  switch (kind) {
    case "header":
    case "footer":
      return { kind, partName: rest };
    case "footnote":
    case "endnote":
    case "comment": {
      const id = Number(rest);
      return Number.isInteger(id) ? { kind, id } : undefined;
    }
    default:
      return undefined;
  }
}

const FOOTNOTES_PART = "/word/footnotes.xml";
const ENDNOTES_PART = "/word/endnotes.xml";
const COMMENTS_PART = "/word/comments.xml";

const bodies = new WeakMap<Docx, Map<string, WmlBody>>();
const views = new WeakMap<Docx, Map<string, Docx>>();

function idOf(el: XmlElement): number | undefined {
  const value = el.attrs.find((a) => a.name.uri === WML_NS && a.name.local === "id")?.value;
  return value === undefined ? undefined : Number(value);
}

/** Replace an element's children in place (the XML AST's arrays are mutable by contract). */
function setChildren(el: XmlElement, children: XmlNode[]): void {
  (el.children as XmlNode[]).splice(0, el.children.length, ...children);
  (el as { selfClosing: boolean }).selfClosing = children.length === 0;
}

function parseHeaderFooter(
  doc: Docx,
  ref: Extract<StoryRef, { partName: string }>,
): WmlBody | undefined {
  const root = getRawPartRoot(doc, ref.partName);
  const rootLocal = ref.kind === "header" ? "hdr" : "ftr";
  if (!root || root.name.uri !== WML_NS || root.name.local !== rootLocal) return undefined;
  const body = parseBody(root);
  registerStory(doc, storyKey(ref), ref.partName, () => {
    setChildren(root, writeBodyChildren(body));
    // Not markRawPartDirty: that would discard this very story.
    doc.rawPartsDirty.add(ref.partName);
  });
  return body;
}

function parseNote(doc: Docx, ref: Extract<StoryRef, { id: number }>): WmlBody | undefined {
  const elements =
    ref.kind === "footnote"
      ? footnotesPart(doc)?.footnotes
      : ref.kind === "endnote"
        ? endnotesPart(doc)?.footnotes
        : commentsPart(doc)?.comments;
  const element = elements?.find((el) => idOf(el) === ref.id);
  if (!element) return undefined;
  const body = parseBody(element);
  const partName =
    ref.kind === "footnote"
      ? FOOTNOTES_PART
      : ref.kind === "endnote"
        ? ENDNOTES_PART
        : COMMENTS_PART;
  registerStory(doc, storyKey(ref), partName, () => {
    setChildren(element, writeBodyChildren(body));
    if (ref.kind === "footnote") doc.footnotesDirty = true;
    else if (ref.kind === "endnote") doc.endnotesDirty = true;
    else doc.commentsDirty = true;
  });
  return body;
}

/**
 * The typed content of a story, parsed on first use and written back to its
 * part on every save. Edits to the returned blocks are saved like body edits.
 * A raw edit of the part (`markRawPartDirty`) discards it, and the next call
 * parses the edited XML. Returns `undefined` when the story does not exist.
 */
export function storyBody(doc: Docx, ref: StoryRef): WmlBody | undefined {
  const owner = storyViewOwner(doc);
  const key = storyKey(ref);
  let map = bodies.get(owner);
  const cached = map?.get(key);
  if (cached && hasStory(owner, key)) return cached;
  const body =
    ref.kind === "header" || ref.kind === "footer"
      ? parseHeaderFooter(owner, ref)
      : parseNote(owner, ref);
  if (!body) return undefined;
  if (!map) {
    map = new Map();
    bodies.set(owner, map);
  }
  map.set(key, body);
  return body;
}

/**
 * A {@link Docx} that shows the story as its body: every body function of this
 * library edits the story through it, and everything else (styles, numbering,
 * section properties, saving) still reaches the real document. The final
 * `<w:sectPr>` is the real document's, so page setup made through a view
 * still lands on the document's last section.
 */
export function storyView(doc: Docx, ref: StoryRef): Docx | undefined {
  const owner = storyViewOwner(doc);
  const key = storyKey(ref);
  const cached = views.get(owner)?.get(key);
  if (cached) return cached;
  if (!storyBody(owner, ref)) return undefined;
  // Resolved on every access: a raw edit of the part re-parses the story.
  const body = (): WmlBody => storyBody(owner, ref) ?? { blocks: [], extras: [] };
  const viewBody: WmlBody = {
    get blocks() {
      return body().blocks;
    },
    set blocks(value) {
      body().blocks = value;
    },
    get extras() {
      return body().extras;
    },
    set extras(value) {
      body().extras = value;
    },
  };
  // Defined separately: an accessor cannot express the optional property.
  Object.defineProperty(viewBody, "sectPr", {
    enumerable: true,
    get: () => owner.document.body.sectPr,
    set: (value: XmlElement | undefined) => {
      if (value) owner.document.body.sectPr = value;
      else delete owner.document.body.sectPr;
    },
  });
  const document: WmlDocument = { rootAttrs: [], body: viewBody, extras: [] };
  const view = new Proxy(owner, {
    get: (target, prop) => (prop === "document" ? document : Reflect.get(target, prop)),
  });
  registerStoryView(view, owner);
  let map = views.get(owner);
  if (!map) {
    map = new Map();
    views.set(owner, map);
  }
  map.set(key, view);
  return view;
}

// --- sections and their headers / footers -----------------------------------

export type HeaderFooterKind = "header" | "footer";

/**
 * The `<w:sectPr>` of every section in document order: one per
 * section-break paragraph, then the body's final one. The final section's
 * entry is `undefined` when the body has no trailing `<w:sectPr>` (every
 * property at its default).
 */
export function sectionProperties(doc: Docx): Array<XmlElement | undefined> {
  const out: Array<XmlElement | undefined> = [];
  for (const block of doc.document.body.blocks) {
    if (block.kind !== "paragraph" || !block.pPr) continue;
    const sectPr = block.pPr.children.find(
      (c): c is XmlElement =>
        c.kind === "element" && c.name.uri === WML_NS && c.name.local === "sectPr",
    );
    if (sectPr) out.push(sectPr);
  }
  out.push(doc.document.body.sectPr);
  return out;
}

function ensureSectionProperties(doc: Docx, sectionIndex: number): XmlElement {
  const all = sectionProperties(doc);
  if (!Number.isInteger(sectionIndex) || sectionIndex < 0 || sectionIndex >= all.length) {
    throw new Error(`Section ${sectionIndex} does not exist (the document has ${all.length}).`);
  }
  const existing = all[sectionIndex];
  if (existing) return existing;
  const sectPr = makePropsElement("sectPr");
  doc.document.body.sectPr = sectPr;
  return sectPr;
}

function ownReference(
  sectPr: XmlElement | undefined,
  kind: HeaderFooterKind,
  type: HeaderFooterType,
): XmlElement | undefined {
  const local = `${kind}Reference`;
  return sectPr?.children.find(
    (c): c is XmlElement =>
      c.kind === "element" &&
      c.name.uri === WML_NS &&
      c.name.local === local &&
      (c.attrs.find((a) => a.name.local === "type")?.value ?? "default") === type,
  );
}

function referenceRelId(ref: XmlElement): string | undefined {
  return ref.attrs.find((a) => a.name.local === "id" && a.name.uri !== WML_NS)?.value;
}

function partNameOfRelId(doc: Docx, relId: string): string | undefined {
  const rel = relationshipById(partRelationships(doc.opc, doc.partName), relId);
  if (!rel || rel.targetMode === "External") return undefined;
  return resolveInternalTarget(doc.partName, rel.target);
}

/**
 * The part a section shows for one header/footer type, following §17.10.5:
 * a section without its own reference of that type uses the previous
 * section's. `section` is the index of the section that owns the reference.
 */
export function resolveHeaderFooter(
  doc: Docx,
  sectionIndex: number,
  kind: HeaderFooterKind,
  type: HeaderFooterType,
): { partName: string; section: number } | undefined {
  const all = sectionProperties(doc);
  for (let i = Math.min(sectionIndex, all.length - 1); i >= 0; i--) {
    const ref = ownReference(all[i], kind, type);
    const relId = ref && referenceRelId(ref);
    const partName = relId && partNameOfRelId(doc, relId);
    if (partName && hasPart(doc.opc, partName)) return { partName, section: i };
  }
  return undefined;
}

function allocatePartName(doc: Docx, kind: HeaderFooterKind): string {
  let n = 1;
  while (hasPart(doc.opc, `/word/${kind}${n}.xml`)) n++;
  return `/word/${kind}${n}.xml`;
}

function addOwnHeaderFooter(
  doc: Docx,
  sectPr: XmlElement,
  kind: HeaderFooterKind,
  type: HeaderFooterType,
  xml: string,
): string {
  const partName = allocatePartName(doc, kind);
  addPart(doc.opc, {
    name: partName,
    contentType: WML_CONTENT_TYPES[kind],
    data: new TextEncoder().encode(xml),
  });
  const rel = addRelationship(partRelationships(doc.opc, doc.partName), {
    type: WML_RELATIONSHIPS[kind],
    target: relativizeTarget(doc.partName, partName),
  });
  if (kind === "header") addSectPrHeaderRef(sectPr, type, rel.id);
  else addSectPrFooterRef(sectPr, type, rel.id);
  doc.dirty = true;
  return partName;
}

function emptyHeaderFooterXml(kind: HeaderFooterKind): string {
  return kind === "header" ? buildHeaderXml("") : buildFooterXml("");
}

/**
 * The part to edit for a section's header or footer of one type, creating an
 * empty one when no section up to this one has it (what Word does when you
 * double-click an empty header area). A linked header is shared with the
 * section it comes from, so editing it edits that one too, as in Word.
 */
export function ensureHeaderFooter(
  doc: Docx,
  sectionIndex: number,
  kind: HeaderFooterKind,
  type: HeaderFooterType,
): string {
  const owner = storyViewOwner(doc);
  const resolved = resolveHeaderFooter(owner, sectionIndex, kind, type);
  if (resolved) return resolved.partName;
  const sectPr = ensureSectionProperties(owner, sectionIndex);
  return addOwnHeaderFooter(owner, sectPr, kind, type, emptyHeaderFooterXml(kind));
}

/** Whether a section shows the previous section's header/footer of this type. */
export function isHeaderFooterLinked(
  doc: Docx,
  sectionIndex: number,
  kind: HeaderFooterKind,
  type: HeaderFooterType,
): boolean {
  return !ownReference(sectionProperties(doc)[sectionIndex], kind, type);
}

/**
 * Word's "Link to Previous". Linking removes the section's own reference (and
 * the part, when nothing else references it), so the previous section's
 * header shows. Unlinking gives the section its own copy of what it showed.
 * The first section has no previous one, so it cannot be linked.
 */
export function setHeaderFooterLinked(
  doc: Docx,
  sectionIndex: number,
  kind: HeaderFooterKind,
  type: HeaderFooterType,
  linked: boolean,
): void {
  const owner = storyViewOwner(doc);
  if (linked && sectionIndex === 0) {
    throw new Error("The first section has no previous section to link to.");
  }
  const sectPr = ensureSectionProperties(owner, sectionIndex);
  const own = ownReference(sectPr, kind, type);
  if (linked) {
    if (!own) return;
    const relId = referenceRelId(own);
    (sectPr.children as XmlNode[]).splice(sectPr.children.indexOf(own), 1);
    if (relId) removeUnreferencedPart(owner, relId);
    owner.dirty = true;
    return;
  }
  if (own) return;
  const inherited = resolveHeaderFooter(owner, sectionIndex, kind, type);
  const xml = inherited ? currentPartXml(owner, inherited.partName) : emptyHeaderFooterXml(kind);
  addOwnHeaderFooter(owner, sectPr, kind, type, xml);
}

/** A part's XML as it would be saved now, pending story edits included. */
function currentPartXml(doc: Docx, partName: string): string {
  getRawPartRoot(doc, partName);
  runStoryFlushes(doc);
  const xmlDoc = doc.rawParts.get(partName);
  if (xmlDoc) return serializeXml(xmlDoc);
  const part = getPart(doc.opc, partName);
  return part ? new TextDecoder("utf-8").decode(part.data) : "";
}

function removeUnreferencedPart(doc: Docx, relId: string): void {
  for (const sectPr of sectionProperties(doc)) {
    for (const c of sectPr?.children ?? []) {
      if (c.kind === "element" && referenceRelId(c) === relId) return;
    }
  }
  const partName = partNameOfRelId(doc, relId);
  removeRelationship(partRelationships(doc.opc, doc.partName), relId);
  if (partName) {
    removePart(doc.opc, partName);
    doc.rawParts.delete(partName);
    doc.rawPartsDirty.delete(partName);
  }
}

/**
 * The blocks inside a block-level content control (`<w:sdt>`, §17.5.2) —
 * the wrapper Word puts around cover pages, watermarks and tables of
 * contents — or `undefined` when `block` is not one. The blocks are a
 * read-only view: they are parsed from the control's XML on each call, and
 * changes to them are not saved.
 */
export function contentControlBlocks(block: WmlBlock): readonly WmlBlock[] | undefined {
  if (block.kind !== "raw") return undefined;
  const { node } = block;
  if (node.name.uri !== WML_NS || node.name.local !== "sdt") return undefined;
  const content = node.children.find(
    (c) => c.kind === "element" && c.name.uri === WML_NS && c.name.local === "sdtContent",
  );
  // sdtContent at block level holds the same children as w:body (§17.5.2.38).
  return content?.kind === "element" ? parseBody(content).blocks : [];
}
