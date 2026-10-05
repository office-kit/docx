/**
 * Document-level Insert-tab operations: cover pages, header and footer
 * content, page numbers and their format, and inserting the body of another
 * document (Insert ▸ Object ▸ Text from File).
 */

import {
  addPart,
  addRelationship,
  getPart,
  hasPart,
  partRelationships,
  relationshipsByType,
  removePart,
  removeRelationship,
} from "../internal/opc/index.js";
import {
  EMPTY_NUMBERING_XML,
  findStyle,
  paragraphToElement,
  parseParagraph,
  parseWmlDocument,
  SEED_ENDNOTES_XML,
  SEED_FOOTNOTES_XML,
  type HeaderFooterType,
  WML_CONTENT_TYPES,
  WML_NS,
  WML_RELATIONSHIPS,
  type WmlBlock,
  type WmlParagraph,
  type WmlRun,
  writeWmlDocument,
} from "../internal/wordprocessingml/index.js";
import { type XmlAttr, type XmlElement, type XmlNode } from "../internal/xml/index.js";
import {
  addFooter,
  addHeader,
  addStyle,
  type Docx,
  endnotesPart,
  footers,
  footnotesPart,
  getRawPartRoot,
  headers,
  markRawPartDirty,
  numberingPart,
  setDocumentSettingOnOff,
  stylesPart,
} from "./docx.js";
import { buildComplexField } from "./insert-fields.js";
import {
  isW,
  nsDecl,
  R_NS,
  SECTPR_ORDER,
  setOrderedChild,
  setParagraphChild,
  textRun,
  wAttrValue,
  wChild,
  wEl,
} from "./insert-xml.js";

// --- section helpers --------------------------------------------------------------

function bodySectPr(doc: Docx): XmlElement {
  if (!doc.document.body.sectPr) {
    doc.document.body.sectPr = { ...wEl("sectPr"), selfClosing: false };
  }
  return doc.document.body.sectPr;
}

function setSectPrChild(doc: Docx, local: string, child: XmlElement | undefined): void {
  doc.document.body.sectPr = setOrderedChild(bodySectPr(doc), local, child, SECTPR_ORDER);
  doc.dirty = true;
}

/** Turn "Different First Page" (`w:titlePg`) on or off for the last section. */
export function setDifferentFirstPage(doc: Docx, on: boolean): void {
  setSectPrChild(doc, "titlePg", on ? wEl("titlePg") : undefined);
}

// --- cover pages ------------------------------------------------------------------

// Hidden bookmark marking the cover page's extent, so it can be found and
// replaced later. Word marks its cover pages with a docPartObj content control
// ("Cover Pages" gallery) instead; removeCoverPage recognizes both.
const COVER_PAGE_BOOKMARK = "_CoverPage";
const COVER_PAGE_GALLERY = "Cover Pages";

function bookmarkSpan(doc: Docx, name: string): { start: number; end: number } | undefined {
  const blocks = doc.document.body.blocks;
  let start = -1;
  let id: string | undefined;
  for (const [i, b] of blocks.entries()) {
    if (b.kind !== "paragraph") continue;
    for (const c of b.children) {
      if (c.kind !== "raw") continue;
      if (start < 0 && isW(c.node, "bookmarkStart") && wAttrValue(c.node, "name") === name) {
        start = i;
        id = wAttrValue(c.node, "id");
      }
      if (start >= 0 && isW(c.node, "bookmarkEnd") && wAttrValue(c.node, "id") === id) {
        return { start, end: i };
      }
    }
  }
  return undefined;
}

function isCoverPageSdt(block: WmlBlock | undefined): boolean {
  if (block?.kind !== "raw" || !isW(block.node, "sdt")) return false;
  const gallery = wChild(wChild(wChild(block.node, "sdtPr"), "docPartObj"), "docPartGallery");
  return gallery !== undefined && wAttrValue(gallery, "val") === COVER_PAGE_GALLERY;
}

/** Whether the document starts with a cover page (ours or Word's). */
export function hasCoverPage(doc: Docx): boolean {
  return !!bookmarkSpan(doc, COVER_PAGE_BOOKMARK) || isCoverPageSdt(doc.document.body.blocks[0]);
}

/** Remove the document's cover page. Returns whether there was one. */
export function removeCoverPage(doc: Docx): boolean {
  const blocks = doc.document.body.blocks;
  const span = bookmarkSpan(doc, COVER_PAGE_BOOKMARK);
  if (span) {
    blocks.splice(span.start, span.end - span.start + 1);
  } else if (isCoverPageSdt(blocks[0])) {
    blocks.splice(0, 1);
  } else {
    return false;
  }
  doc.dirty = true;
  return true;
}

function freeBookmarkId(doc: Docx): string {
  let max = -1;
  const scan = (node: XmlElement): void => {
    if (isW(node, "bookmarkStart") || isW(node, "bookmarkEnd")) {
      const n = Number(wAttrValue(node, "id"));
      if (Number.isFinite(n) && n > max) max = n;
    }
    for (const c of node.children) if (c.kind === "element") scan(c);
  };
  for (const b of doc.document.body.blocks) {
    if (b.kind === "paragraph") for (const c of b.children) if (c.kind === "raw") scan(c.node);
    if (b.kind === "raw") scan(b.node);
  }
  return String(max + 1);
}

/**
 * Put a cover page at the start of the document, replacing any existing one:
 * `paragraphs` followed by a page break, with "Different First Page" turned on
 * so the cover shows no header or footer (as Word does).
 */
export function insertCoverPage(doc: Docx, paragraphs: readonly WmlParagraph[]): void {
  if (paragraphs.length === 0) throw new Error("A cover page needs content.");
  removeCoverPage(doc);
  const id = freeBookmarkId(doc);
  const content = paragraphs.map((p) => structuredClone(p));
  const breakPara: WmlParagraph = {
    kind: "paragraph",
    children: [{ kind: "run", pieces: [{ kind: "break", breakType: "page" }], extras: [] }],
    extras: [],
  };
  content[0]?.children.unshift({
    kind: "raw",
    node: wEl("bookmarkStart", { id, name: COVER_PAGE_BOOKMARK }),
  });
  breakPara.children.push({ kind: "raw", node: wEl("bookmarkEnd", { id }) });
  doc.document.body.blocks.unshift(...content, breakPara);
  setDifferentFirstPage(doc, true);
  doc.dirty = true;
}

// --- header / footer content -----------------------------------------------------

export type HeaderFooterKind = "header" | "footer";

function referenceLocal(kind: HeaderFooterKind): string {
  return kind === "header" ? "headerReference" : "footerReference";
}

function referenceOf(
  doc: Docx,
  kind: HeaderFooterKind,
  type: HeaderFooterType,
): XmlElement | undefined {
  const sectPr = doc.document.body.sectPr;
  return sectPr?.children.find(
    (c): c is XmlElement =>
      isW(c, referenceLocal(kind)) && (wAttrValue(c, "type") ?? "default") === type,
  );
}

function relIdOf(ref: XmlElement): string | undefined {
  return ref.attrs.find((a) => a.name.uri === R_NS && a.name.local === "id")?.value;
}

function partNameFor(
  doc: Docx,
  kind: HeaderFooterKind,
  type: HeaderFooterType,
): string | undefined {
  const ref = referenceOf(doc, kind, type);
  const relId = ref ? relIdOf(ref) : undefined;
  if (!relId) return undefined;
  const list = kind === "header" ? headers(doc) : footers(doc);
  return list.find((h) => h.relId === relId)?.partName;
}

/** The paragraphs of the last section's header or footer of `type`, if it has one. */
export function headerFooterParagraphs(
  doc: Docx,
  kind: HeaderFooterKind,
  type: HeaderFooterType = "default",
): WmlParagraph[] | undefined {
  const partName = partNameFor(doc, kind, type);
  const root = partName ? getRawPartRoot(doc, partName) : undefined;
  if (!root) return undefined;
  return root.children.filter((c): c is XmlElement => isW(c, "p")).map(parseParagraph);
}

/**
 * Replace the content of the last section's header or footer of `type`
 * (creating the part and its reference when missing). A `first` header turns
 * on "Different First Page"; an `even` one turns on "Different Odd & Even
 * Pages" (`w:evenAndOddHeaders`).
 */
export function setHeaderFooterParagraphs(
  doc: Docx,
  kind: HeaderFooterKind,
  type: HeaderFooterType,
  paragraphs: readonly WmlParagraph[],
): string {
  let partName = partNameFor(doc, kind, type);
  if (!partName) {
    const relId = kind === "header" ? addHeader(doc, "", type) : addFooter(doc, "", type);
    const list = kind === "header" ? headers(doc) : footers(doc);
    partName = list.find((h) => h.relId === relId)?.partName;
  }
  const root = partName ? getRawPartRoot(doc, partName) : undefined;
  if (!partName || !root) throw new Error(`Could not create the ${kind} part.`);
  ensureHeaderFooterStyle(doc, kind);
  const children: XmlNode[] = (paragraphs.length ? paragraphs : [emptyParagraph()]).map((p) =>
    paragraphToElement(p),
  );
  const attrs = [...root.attrs];
  // Content may carry hyperlinks / pictures (r:id); declare `r` on the root.
  if (!attrs.some((a) => a.isNamespaceDecl && a.name.local === "r")) attrs.push(nsDecl("r", R_NS));
  replaceRoot(doc, partName, { ...root, attrs, children, selfClosing: false });
  if (type === "first") setDifferentFirstPage(doc, true);
  if (type === "even") setDocumentSettingOnOff(doc, "evenAndOddHeaders", true);
  doc.dirty = true;
  return partName;
}

/** Swap a raw part's root element in place (the cached XmlDocument is reused). */
function replaceRoot(doc: Docx, partName: string, root: XmlElement): void {
  const cached = doc.rawParts.get(partName);
  if (!cached) throw new Error(`Part ${partName} is not open.`);
  doc.rawParts.set(partName, { ...cached, root });
  markRawPartDirty(doc, partName);
}

function emptyParagraph(): WmlParagraph {
  return { kind: "paragraph", children: [], extras: [] };
}

/**
 * Remove the last section's header or footer of `type`: its reference, and
 * the part itself when nothing else refers to it.
 */
export function removeHeaderFooter(
  doc: Docx,
  kind: HeaderFooterKind,
  type: HeaderFooterType = "default",
): boolean {
  const ref = referenceOf(doc, kind, type);
  if (!ref) return false;
  const relId = relIdOf(ref);
  const partName = partNameFor(doc, kind, type);
  const sectPr = bodySectPr(doc);
  doc.document.body.sectPr = { ...sectPr, children: sectPr.children.filter((c) => c !== ref) };
  if (relId && partName && !relIdStillUsed(doc, relId)) {
    removeRelationship(partRelationships(doc.opc, doc.partName), relId);
    removePart(doc.opc, partName);
    doc.rawParts.delete(partName);
    doc.rawPartsDirty.delete(partName);
  }
  doc.dirty = true;
  return true;
}

function relIdStillUsed(doc: Docx, relId: string): boolean {
  const uses = (sectPr: XmlElement | undefined): boolean =>
    !!sectPr?.children.some(
      (c) => c.kind === "element" && c.attrs.some((a) => a.name.uri === R_NS && a.value === relId),
    );
  if (uses(doc.document.body.sectPr)) return true;
  return doc.document.body.blocks.some(
    (b) => b.kind === "paragraph" && uses(b.pPr ? wChild(b.pPr, "sectPr") : undefined),
  );
}

const DEFAULT_TEXT_WIDTH_TWIPS = 9360;

/** The text width of the last section (page width minus side margins). */
function textWidthTwips(doc: Docx): number {
  const sectPr = doc.document.body.sectPr;
  const pgSz = wChild(sectPr, "pgSz");
  const pgMar = wChild(sectPr, "pgMar");
  const w = Number(pgSz ? wAttrValue(pgSz, "w") : undefined);
  const left = Number(pgMar ? wAttrValue(pgMar, "left") : undefined);
  const right = Number(pgMar ? wAttrValue(pgMar, "right") : undefined);
  const width = w - (left || 0) - (right || 0);
  return Number.isFinite(width) && width > 0 ? width : DEFAULT_TEXT_WIDTH_TWIPS;
}

const HEADER_STYLE_IDS: Readonly<Record<HeaderFooterKind, { id: string; name: string }>> = {
  header: { id: "Header", name: "header" },
  footer: { id: "Footer", name: "footer" },
};
// Word's built-in Header / Footer styles have this UI priority.
const HEADER_STYLE_PRIORITY = 99;

/**
 * Define Word's built-in Header / Footer paragraph style (center and right
 * tab stops across the text width, no spacing after) when the document lacks
 * it, so header presets that reference it look as in Word.
 */
export function ensureHeaderFooterStyle(doc: Docx, kind: HeaderFooterKind): string {
  const { id, name } = HEADER_STYLE_IDS[kind];
  const existing = stylesPart(doc);
  if (existing && findStyle(existing, id)) return id;
  addStyle(doc, {
    type: "paragraph",
    styleId: id,
    name,
    basedOn: "Normal",
    uiPriority: HEADER_STYLE_PRIORITY,
  });
  const part = stylesPart(doc);
  const style = part ? findStyle(part, id) : undefined;
  if (!part || !style) return id;
  const width = textWidthTwips(doc);
  const pPr = wEl("pPr", {}, [
    wEl("tabs", {}, [
      wEl("tab", { val: "center", pos: String(Math.round(width / 2)) }),
      wEl("tab", { val: "right", pos: String(width) }),
    ]),
    wEl("spacing", { after: "0", line: "240", lineRule: "auto" }),
  ]);
  const idx = part.styles.indexOf(style);
  part.styles[idx] = { ...style, children: [...style.children, wEl("unhideWhenUsed"), pPr] };
  doc.stylesDirty = true;
  return id;
}

// --- page numbers ---------------------------------------------------------------------

export interface PageNumberOptions {
  /** Top of Page (header), Bottom of Page (footer), or Page Margins (side of the page). */
  readonly position: "top" | "bottom" | "margin";
  readonly align: "left" | "center" | "right";
  /** `plain`: "1"; `pageXofY`: "Page 1 of 3" (PAGE and NUMPAGES fields). */
  readonly style?: "plain" | "pageXofY";
}

/**
 * Put a page number in the header (top), footer (bottom) or page margin of the
 * last section, replacing that header / footer's content as Word's Page
 * Number gallery does. A margin number is a frame (`w:framePr`) anchored to
 * the page's side, vertically centered.
 */
export function insertPageNumbers(doc: Docx, options: PageNumberOptions): void {
  const kind: HeaderFooterKind = options.position === "bottom" ? "footer" : "header";
  const para: WmlParagraph = { kind: "paragraph", children: [], extras: [] };
  setParagraphChild(para, "pStyle", wEl("pStyle", { val: HEADER_STYLE_IDS[kind].id }));
  if (options.position === "margin") {
    setParagraphChild(
      para,
      "framePr",
      wEl("framePr", {
        wrap: "around",
        vAnchor: "page",
        hAnchor: "page",
        xAlign: options.align === "left" ? "left" : options.align === "right" ? "right" : "center",
        yAlign: "center",
      }),
    );
  } else {
    setParagraphChild(
      para,
      "jc",
      wEl("jc", { val: options.align === "left" ? "left" : options.align }),
    );
  }
  if (options.style === "pageXofY") {
    para.children.push(
      textRun("Page "),
      ...buildComplexField("PAGE", "1"),
      textRun(" of "),
      ...buildComplexField("NUMPAGES", "1"),
    );
  } else {
    para.children.push(...buildComplexField("PAGE", "1"));
  }
  setHeaderFooterParagraphs(doc, kind, "default", [para]);
}

/** Remove PAGE / NUMPAGES fields from every header and footer of the document. */
export function removePageNumbers(doc: Docx): number {
  let removed = 0;
  for (const { partName } of [...headers(doc), ...footers(doc)]) {
    const root = getRawPartRoot(doc, partName);
    if (!root) continue;
    let changed = false;
    const children: XmlNode[] = [];
    for (const c of root.children) {
      if (!isW(c, "p")) {
        children.push(c);
        continue;
      }
      const para = parseParagraph(c);
      const n = stripPageFields(para);
      if (n === 0) {
        children.push(c);
        continue;
      }
      removed += n;
      changed = true;
      // A framed paragraph held only the number: drop it entirely.
      const framed = para.pPr ? !!wChild(para.pPr, "framePr") : false;
      if (!framed) children.push(paragraphToElement(para));
    }
    if (changed) replaceRoot(doc, partName, { ...root, children });
  }
  if (removed) doc.dirty = true;
  return removed;
}

const PAGE_FIELD_TYPES: ReadonlySet<string> = new Set(["PAGE", "NUMPAGES", "SECTIONPAGES"]);

/** Remove page-number fields (and "Page … of" text beside them) from a paragraph. */
function stripPageFields(para: WmlParagraph): number {
  let removed = 0;
  let begin = -1;
  let instr = "";
  for (let i = 0; i < para.children.length; i++) {
    const c = para.children[i];
    if (c?.kind !== "run") continue;
    for (const p of c.pieces) {
      if (p.kind === "fieldChar" && p.charType === "begin" && begin < 0) {
        begin = i;
        instr = "";
      } else if (p.kind === "instrText" && begin >= 0) instr += p.value;
      else if (p.kind === "fieldChar" && p.charType === "end" && begin >= 0) {
        const type = instr.trim().split(/\s+/)[0]?.toUpperCase() ?? "";
        if (PAGE_FIELD_TYPES.has(type)) {
          para.children.splice(begin, i - begin + 1);
          i = begin - 1;
          removed++;
        }
        begin = -1;
        break;
      }
    }
  }
  if (removed) {
    para.children = para.children.filter(
      (c) => !(c.kind === "run" && /^\s*(Page|of)\s*$/i.test(runPlainText(c))),
    );
  }
  return removed;
}

function runPlainText(run: WmlRun): string {
  return run.pieces.map((p) => (p.kind === "text" ? p.value : "")).join("");
}

/** ST_NumberFormat values Word's Format Page Numbers dialog offers (§17.18.59). */
export const PAGE_NUMBER_FORMATS = [
  "decimal",
  "upperRoman",
  "lowerRoman",
  "upperLetter",
  "lowerLetter",
  "numberInDash",
  "decimalFullWidth",
  "japaneseCounting",
  "aiueoFullWidth",
  "iroha",
  "chineseCounting",
  "ideographTraditional",
] as const;
export type PageNumberFormat = (typeof PAGE_NUMBER_FORMATS)[number];

/** ST_ChapterSep (§17.18.6). */
export const CHAPTER_SEPARATORS = ["hyphen", "period", "colon", "emDash", "enDash"] as const;
export type ChapterSeparator = (typeof CHAPTER_SEPARATORS)[number];

const isPageNumberFormat = (v: string): v is PageNumberFormat =>
  PAGE_NUMBER_FORMATS.some((f) => f === v);
const isChapterSeparator = (v: string): v is ChapterSeparator =>
  CHAPTER_SEPARATORS.some((f) => f === v);

export interface PageNumberFormatOptions {
  readonly format?: PageNumberFormat;
  /** Start at this number; undefined continues from the previous section. */
  readonly start?: number;
  /** Include chapter number: the heading level (1–9) whose number prefixes the page number. */
  readonly chapterStyle?: number;
  readonly chapterSeparator?: ChapterSeparator;
}

/** Set the last section's page numbering (`w:pgNumType`, §17.6.12). */
export function setPageNumberFormat(doc: Docx, options: PageNumberFormatOptions): void {
  const attrs: Record<string, string> = {};
  if (options.format && options.format !== "decimal") attrs.fmt = options.format;
  if (options.start !== undefined) {
    if (!Number.isInteger(options.start) || options.start < 0) {
      throw new RangeError(`Start at must be a whole number ≥ 0, got ${options.start}.`);
    }
    attrs.start = String(options.start);
  }
  if (options.chapterStyle !== undefined) {
    attrs.chapStyle = String(options.chapterStyle);
    attrs.chapSep = options.chapterSeparator ?? "hyphen";
  }
  setSectPrChild(doc, "pgNumType", Object.keys(attrs).length ? wEl("pgNumType", attrs) : undefined);
}

/** Read the last section's page numbering. */
export function getPageNumberFormat(doc: Docx): PageNumberFormatOptions {
  const el = wChild(doc.document.body.sectPr, "pgNumType");
  if (!el) return { format: "decimal" };
  const fmt = wAttrValue(el, "fmt");
  const start = wAttrValue(el, "start");
  const chap = wAttrValue(el, "chapStyle");
  const sep = wAttrValue(el, "chapSep");
  // A format this dialog does not offer (e.g. "chicago") reads as decimal.
  return {
    format: fmt !== undefined && isPageNumberFormat(fmt) ? fmt : "decimal",
    ...(start !== undefined ? { start: Number(start) } : {}),
    ...(chap !== undefined ? { chapterStyle: Number(chap) } : {}),
    ...(sep !== undefined && isChapterSeparator(sep) ? { chapterSeparator: sep } : {}),
  };
}

// --- insert another document's content ---------------------------------------------

const NUMBERING_PART_NAME = "/word/numbering.xml";
const FOOTNOTES_PART_NAME = "/word/footnotes.xml";
const ENDNOTES_PART_NAME = "/word/endnotes.xml";

function ensureNumbering(doc: Docx): NonNullable<ReturnType<typeof numberingPart>> {
  const existing = numberingPart(doc);
  if (existing) return existing;
  addPart(doc.opc, {
    name: NUMBERING_PART_NAME,
    contentType: WML_CONTENT_TYPES.numbering,
    data: new TextEncoder().encode(EMPTY_NUMBERING_XML),
  });
  const rels = partRelationships(doc.opc, doc.partName);
  if (relationshipsByType(rels, WML_RELATIONSHIPS.numbering).length === 0) {
    addRelationship(rels, { type: WML_RELATIONSHIPS.numbering, target: "numbering.xml" });
  }
  const part = numberingPart(doc);
  if (!part) throw new Error("Could not create the numbering part.");
  doc.numberingDirty = true;
  return part;
}

function ensureNotes(
  doc: Docx,
  kind: "footnote" | "endnote",
): NonNullable<ReturnType<typeof footnotesPart>> {
  const existing = kind === "footnote" ? footnotesPart(doc) : endnotesPart(doc);
  if (existing) return existing;
  addPart(doc.opc, {
    name: kind === "footnote" ? FOOTNOTES_PART_NAME : ENDNOTES_PART_NAME,
    contentType: kind === "footnote" ? WML_CONTENT_TYPES.footnotes : WML_CONTENT_TYPES.endnotes,
    data: new TextEncoder().encode(kind === "footnote" ? SEED_FOOTNOTES_XML : SEED_ENDNOTES_XML),
  });
  const rels = partRelationships(doc.opc, doc.partName);
  const type = kind === "footnote" ? WML_RELATIONSHIPS.footnotes : WML_RELATIONSHIPS.endnotes;
  if (relationshipsByType(rels, type).length === 0) {
    addRelationship(rels, { type, target: kind === "footnote" ? "footnotes.xml" : "endnotes.xml" });
  }
  const part = kind === "footnote" ? footnotesPart(doc) : endnotesPart(doc);
  if (!part) throw new Error(`Could not create the ${kind}s part.`);
  return part;
}

function idOf(el: XmlElement, local: string): string | undefined {
  return el.attrs.find((a) => a.name.uri === WML_NS && a.name.local === local)?.value;
}

function maxNumericAttr(els: readonly XmlElement[], local: string): number {
  let max = 0;
  for (const e of els) {
    const n = Number(idOf(e, local));
    if (Number.isFinite(n) && n > max) max = n;
  }
  return max;
}

function setWAttr(el: XmlElement, local: string, value: string): XmlElement {
  return {
    ...el,
    attrs: el.attrs.map((a) =>
      a.name.uri === WML_NS && a.name.local === local ? { ...a, value } : a,
    ),
  };
}

/** Map every element of a tree (children first). */
function mapTree(
  node: XmlElement,
  fn: (el: XmlElement) => XmlElement | undefined,
): XmlElement | undefined {
  const children: XmlNode[] = [];
  for (const c of node.children) {
    if (c.kind !== "element") {
      children.push(c);
      continue;
    }
    const mapped = mapTree(c, fn);
    if (mapped) children.push(mapped);
  }
  return fn({ ...node, children });
}

function forEachElement(node: XmlElement, fn: (el: XmlElement) => void): void {
  fn(node);
  for (const c of node.children) if (c.kind === "element") forEachElement(c, fn);
}

function resolveTarget(target: string): string {
  const parts = `/word/${target}`.split("/");
  const out: string[] = [];
  for (const p of parts) {
    if (p === "..") out.pop();
    else if (p !== "." && p !== "") out.push(p);
  }
  return `/${out.join("/")}`;
}

function uniquePartName(doc: Docx, partName: string): string {
  if (!hasPart(doc.opc, partName)) return partName;
  const dot = partName.lastIndexOf(".");
  const stem = dot > 0 ? partName.slice(0, dot) : partName;
  const ext = dot > 0 ? partName.slice(dot) : "";
  let n = 1;
  while (hasPart(doc.opc, `${stem}_${n}${ext}`)) n++;
  return `${stem}_${n}${ext}`;
}

const STRIPPED_INLINES: ReadonlySet<string> = new Set(["commentRangeStart", "commentRangeEnd"]);

/**
 * Insert the body of `source` before body block `blockIndex` (Insert ▸ Object
 * ▸ Text from File). Styles the target lacks are copied (the target's own
 * definition wins on a clash), list definitions are copied under fresh ids,
 * pictures and links get their parts and relationships, footnotes and
 * endnotes are copied with new ids, and bookmark / drawing ids are renumbered
 * so nothing collides. Comments and the source's section breaks are dropped.
 * Returns the number of blocks inserted.
 */
export function insertDocumentContent(doc: Docx, source: Docx, blockIndex: number): number {
  const total = doc.document.body.blocks.length;
  if (!Number.isInteger(blockIndex) || blockIndex < 0 || blockIndex > total) {
    throw new RangeError(`Block index ${blockIndex} is outside the body (0–${total}).`);
  }
  const written = writeWmlDocument(source.document).root;
  const sourceBody = written.children.find((c): c is XmlElement => isW(c, "body"));
  if (!sourceBody) return 0;
  let content = sourceBody.children.filter(
    (c): c is XmlElement => c.kind === "element" && !isW(c, "sectPr"),
  );

  // Styles the target lacks.
  const sourceStyles = stylesPart(source);
  const copiedStyles: XmlElement[] = [];
  if (sourceStyles) {
    const targetStyles = stylesPart(doc);
    for (const style of sourceStyles.styles) {
      const id = idOf(style, "styleId");
      if (!id || (targetStyles && findStyle(targetStyles, id))) continue;
      copiedStyles.push(structuredClone(style));
    }
  }

  // Lists: every num the content or copied styles use, under fresh ids.
  const numMap = new Map<string, string>();
  const usedNums = new Set<string>();
  const collectNums = (el: XmlElement): void =>
    forEachElement(el, (e) => {
      if (isW(e, "numId")) {
        const v = idOf(e, "val");
        if (v && v !== "0") usedNums.add(v);
      }
    });
  for (const c of [...content, ...copiedStyles]) collectNums(c);
  const sourceNumbering = numberingPart(source);
  if (usedNums.size && sourceNumbering) {
    const target = ensureNumbering(doc);
    let nextAbstract = maxNumericAttr(target.abstractNums, "abstractNumId") + 1;
    let nextNum = maxNumericAttr(target.nums, "numId") + 1;
    const abstractMap = new Map<string, string>();
    const sourceNums = new Map(sourceNumbering.nums.map((n) => [idOf(n, "numId"), n] as const));
    const sourceAbstracts = new Map(
      sourceNumbering.abstractNums.map((a) => [idOf(a, "abstractNumId"), a] as const),
    );
    for (const oldNum of usedNums) {
      const num = sourceNums.get(oldNum);
      const absRef = num ? wChild(num, "abstractNumId") : undefined;
      const oldAbs = absRef ? idOf(absRef, "val") : undefined;
      const abs = oldAbs !== undefined ? sourceAbstracts.get(oldAbs) : undefined;
      if (!num || oldAbs === undefined || !abs) continue;
      let newAbs = abstractMap.get(oldAbs);
      if (newAbs === undefined) {
        newAbs = String(nextAbstract++);
        abstractMap.set(oldAbs, newAbs);
        target.abstractNums.push(setWAttr(structuredClone(abs), "abstractNumId", newAbs));
      }
      const newNum = String(nextNum++);
      numMap.set(oldNum, newNum);
      const absEl = newAbs;
      const copied = mapTree(setWAttr(structuredClone(num), "numId", newNum), (e) =>
        isW(e, "abstractNumId") ? setWAttr(e, "val", absEl) : e,
      );
      if (copied) target.nums.push(copied);
    }
    doc.numberingDirty = true;
  }

  // Relationships (pictures, links, embedded objects) referenced by r:* attributes.
  const relMap = new Map<string, string>();
  const sourceRels = partRelationships(source.opc, source.partName);
  const targetRels = partRelationships(doc.opc, doc.partName);
  const relIds = new Set<string>();
  for (const c of content) {
    forEachElement(c, (e) => {
      for (const a of e.attrs) if (a.name.uri === R_NS && !a.isNamespaceDecl) relIds.add(a.value);
    });
  }
  for (const oldId of relIds) {
    const rel = sourceRels.relationships.find((r) => r.id === oldId);
    if (!rel) continue;
    if (rel.targetMode === "External") {
      relMap.set(
        oldId,
        addRelationship(targetRels, { type: rel.type, target: rel.target, targetMode: "External" })
          .id,
      );
      continue;
    }
    const sourcePartName = resolveTarget(rel.target);
    const part = getPart(source.opc, sourcePartName);
    if (!part) continue;
    const newPartName = uniquePartName(doc, sourcePartName);
    addPart(doc.opc, { name: newPartName, contentType: part.contentType, data: part.data });
    const newRel = addRelationship(targetRels, {
      type: rel.type,
      target: newPartName.replace(/^\/word\//, ""),
    });
    relMap.set(oldId, newRel.id);
  }

  // Footnotes / endnotes referenced by the content.
  const noteMaps = { footnote: new Map<string, string>(), endnote: new Map<string, string>() };
  for (const kind of ["footnote", "endnote"] as const) {
    const refLocal = `${kind}Reference`;
    const ids = new Set<string>();
    for (const c of content)
      forEachElement(c, (e) => {
        if (isW(e, refLocal)) {
          const id = idOf(e, "id");
          if (id) ids.add(id);
        }
      });
    if (!ids.size) continue;
    const sourcePart = kind === "footnote" ? footnotesPart(source) : endnotesPart(source);
    if (!sourcePart) continue;
    const target = ensureNotes(doc, kind);
    let next = maxNumericAttr(target.footnotes, "id") + 1;
    const byId = new Map(sourcePart.footnotes.map((f) => [idOf(f, "id"), f] as const));
    for (const id of ids) {
      const note = byId.get(id);
      if (!note) continue;
      const newId = String(next++);
      noteMaps[kind].set(id, newId);
      target.footnotes.push(setWAttr(structuredClone(note), "id", newId));
    }
    if (kind === "footnote") doc.footnotesDirty = true;
    else doc.endnotesDirty = true;
  }

  // Bookmarks: fresh ids; a name the target already uses keeps the target's.
  const targetBookmarkNames = new Set<string>();
  let nextBookmark = 0;
  for (const b of doc.document.body.blocks) {
    const node =
      b.kind === "paragraph" ? paragraphToElement(b) : b.kind === "raw" ? b.node : undefined;
    if (!node) continue;
    forEachElement(node, (e) => {
      if (isW(e, "bookmarkStart")) {
        const name = idOf(e, "name");
        if (name) targetBookmarkNames.add(name);
      }
      if (isW(e, "bookmarkStart") || isW(e, "bookmarkEnd")) {
        const n = Number(idOf(e, "id"));
        if (Number.isFinite(n) && n >= nextBookmark) nextBookmark = n + 1;
      }
    });
  }
  const bookmarkMap = new Map<string, string | null>();
  for (const c of content) {
    forEachElement(c, (e) => {
      if (!isW(e, "bookmarkStart")) return;
      const id = idOf(e, "id");
      const name = idOf(e, "name");
      if (id === undefined) return;
      bookmarkMap.set(id, name && targetBookmarkNames.has(name) ? null : String(nextBookmark++));
    });
  }

  let nextDocPr = 1 + maxDocPrId(doc);
  const remapAttrs = (attrs: readonly XmlAttr[]): XmlAttr[] =>
    attrs.map((a) =>
      a.name.uri === R_NS && !a.isNamespaceDecl && relMap.has(a.value)
        ? { ...a, value: relMap.get(a.value) ?? a.value }
        : a,
    );
  const transform = (e: XmlElement): XmlElement | undefined => {
    if (e.name.uri === WML_NS && STRIPPED_INLINES.has(e.name.local)) return undefined;
    if (isW(e, "sectPr")) return undefined;
    let out: XmlElement = { ...e, attrs: remapAttrs(e.attrs) };
    if (isW(out, "numId")) {
      const v = idOf(out, "val");
      const mapped = v ? numMap.get(v) : undefined;
      if (mapped) out = setWAttr(out, "val", mapped);
    }
    for (const kind of ["footnote", "endnote"] as const) {
      if (isW(out, `${kind}Reference`)) {
        const mapped = noteMaps[kind].get(idOf(out, "id") ?? "");
        if (mapped) out = setWAttr(out, "id", mapped);
      }
    }
    if (isW(out, "bookmarkStart") || isW(out, "bookmarkEnd")) {
      const mapped = bookmarkMap.get(idOf(out, "id") ?? "");
      if (mapped === null) return undefined;
      if (mapped) out = setWAttr(out, "id", mapped);
    }
    if (out.name.local === "docPr" && out.attrs.some((a) => a.name.local === "id")) {
      out = {
        ...out,
        attrs: out.attrs.map((a) =>
          a.name.local === "id" ? { ...a, value: String(nextDocPr++) } : a,
        ),
      };
    }
    return out;
  };
  // Runs holding a comment reference mark go with the comment.
  const withoutCommentRuns = (e: XmlElement): XmlElement | undefined =>
    isW(e, "r") && e.children.some((c) => isW(c, "commentReference")) ? undefined : e;
  content = content.flatMap((c) => {
    const pruned = mapTree(c, withoutCommentRuns);
    const mapped = pruned ? mapTree(pruned, transform) : undefined;
    return mapped ? [mapped] : [];
  });

  // Copied styles may use lists too.
  if (copiedStyles.length) {
    const target = stylesPart(doc);
    if (target) {
      for (const s of copiedStyles) {
        const mapped = mapTree(s, (e) => {
          if (!isW(e, "numId")) return e;
          const v = idOf(e, "val");
          const m = v ? numMap.get(v) : undefined;
          return m ? setWAttr(e, "val", m) : e;
        });
        if (mapped) target.styles.push(mapped);
      }
      doc.stylesDirty = true;
    }
  }

  const wrapper = parseWmlDocument({
    prologue: [],
    epilogue: [],
    root: { ...written, children: [{ ...sourceBody, children: content }] },
  });
  const blocks = wrapper.body.blocks;
  doc.document.body.blocks.splice(blockIndex, 0, ...blocks);
  doc.dirty = true;
  return blocks.length;
}

function maxDocPrId(doc: Docx): number {
  let max = 0;
  const visit = (node: XmlElement): void =>
    forEachElement(node, (e) => {
      if (e.name.local === "docPr") {
        const n = Number(e.attrs.find((a) => a.name.local === "id")?.value);
        if (Number.isFinite(n) && n > max) max = n;
      }
    });
  for (const b of doc.document.body.blocks) {
    if (b.kind === "paragraph") visit(paragraphToElement(b));
    else if (b.kind === "raw") visit(b.node);
  }
  return max;
}
