/**
 * The Insert-tab operations: content placed at a caret (fields, symbols,
 * links, bookmarks, equations, drop caps, signature lines) and document-level
 * inserts (cover pages, headers and footers, page numbers, text from another
 * document). Positions are a paragraph plus a character offset in the unit
 * {@link runTextLength} counts, so an editor's caret maps onto them directly.
 */

import {
  addRelationship,
  addPart,
  hasPart,
  partRelationships,
  setContentTypeDefault,
} from "../internal/opc/index.js";
import { buildMathFromLinear, mathToLinear, OMML_NS } from "../internal/math/linear.js";
import {
  extensionForImageContentType,
  sniffImageContentType,
  WML_NS,
  WML_RELATIONSHIPS,
  type WmlInline,
  type WmlParagraph,
  type WmlRun,
} from "../internal/wordprocessingml/index.js";
import type { XmlElement } from "../internal/xml/index.js";
import { bookmarks, type Docx } from "./docx.js";
import {
  assertOffset,
  attr,
  bodyParagraphs,
  childIndexAtOffset,
  el,
  isW,
  nsDecl,
  paragraphSlot,
  R_NS,
  runVisibleText,
  setParagraphChild,
  setRunChild,
  spliceInlinesAt,
  textRun,
  VML_NS,
  VML_OFFICE_NS,
  wAttr,
  wAttrValue,
  wChild,
  wEl,
  xmlVisibleText,
} from "./insert-xml.js";

export * from "./insert-fields.js";
export * from "./insert-content.js";
export { mathToLinear as equationToLinear } from "../internal/math/linear.js";

// --- symbols and special characters --------------------------------------------

/**
 * What {@link insertSymbol} places: a character from a symbol font as
 * `<w:sym>` (§17.3.3.30, `char` is the hex code, e.g. `"F0E0"` in Wingdings),
 * or one of the two hyphen elements Word's Special Characters tab inserts.
 */
export type SymbolSpec =
  | { readonly font: string; readonly char: string }
  | "noBreakHyphen"
  | "softHyphen";

const SYM_CHAR = /^[0-9A-Fa-f]{4}$/;

/** Insert a symbol-font character or a special hyphen at a character offset. */
export function insertSymbol(
  doc: Docx,
  paragraph: WmlParagraph,
  offset: number,
  symbol: SymbolSpec,
  rPr?: XmlElement,
): WmlRun {
  assertOffset(paragraph, offset);
  let run: WmlRun;
  if (symbol === "noBreakHyphen" || symbol === "softHyphen") {
    run = { kind: "run", pieces: [{ kind: symbol }], extras: [] };
  } else {
    if (!SYM_CHAR.test(symbol.char)) {
      throw new RangeError(`w:sym char must be four hex digits, got ${JSON.stringify(symbol.char)}.`);
    }
    if (symbol.font.trim() === "") throw new Error("A symbol needs its font.");
    run = {
      kind: "run",
      pieces: [{ kind: "symbol", font: symbol.font, char: symbol.char.toUpperCase() }],
      extras: [],
    };
  }
  if (rPr) run.rPr = structuredClone(rPr);
  spliceInlinesAt(paragraph, offset, [run]);
  doc.dirty = true;
  return run;
}

// --- hyperlinks -------------------------------------------------------------------

// Schemes a hyperlink may target. Anything else (notably `javascript:` /
// `data:`) is rejected: the target ends up as a clickable link in Word and in
// any HTML rendering of the document.
const ALLOWED_LINK_PROTOCOLS: ReadonlySet<string> = new Set([
  "http:",
  "https:",
  "mailto:",
  "ftp:",
  "file:",
  "tel:",
]);

/** Where a link points: a web/mail address, a bookmark in this document, or both. */
export interface HyperlinkTarget {
  /** External address (`https://…`, `mailto:…`). */
  readonly url?: string;
  /** Bookmark name (`w:anchor`); with `url`, a location inside that resource. */
  readonly bookmark?: string;
  /** ScreenTip (`w:tooltip`). */
  readonly tooltip?: string;
  /** Target frame (`w:tgtFrame`): `_blank`, `_top`, `_self`, `_parent` or a frame name. */
  readonly targetFrame?: string;
}

export interface HyperlinkInfo extends HyperlinkTarget {
  readonly element: XmlElement;
  readonly text: string;
}

function assertSafeUrl(url: string): void {
  let protocol: string;
  try {
    protocol = new URL(url).protocol;
  } catch {
    throw new Error(`Invalid hyperlink address: ${JSON.stringify(url)}.`);
  }
  if (!ALLOWED_LINK_PROTOCOLS.has(protocol)) {
    throw new Error(`Unsupported hyperlink scheme ${JSON.stringify(protocol)}.`);
  }
}

function hyperlinkAttrs(doc: Docx, target: HyperlinkTarget): XmlElement["attrs"] {
  if (!target.url && !target.bookmark) throw new Error("A link needs an address or a bookmark.");
  const attrs = [];
  if (target.url) {
    assertSafeUrl(target.url);
    const rel = addRelationship(partRelationships(doc.opc, doc.partName), {
      type: WML_RELATIONSHIPS.hyperlink,
      target: target.url,
      targetMode: "External",
    });
    attrs.push(attr(R_NS, "r", "id", rel.id));
  }
  if (target.bookmark) attrs.push(wAttr("anchor", target.bookmark));
  if (target.targetFrame) attrs.push(wAttr("tgtFrame", target.targetFrame));
  if (target.tooltip) attrs.push(wAttr("tooltip", target.tooltip));
  attrs.push(wAttr("history", "1"));
  return attrs;
}

function linkRun(text: string): WmlRun {
  return textRun(text, wEl("rPr", {}, [wEl("rStyle", { val: "Hyperlink" })]));
}

/**
 * Insert a `<w:hyperlink>` showing `text` at a character offset. Built-in
 * character style `Hyperlink` is referenced (Word defines it on first use; a
 * document without it shows the text unstyled).
 */
export function insertHyperlink(
  doc: Docx,
  paragraph: WmlParagraph,
  offset: number,
  text: string,
  target: HyperlinkTarget,
): XmlElement {
  assertOffset(paragraph, offset);
  if (text === "") throw new Error("A link needs text to display.");
  const element = el(WML_NS, "w", "hyperlink", hyperlinkAttrs(doc, target), [
    runElement(linkRun(text)),
  ]);
  spliceInlinesAt(paragraph, offset, [{ kind: "raw", node: element }]);
  doc.dirty = true;
  return element;
}

function runElement(run: WmlRun): XmlElement {
  // Hyperlink children are raw XML; build the run through the paragraph parser
  // round trip's inverse: a <w:r> with the run's rPr and text.
  const children: XmlElement[] = [];
  if (run.rPr) children.push(run.rPr);
  for (const p of run.pieces) {
    if (p.kind === "text") {
      children.push(
        el(WML_NS, "w", "t", p.preserveSpace ? [preserveAttr()] : [], [{ kind: "text", value: p.value }]),
      );
    } else if (p.kind === "tab") children.push(wEl("tab"));
    else if (p.kind === "break") children.push(wEl("br"));
  }
  return el(WML_NS, "w", "r", [], children);
}

function preserveAttr(): XmlElement["attrs"][number] {
  return {
    name: { uri: "http://www.w3.org/XML/1998/namespace", local: "space", prefix: "xml" },
    value: "preserve",
    isNamespaceDecl: false,
  };
}

/** The hyperlinks directly inside a paragraph, in order. */
export function paragraphHyperlinks(doc: Docx, paragraph: WmlParagraph): HyperlinkInfo[] {
  const rels = partRelationships(doc.opc, doc.partName);
  const out: HyperlinkInfo[] = [];
  for (const child of paragraph.children) {
    if (child.kind !== "raw" || !isW(child.node, "hyperlink")) continue;
    const node = child.node;
    const relId = node.attrs.find((a) => a.name.uri === R_NS && a.name.local === "id")?.value;
    const url = relId ? rels.relationships.find((r) => r.id === relId)?.target : undefined;
    const bookmark = wAttrValue(node, "anchor");
    const tooltip = wAttrValue(node, "tooltip");
    const targetFrame = wAttrValue(node, "tgtFrame");
    out.push({
      element: node,
      text: xmlVisibleText(node),
      ...(url !== undefined ? { url } : {}),
      ...(bookmark !== undefined ? { bookmark } : {}),
      ...(tooltip !== undefined ? { tooltip } : {}),
      ...(targetFrame !== undefined ? { targetFrame } : {}),
    });
  }
  return out;
}

/**
 * Change an existing hyperlink's target and (optionally) its text. The old
 * relationship is left in place: other links may share it, and an unused
 * external relationship is harmless.
 */
export function editHyperlink(
  doc: Docx,
  paragraph: WmlParagraph,
  hyperlink: XmlElement,
  target: HyperlinkTarget,
  text?: string,
): XmlElement {
  const index = paragraph.children.findIndex((c) => c.kind === "raw" && c.node === hyperlink);
  if (index < 0) throw new Error("The hyperlink is not in this paragraph.");
  const children =
    text === undefined || text === xmlVisibleText(hyperlink)
      ? hyperlink.children
      : [runElement(linkRun(text))];
  const next = el(WML_NS, "w", "hyperlink", hyperlinkAttrs(doc, target), children);
  paragraph.children[index] = { kind: "raw", node: next };
  doc.dirty = true;
  return next;
}

/** Remove a hyperlink but keep its text (Word's Remove Link). */
export function removeHyperlink(doc: Docx, paragraph: WmlParagraph, hyperlink: XmlElement): void {
  const index = paragraph.children.findIndex((c) => c.kind === "raw" && c.node === hyperlink);
  if (index < 0) throw new Error("The hyperlink is not in this paragraph.");
  const runs: WmlInline[] = [];
  for (const c of hyperlink.children) {
    if (c.kind !== "element") continue;
    if (isW(c, "r")) {
      // Drop the Hyperlink character style; keep any other formatting.
      const rPr = wChild(c, "rPr");
      const keptRPr = rPr
        ? { ...rPr, children: rPr.children.filter((x) => !isW(x, "rStyle")) }
        : undefined;
      const rest = c.children.filter((x) => !isW(x, "rPr"));
      const node = el(WML_NS, "w", "r", c.attrs, keptRPr && keptRPr.children.length ? [keptRPr, ...rest] : rest);
      runs.push({ kind: "raw", node });
    } else {
      runs.push({ kind: "raw", node: c });
    }
  }
  paragraph.children.splice(index, 1, ...runs);
  doc.dirty = true;
}

// --- bookmarks ----------------------------------------------------------------------

// Word: a bookmark name starts with a letter (or `_` for hidden ones), uses
// letters, digits and underscores, and is at most 40 characters.
const BOOKMARK_NAME = /^[\p{L}_][\p{L}\p{N}_]{0,39}$/u;

/** Whether `name` is a bookmark name Word accepts. */
export function isValidBookmarkName(name: string): boolean {
  return BOOKMARK_NAME.test(name);
}

/** A point in the body: a paragraph and a character offset in it. */
export interface TextPoint {
  readonly paragraph: WmlParagraph;
  readonly offset: number;
}

function nextBookmarkId(doc: Docx): number {
  let max = -1;
  const scan = (node: XmlElement): void => {
    if (isW(node, "bookmarkStart") || isW(node, "bookmarkEnd")) {
      const n = Number(wAttrValue(node, "id"));
      if (Number.isFinite(n) && n > max) max = n;
    }
    for (const c of node.children) if (c.kind === "element") scan(c);
  };
  for (const p of bodyParagraphs(doc)) {
    for (const c of p.children) if (c.kind === "raw") scan(c.node);
  }
  for (const b of doc.document.body.blocks) if (b.kind === "raw") scan(b.node);
  return max + 1;
}

/**
 * Add a bookmark spanning `start`–`end` (a point bookmark when `end` is
 * omitted). The two points may be in different paragraphs, start first.
 * Throws when the name is invalid or already used.
 */
export function insertBookmark(doc: Docx, name: string, start: TextPoint, end: TextPoint = start): number {
  if (!isValidBookmarkName(name)) {
    throw new Error(
      `Invalid bookmark name ${JSON.stringify(name)}: use letters, digits and underscores, start with a letter, at most 40 characters.`,
    );
  }
  if (bookmarks(doc).some((b) => b.name === name)) {
    throw new Error(`A bookmark named ${JSON.stringify(name)} already exists.`);
  }
  assertOffset(start.paragraph, start.offset);
  assertOffset(end.paragraph, end.offset);
  const id = String(nextBookmarkId(doc));
  // End first: when both points share a paragraph, inserting the end marker
  // does not shift the start offset.
  spliceInlinesAt(end.paragraph, end.offset, [{ kind: "raw", node: wEl("bookmarkEnd", { id }) }]);
  const startIndex = childIndexAtOffset(start.paragraph, start.offset);
  start.paragraph.children.splice(startIndex, 0, {
    kind: "raw",
    node: wEl("bookmarkStart", { id, name }),
  });
  doc.dirty = true;
  return Number(id);
}

const REF_BOOKMARK_PREFIX = "_Ref";
// Word numbers its hidden reference bookmarks with nine digits.
const REF_BOOKMARK_DIGITS = 9;

/**
 * The hidden `_Ref…` bookmark Word puts around a cross-reference target (a
 * heading's text, a caption, a footnote reference), creating it when the
 * paragraph has none. With `run`, the bookmark wraps just that run (a
 * footnote reference mark); otherwise the paragraph's whole text.
 */
export function ensureReferenceBookmark(doc: Docx, paragraph: WmlParagraph, run?: WmlRun): string {
  const existing = paragraph.children.find(
    (c) =>
      c.kind === "raw" &&
      isW(c.node, "bookmarkStart") &&
      (wAttrValue(c.node, "name") ?? "").startsWith(REF_BOOKMARK_PREFIX) &&
      (!run || paragraph.children[paragraph.children.indexOf(c) + 1] === run),
  );
  if (existing?.kind === "raw") return wAttrValue(existing.node, "name") ?? "";
  const taken = new Set(bookmarks(doc).map((b) => b.name));
  let n = taken.size + 1;
  let name = `${REF_BOOKMARK_PREFIX}${String(n).padStart(REF_BOOKMARK_DIGITS, "0")}`;
  while (taken.has(name)) {
    n++;
    name = `${REF_BOOKMARK_PREFIX}${String(n).padStart(REF_BOOKMARK_DIGITS, "0")}`;
  }
  const id = String(nextBookmarkId(doc));
  const start: WmlInline = { kind: "raw", node: wEl("bookmarkStart", { id, name }) };
  const end: WmlInline = { kind: "raw", node: wEl("bookmarkEnd", { id }) };
  if (run) {
    const at = paragraph.children.indexOf(run);
    if (at < 0) throw new Error("The run is not in this paragraph.");
    paragraph.children.splice(at, 1, start, run, end);
  } else {
    // Around the text runs, leaving the paragraph's own bookmarks outside.
    const first = paragraph.children.findIndex((c) => c.kind === "run");
    const lastRun = paragraph.children.findLastIndex((c) => c.kind === "run");
    if (first < 0) paragraph.children.push(start, end);
    else {
      paragraph.children.splice(lastRun + 1, 0, end);
      paragraph.children.splice(first, 0, start);
    }
  }
  doc.dirty = true;
  return name;
}

// --- equations ------------------------------------------------------------------------

/**
 * Insert an equation built from Word's linear format (`a/b`, `x^2`,
 * `√(x+1)`, `∑_(i=1)^n▒i`, `■(a&b@c&d)`, `\alpha` …) at a character offset.
 * `display: true` writes `<m:oMathPara>` — a centered equation on its own
 * line, as Word inserts into an empty paragraph.
 */
export function insertEquation(
  doc: Docx,
  paragraph: WmlParagraph,
  offset: number,
  linear: string,
  options: { readonly display?: boolean } = {},
): XmlElement {
  assertOffset(paragraph, offset);
  const element = buildMathFromLinear(linear, options.display ?? false);
  spliceInlinesAt(paragraph, offset, [{ kind: "raw", node: element }]);
  doc.dirty = true;
  return element;
}

function isMath(node: XmlElement): boolean {
  return node.name.uri === OMML_NS && (node.name.local === "oMath" || node.name.local === "oMathPara");
}

/** The equations (`m:oMath` / `m:oMathPara`) directly in a paragraph, in order. */
export function paragraphEquations(paragraph: WmlParagraph): XmlElement[] {
  return paragraph.children.flatMap((c) => (c.kind === "raw" && isMath(c.node) ? [c.node] : []));
}

/** Replace the `index`-th equation of a paragraph with one built from `linear`. */
export function setEquation(doc: Docx, paragraph: WmlParagraph, index: number, linear: string): XmlElement {
  const target = paragraphEquations(paragraph)[index];
  if (!target) throw new RangeError(`The paragraph has no equation ${index}.`);
  const display = target.name.local === "oMathPara";
  const next = buildMathFromLinear(linear, display);
  const at = paragraph.children.findIndex((c) => c.kind === "raw" && c.node === target);
  paragraph.children[at] = { kind: "raw", node: next };
  doc.dirty = true;
  return next;
}

/** Build an equation element without placing it (for headers, notes, tests). */
export function buildEquation(linear: string, options: { readonly display?: boolean } = {}): XmlElement {
  return buildMathFromLinear(linear, options.display ?? false);
}

/** Read an equation back as linear-format text (for re-editing). */
export function equationLinear(element: XmlElement): string {
  return mathToLinear(element);
}

// --- drop caps -----------------------------------------------------------------------

export interface DropCapOptions {
  /** `drop`: in the text; `margin`: in the left margin; `none`: remove. */
  readonly position: "drop" | "margin" | "none";
  /** Lines to drop (1–10). Default 3. */
  readonly lines?: number;
  /** Distance from text in twips (`w:hSpace`). Default 0. */
  readonly distanceTwips?: number;
  /** Font of the dropped letter; the paragraph's font when omitted. */
  readonly font?: string;
}

const MIN_DROP_LINES = 1;
const MAX_DROP_LINES = 10;
const DEFAULT_DROP_LINES = 3;
const DEFAULT_FONT_HALF_POINTS = 22;
// Word sizes the letter so its cap height spans `lines` lines and sets an
// exact line height and a baseline shift to match. These ratios reproduce what
// Word for Mac writes for its default fonts (3 lines of 11 pt → sz 111,
// line 827, position −8).
const DROP_CAP_SIZE_PER_LINE = 1.68;
const DROP_CAP_LINE_RATIO = 0.745;
const DROP_CAP_POSITION_RATIO = 0.072;
const TWIPS_PER_POINT = 20;

function isDropCapParagraph(para: WmlParagraph | undefined): boolean {
  const framePr = para?.pPr ? wChild(para.pPr, "framePr") : undefined;
  const dropCap = framePr ? wAttrValue(framePr, "dropCap") : undefined;
  return dropCap === "drop" || dropCap === "margin";
}

/** The drop cap state of a paragraph (the frame paragraph Word puts before it). */
export function getDropCap(doc: Docx, paragraph: WmlParagraph): DropCapOptions {
  const prev = paragraphSlot(doc, paragraph).sibling(-1);
  if (!prev || !isDropCapParagraph(prev)) return { position: "none" };
  const framePr = prev.pPr ? wChild(prev.pPr, "framePr") : undefined;
  const position = framePr && wAttrValue(framePr, "dropCap") === "margin" ? "margin" : "drop";
  const lines = Number(framePr ? wAttrValue(framePr, "lines") : undefined) || DEFAULT_DROP_LINES;
  const distance = Number(framePr ? wAttrValue(framePr, "hSpace") : undefined) || 0;
  const run = prev.children.find((c): c is WmlRun => c.kind === "run");
  const fonts = run?.rPr ? wChild(run.rPr, "rFonts") : undefined;
  const font = fonts ? wAttrValue(fonts, "ascii") : undefined;
  return { position, lines, distanceTwips: distance, ...(font ? { font } : {}) };
}

/**
 * Make (or change, or with `position: "none"` remove) a drop cap on a
 * paragraph. As Word does, the first letter moves into a frame paragraph
 * (`w:framePr w:dropCap`, §17.3.1.11) placed just before it, and removing
 * the drop cap moves the letter back.
 */
export function setDropCap(doc: Docx, paragraph: WmlParagraph, options: DropCapOptions): void {
  const slot = paragraphSlot(doc, paragraph);
  const prev = slot.sibling(-1);
  if (prev && isDropCapParagraph(prev)) {
    // Undo the existing drop cap first; a new one is built from scratch.
    const letterRuns = prev.children.filter((c): c is WmlRun => c.kind === "run");
    for (const r of letterRuns) {
      setRunChild(r, "position", undefined);
      setRunChild(r, "sz", undefined);
      setRunChild(r, "szCs", undefined);
      if (r.rPr && r.rPr.children.length === 0) delete r.rPr;
    }
    paragraph.children.unshift(...letterRuns);
    slot.remove(slot.index - 1, 1);
  }
  if (options.position === "none") {
    doc.dirty = true;
    return;
  }
  const lines = options.lines ?? DEFAULT_DROP_LINES;
  if (!Number.isInteger(lines) || lines < MIN_DROP_LINES || lines > MAX_DROP_LINES) {
    throw new RangeError(`Lines to drop must be ${MIN_DROP_LINES}–${MAX_DROP_LINES}, got ${lines}.`);
  }
  const distance = options.distanceTwips ?? 0;
  if (!Number.isFinite(distance) || distance < 0) throw new RangeError("Distance must be ≥ 0.");
  const here = paragraphSlot(doc, paragraph);
  const firstRun = paragraph.children.find(
    (c): c is WmlRun => c.kind === "run" && runVisibleText(c).trim() !== "",
  );
  if (!firstRun) throw new Error("A drop cap needs a paragraph that starts with text.");
  const textPiece = firstRun.pieces.find((p) => p.kind === "text");
  if (!textPiece || textPiece.kind !== "text") throw new Error("A drop cap needs text.");
  const letter = [...textPiece.value][0] ?? "";
  textPiece.value = textPiece.value.slice(letter.length);
  const sizeEl = firstRun.rPr ? wChild(firstRun.rPr, "sz") : undefined;
  const baseHalfPoints = Number(sizeEl ? wAttrValue(sizeEl, "val") : undefined) || DEFAULT_FONT_HALF_POINTS;
  const sz = Math.round(lines * baseHalfPoints * DROP_CAP_SIZE_PER_LINE);
  const line = Math.round((sz / 2) * DROP_CAP_LINE_RATIO * TWIPS_PER_POINT);
  const position = -Math.round(sz * DROP_CAP_POSITION_RATIO);
  const letterRun = textRun(letter, firstRun.rPr);
  if (options.font) {
    setRunChild(letterRun, "rFonts", wEl("rFonts", { ascii: options.font, hAnsi: options.font }));
  }
  setRunChild(letterRun, "position", wEl("position", { val: String(position) }));
  setRunChild(letterRun, "sz", wEl("sz", { val: String(sz) }));
  setRunChild(letterRun, "szCs", wEl("szCs", { val: String(sz) }));
  const frame: WmlParagraph = { kind: "paragraph", children: [letterRun], extras: [] };
  if (paragraph.pPr) {
    const style = wChild(paragraph.pPr, "pStyle");
    if (style) setParagraphChild(frame, "pStyle", structuredClone(style));
  }
  setParagraphChild(frame, "keepNext", wEl("keepNext"));
  const frameAttrs: Record<string, string> = {
    dropCap: options.position,
    lines: String(lines),
    wrap: "around",
    vAnchor: "text",
    hAnchor: options.position === "margin" ? "page" : "text",
  };
  if (distance > 0) frameAttrs.hSpace = String(Math.round(distance));
  setParagraphChild(frame, "framePr", wEl("framePr", frameAttrs));
  setParagraphChild(frame, "spacing", wEl("spacing", { after: "0", line: String(line), lineRule: "exact" }));
  setParagraphChild(frame, "textAlignment", wEl("textAlignment", { val: "baseline" }));
  here.insert(here.index, frame);
  doc.dirty = true;
}

// --- signature line (VML, ECMA-376 Part 4 §19.1.2.24 o:signatureline) ----------------

export interface SignatureLineOptions {
  /** PNG/JPEG of the line as Word draws it (the "X", rule and signer text). */
  readonly image: Uint8Array;
  readonly widthPt: number;
  readonly heightPt: number;
  readonly suggestedSigner?: string;
  readonly suggestedSignerTitle?: string;
  readonly suggestedSignerEmail?: string;
  readonly instructions?: string;
  readonly allowComments?: boolean;
  readonly showSignDate?: boolean;
}

// The provider id Word writes for its own (default) signature provider.
const DEFAULT_SIGNATURE_PROVIDER = "{00000000-0000-0000-0000-000000000000}";
// `o:spt="75"` is VML's picture-frame shape type (ECMA-376 Part 4 §19.1.2.20).
const PICTURE_SHAPE_TYPE = "75";

function nextMediaPartName(doc: Docx, ext: string): string {
  let n = 1;
  while (hasPart(doc.opc, `/word/media/image${n}.${ext}`)) n++;
  return `/word/media/image${n}.${ext}`;
}

/**
 * Insert a Microsoft Office signature line: a VML picture carrying
 * `<o:signatureline>` with the suggested signer and options, which Word shows
 * as a signable line. The caller supplies the picture of the line.
 */
export function insertSignatureLine(
  doc: Docx,
  paragraph: WmlParagraph,
  offset: number,
  options: SignatureLineOptions,
): WmlRun {
  assertOffset(paragraph, offset);
  const contentType = sniffImageContentType(options.image);
  if (!contentType) throw new Error("The signature line picture must be a PNG, JPEG, GIF or BMP.");
  if (!(options.widthPt > 0) || !(options.heightPt > 0)) {
    throw new RangeError("The signature line needs a positive size.");
  }
  const ext = extensionForImageContentType(contentType);
  const partName = nextMediaPartName(doc, ext);
  addPart(doc.opc, { name: partName, contentType, data: options.image });
  setContentTypeDefault(doc.opc.contentTypes, ext, contentType);
  const rel = addRelationship(partRelationships(doc.opc, doc.partName), {
    type: WML_RELATIONSHIPS.image,
    target: partName.replace(/^\/word\//, ""),
  });
  const v = (local: string, attrs: XmlElement["attrs"] = [], children: XmlElement[] = []): XmlElement =>
    el(VML_NS, "v", local, attrs, children);
  const o = (local: string, attrs: XmlElement["attrs"] = []): XmlElement =>
    el(VML_OFFICE_NS, "o", local, attrs);
  const plain = (local: string, value: string): XmlElement["attrs"][number] => ({
    name: { uri: "", local, prefix: "" },
    value,
    isNamespaceDecl: false,
  });
  const sigAttrs = [
    attr(VML_NS, "v", "ext", "edit"),
    plain("id", `{${crypto.randomUUID().toUpperCase()}}`),
    plain("provid", DEFAULT_SIGNATURE_PROVIDER),
  ];
  if (options.suggestedSigner) sigAttrs.push(attr(VML_OFFICE_NS, "o", "suggestedsigner", options.suggestedSigner));
  if (options.suggestedSignerTitle) {
    sigAttrs.push(attr(VML_OFFICE_NS, "o", "suggestedsigner2", options.suggestedSignerTitle));
  }
  if (options.suggestedSignerEmail) {
    sigAttrs.push(attr(VML_OFFICE_NS, "o", "suggestedsigneremail", options.suggestedSignerEmail));
  }
  if (options.instructions) {
    sigAttrs.push(plain("signinginstructionsset", "t"), attr(VML_OFFICE_NS, "o", "signinginstructions", options.instructions));
  }
  if (options.allowComments) sigAttrs.push(plain("allowcomments", "t"));
  if (options.showSignDate === false) sigAttrs.push(plain("showsigndate", "f"));
  sigAttrs.push(plain("issignatureline", "t"));
  const shape = v(
    "shape",
    [
      plain("id", `_x0000_i${nextBookmarkId(doc) + 1025}`),
      plain("type", `#_x0000_t${PICTURE_SHAPE_TYPE}`),
      plain("alt", "Microsoft Office Signature Line..."),
      plain("style", `width:${options.widthPt}pt;height:${options.heightPt}pt`),
    ],
    [
      v("imagedata", [attr(R_NS, "r", "id", rel.id), attr(VML_OFFICE_NS, "o", "title", "")]),
      o("lock", [
        attr(VML_NS, "v", "ext", "edit"),
        plain("ungrouping", "t"),
        plain("rotation", "t"),
        plain("cropping", "t"),
        plain("verticies", "t"),
        plain("text", "t"),
        plain("grouping", "t"),
      ]),
      o("signatureline", sigAttrs),
    ],
  );
  const shapetype = v(
    "shapetype",
    [
      plain("id", `_x0000_t${PICTURE_SHAPE_TYPE}`),
      plain("coordsize", "21600,21600"),
      attr(VML_OFFICE_NS, "o", "spt", PICTURE_SHAPE_TYPE),
      attr(VML_OFFICE_NS, "o", "preferrelative", "t"),
      plain("filled", "f"),
      plain("stroked", "f"),
    ],
    [v("stroke", [plain("joinstyle", "miter")]), o("lock", [attr(VML_NS, "v", "ext", "edit"), plain("aspectratio", "t")])],
  );
  const pict: XmlElement = {
    ...el(WML_NS, "w", "pict", [nsDecl("v", VML_NS), nsDecl("o", VML_OFFICE_NS), nsDecl("r", R_NS)], [shapetype, shape]),
  };
  const run: WmlRun = { kind: "run", pieces: [{ kind: "pict", node: pict }], extras: [] };
  spliceInlinesAt(paragraph, offset, [run]);
  doc.dirty = true;
  return run;
}
