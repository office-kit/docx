/**
 * Small XML / run construction helpers shared by the Insert-tab APIs
 * (`insert.ts`, `insert-content.ts`). Not re-exported from the package index.
 *
 * @internal
 */

import { isolateParagraphRunRange, type Docx, runTextLength } from "./docx.js";
import {
  parseParagraph,
  WML_NS,
  type WmlInline,
  type WmlParagraph,
  type WmlRun,
  type WmlRunPiece,
} from "../internal/wordprocessingml/index.js";
import {
  XML_NAMESPACE,
  type XmlAttr,
  type XmlElement,
  type XmlNode,
} from "../internal/xml/index.js";

export const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
export const VML_NS = "urn:schemas-microsoft-com:vml";
export const VML_OFFICE_NS = "urn:schemas-microsoft-com:office:office";
const XMLNS_URI = "http://www.w3.org/2000/xmlns/";

export function wAttr(local: string, value: string): XmlAttr {
  return { name: { uri: WML_NS, local, prefix: "w" }, value, isNamespaceDecl: false };
}

export function nsDecl(prefix: string, uri: string): XmlAttr {
  return {
    name: { uri: XMLNS_URI, local: prefix, prefix: "xmlns" },
    value: uri,
    isNamespaceDecl: true,
  };
}

/** `<w:local w:k="v" …>children</w:local>`. */
export function wEl(
  local: string,
  attrs: Readonly<Record<string, string>> = {},
  children: readonly XmlNode[] = [],
): XmlElement {
  return {
    kind: "element",
    name: { uri: WML_NS, local, prefix: "w" },
    attrs: Object.entries(attrs).map(([k, v]) => wAttr(k, v)),
    children,
    xmlSpace: "default",
    selfClosing: children.length === 0,
  };
}

export function el(
  uri: string,
  prefix: string,
  local: string,
  attrs: readonly XmlAttr[] = [],
  children: readonly XmlNode[] = [],
): XmlElement {
  return {
    kind: "element",
    name: { uri, local, prefix },
    attrs,
    children,
    xmlSpace: "default",
    selfClosing: children.length === 0,
  };
}

export function attr(uri: string, prefix: string, local: string, value: string): XmlAttr {
  return { name: { uri, local, prefix }, value, isNamespaceDecl: false };
}

declare const NAMED_ELEMENT: unique symbol;
/**
 * An element whose name a guard has checked. The brand keeps the guard's
 * false branch from excluding every `XmlElement` (a name check says nothing
 * about other elements).
 */
export type NamedElement = XmlElement & { readonly [NAMED_ELEMENT]: true };

export function isW(node: XmlNode | undefined, local: string): node is NamedElement {
  return node?.kind === "element" && node.name.uri === WML_NS && node.name.local === local;
}

export function wChild(parent: XmlElement | undefined, local: string): XmlElement | undefined {
  return parent?.children.find((c): c is XmlElement => isW(c, local));
}

export function wAttrValue(node: XmlElement, local: string): string | undefined {
  return node.attrs.find((a) => a.name.uri === WML_NS && a.name.local === local)?.value;
}

/** Text pieces for `text`, with `\t` as tabs and `\n` as line breaks. */
export function textPieces(text: string): WmlRunPiece[] {
  const pieces: WmlRunPiece[] = [];
  for (const part of text.split(/(\t|\n)/)) {
    if (part === "\t") pieces.push({ kind: "tab" });
    else if (part === "\n") pieces.push({ kind: "break" });
    else if (part)
      pieces.push({ kind: "text", value: part, preserveSpace: /^\s|\s$|\s\s/.test(part) });
  }
  return pieces;
}

export function textRun(text: string, rPr?: XmlElement): WmlRun {
  return {
    kind: "run",
    ...(rPr ? { rPr: structuredClone(rPr) } : {}),
    pieces: textPieces(text),
    extras: [],
  };
}

export function emptyRun(): WmlRun {
  return { kind: "run", pieces: [], extras: [] };
}

/**
 * Schema order of the `<w:pPr>` children we write (§17.3.1.26). Elements we
 * insert go before the first existing sibling that comes later in this list.
 */
export const PPR_ORDER = [
  "pStyle",
  "keepNext",
  "keepLines",
  "pageBreakBefore",
  "framePr",
  "widowControl",
  "numPr",
  "suppressLineNumbers",
  "pBdr",
  "shd",
  "tabs",
  "suppressAutoHyphens",
  "kinsoku",
  "wordWrap",
  "overflowPunct",
  "topLinePunct",
  "autoSpaceDE",
  "autoSpaceDN",
  "bidi",
  "adjustRightInd",
  "snapToGrid",
  "spacing",
  "ind",
  "contextualSpacing",
  "mirrorIndents",
  "suppressOverlap",
  "jc",
  "textDirection",
  "textAlignment",
  "textboxTightWrap",
  "outlineLvl",
  "divId",
  "cnfStyle",
  "rPr",
  "sectPr",
  "pPrChange",
] as const;

/** Schema order of `<w:rPr>` children (§17.3.2.28). */
export const RPR_ORDER = [
  "rStyle",
  "rFonts",
  "b",
  "bCs",
  "i",
  "iCs",
  "caps",
  "smallCaps",
  "strike",
  "dstrike",
  "outline",
  "shadow",
  "emboss",
  "imprint",
  "noProof",
  "snapToGrid",
  "vanish",
  "webHidden",
  "color",
  "spacing",
  "w",
  "kern",
  "position",
  "sz",
  "szCs",
  "highlight",
  "u",
  "effect",
  "bdr",
  "shd",
  "fitText",
  "vertAlign",
  "rtl",
  "cs",
  "em",
  "lang",
  "eastAsianLayout",
  "specVanish",
  "oMath",
] as const;

/** Schema order of `<w:sectPr>` children (§17.6.17). */
export const SECTPR_ORDER = [
  "headerReference",
  "footerReference",
  "footnotePr",
  "endnotePr",
  "type",
  "pgSz",
  "pgMar",
  "paperSrc",
  "pgBorders",
  "lnNumType",
  "pgNumType",
  "cols",
  "formProt",
  "vAlign",
  "noEndnote",
  "titlePg",
  "textDirection",
  "bidi",
  "rtlGutter",
  "docGrid",
  "printerSettings",
  "sectPrChange",
] as const;

/**
 * Replace (or with `undefined`, remove) the `w:` child named `child.name.local`
 * of `container`, inserting a new one at its schema position.
 */
export function setOrderedChild(
  container: XmlElement,
  local: string,
  child: XmlElement | undefined,
  order: readonly string[],
): XmlElement {
  const kept = container.children.filter((c) => !isW(c, local));
  if (!child) return { ...container, children: kept, selfClosing: false };
  const rank = order.indexOf(local);
  let at = kept.length;
  for (const [i, c] of kept.entries()) {
    if (c.kind !== "element" || c.name.uri !== WML_NS) continue;
    const r = order.indexOf(c.name.local);
    if (r > rank) {
      at = i;
      break;
    }
  }
  const children = [...kept.slice(0, at), child, ...kept.slice(at)];
  return { ...container, children, selfClosing: false };
}

export function emptyProps(local: string): XmlElement {
  return { ...wEl(local), selfClosing: false };
}

/** Set a child of a paragraph's `<w:pPr>` (created when missing). */
export function setParagraphChild(
  para: WmlParagraph,
  local: string,
  child: XmlElement | undefined,
): void {
  const pPr = para.pPr ?? emptyProps("pPr");
  para.pPr = setOrderedChild(pPr, local, child, PPR_ORDER);
}

/** Set a child of a run's `<w:rPr>` (created when missing). */
export function setRunChild(run: WmlRun, local: string, child: XmlElement | undefined): void {
  const rPr = run.rPr ?? emptyProps("rPr");
  run.rPr = setOrderedChild(rPr, local, child, RPR_ORDER);
}

/**
 * Turn `<w:r>` elements a builder left as raw inlines (e.g. `appendField`)
 * into typed runs so run-level code (offsets, field walking) sees them.
 */
export function normalizeRawRuns(para: WmlParagraph): void {
  if (!para.children.some((c) => c.kind === "raw" && isW(c.node, "r"))) return;
  para.children = para.children.map((child): WmlInline => {
    if (child.kind !== "raw" || !isW(child.node, "r")) return child;
    const parsed = parseParagraph(wEl("p", {}, [child.node])).children[0];
    return parsed ?? child;
  });
}

/**
 * Where content inserted at paragraph-absolute character `offset` goes in
 * `para.children`: the run containing the offset is split (when it is plain
 * text) and the index after the last run that ends at or before the offset is
 * returned. A field's closing `fldChar` sitting right at the offset stays
 * before the insertion point, so content lands after the field rather than in
 * its result.
 */
export function childIndexAtOffset(para: WmlParagraph, offset: number): number {
  normalizeRawRuns(para);
  if (offset > 0) isolateParagraphRunRange(para, 0, offset);
  let cursor = 0;
  let index = 0;
  for (const [i, child] of para.children.entries()) {
    if (child.kind !== "run") continue;
    const start = cursor;
    const end = cursor + runTextLength(child);
    cursor = end;
    const before = end < offset || (end === offset && start < offset);
    // A run that could not be split at the offset (tabs, fields) stays whole.
    const straddles = start < offset && end > offset;
    const closesAtOffset = end === start && start === offset && closesField(child);
    if (!before && !straddles && !closesAtOffset) break;
    index = i + 1;
  }
  return index;
}

function closesField(run: WmlRun): boolean {
  return run.pieces.some(
    (p) => p.kind === "fieldChar" && (p.charType === "end" || p.charType === "separate"),
  );
}

/** Splice inlines at a character offset; returns the index of the first one. */
export function spliceInlinesAt(
  para: WmlParagraph,
  offset: number,
  inlines: readonly WmlInline[],
): number {
  const at = childIndexAtOffset(para, offset);
  para.children.splice(at, 0, ...inlines);
  return at;
}

/** Throw a RangeError unless `offset` is a valid caret offset in `para`. */
export function assertOffset(para: WmlParagraph, offset: number): void {
  let length = 0;
  for (const c of para.children) if (c.kind === "run") length += runTextLength(c);
  if (!Number.isInteger(offset) || offset < 0 || offset > length) {
    throw new RangeError(`Offset ${offset} is outside the paragraph (0–${length}).`);
  }
}

export function xmlSpaceAttr(): XmlAttr {
  return {
    name: { uri: XML_NAMESPACE, local: "space", prefix: "xml" },
    value: "preserve",
    isNamespaceDecl: false,
  };
}

/** Visible text inside an XML subtree (`w:t` only; field codes excluded). */
export function xmlVisibleText(node: XmlElement): string {
  if (isW(node, "t")) {
    return node.children
      .map((c) => (c.kind === "text" || c.kind === "cdata" ? c.value : ""))
      .join("");
  }
  if (isW(node, "tab")) return "\t";
  let out = "";
  for (const c of node.children) if (c.kind === "element") out += xmlVisibleText(c);
  return out;
}

/** Visible text of a run (field codes and deleted text excluded). */
export function runVisibleText(run: WmlRun): string {
  let out = "";
  for (const p of run.pieces) {
    if (p.kind === "text") out += p.value;
    else if (p.kind === "tab") out += "\t";
    else if (p.kind === "noBreakHyphen") out += "-";
  }
  return out;
}

/** Every paragraph of the body in document order, table cells included. */
export function bodyParagraphs(doc: Docx): WmlParagraph[] {
  const out: WmlParagraph[] = [];
  for (const block of doc.document.body.blocks) {
    if (block.kind === "paragraph") out.push(block);
    else if (block.kind === "table") {
      for (const row of block.rows) for (const cell of row.cells) out.push(...cell.paragraphs);
    }
  }
  return out;
}

/** Where a paragraph lives: its index in the body or in a table cell. */
export interface ParagraphSlot {
  readonly index: number;
  /** The paragraph `delta` positions away in the same container, if any. */
  sibling(delta: number): WmlParagraph | undefined;
  insert(at: number, ...paras: WmlParagraph[]): void;
  remove(at: number, count: number): void;
}

/** Locate `para` in the body or in one of its tables' cells. */
export function paragraphSlot(doc: Docx, para: WmlParagraph): ParagraphSlot {
  const blocks = doc.document.body.blocks;
  const top = blocks.indexOf(para);
  if (top >= 0) {
    return {
      index: top,
      sibling: (delta) => {
        const b = blocks[top + delta];
        return b?.kind === "paragraph" ? b : undefined;
      },
      insert: (at, ...paras) => blocks.splice(at, 0, ...paras),
      remove: (at, count) => blocks.splice(at, count),
    };
  }
  for (const block of blocks) {
    if (block.kind !== "table") continue;
    for (const row of block.rows) {
      for (const cell of row.cells) {
        const list = cell.paragraphs;
        const i = list.indexOf(para);
        if (i < 0) continue;
        return {
          index: i,
          sibling: (delta) => list[i + delta],
          insert: (at, ...paras) => list.splice(at, 0, ...paras),
          remove: (at, count) => list.splice(at, count),
        };
      }
    }
  }
  throw new Error("The paragraph is not part of this document's body.");
}
