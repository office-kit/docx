import type { XmlAttr, XmlElement, XmlNode } from "../xml/index.js";
import { WML_NS } from "./namespaces.js";
import { revisionElement, wmlAttrValue } from "./revisions.js";
import type { WmlBlock, WmlParagraph, WmlRunPiece } from "./types.js";

/** A piece as it reads inside a deletion: text becomes `delText`, a field code `delInstrText`. */
export function toDeletedPiece(piece: WmlRunPiece): WmlRunPiece {
  if (piece.kind === "text") return { ...piece, kind: "delText" };
  if (piece.kind === "instrText") return { ...piece, kind: "delInstrText" };
  return piece;
}

/** ISO 8601 in whole seconds, as Word writes `w:date`; now when `date` is omitted. */
export function revisionDate(date: string | undefined): string {
  return date ?? new Date().toISOString().replace(/\.\d{3}Z$/, "Z");
}

function isWml(node: XmlNode, local: string): node is XmlElement {
  return node.kind === "element" && node.name.uri === WML_NS && node.name.local === local;
}

function collectIds(node: XmlNode, ids: number[]): void {
  if (node.kind !== "element") return;
  const id = wmlAttrValue(node, "id");
  if (id !== undefined && /^\d+$/.test(id)) ids.push(Number(id));
  for (const child of node.children) collectIds(child, ids);
}

function paragraphsOf(blocks: readonly WmlBlock[]): WmlParagraph[] {
  return blocks.flatMap((block) =>
    block.kind === "paragraph"
      ? [block]
      : block.kind === "table"
        ? block.rows.flatMap((row) => row.cells.flatMap((cell) => cell.paragraphs))
        : [],
  );
}

/**
 * A `w:id` for a new annotation: one past every id in use by revisions,
 * bookmarks, comment anchors (and `extra`, e.g. the comments part), so ids
 * stay unique across annotation types (§17.13.1).
 */
export function nextAnnotationId(
  blocks: readonly WmlBlock[],
  sectPr: XmlElement | undefined,
  extra: readonly XmlElement[] = [],
): number {
  const ids: number[] = [-1];
  for (const para of paragraphsOf(blocks)) {
    if (para.pPr) collectIds(para.pPr, ids);
    for (const child of para.children) {
      if (child.kind === "raw") collectIds(child.node, ids);
      else {
        if (child.rPr) collectIds(child.rPr, ids);
        const id = child.revision && wmlAttrValue(child.revision, "id");
        if (id !== undefined && /^\d+$/.test(id)) ids.push(Number(id));
      }
    }
  }
  for (const block of blocks) {
    if (block.kind !== "table") continue;
    if (block.tblPr) collectIds(block.tblPr, ids);
    for (const row of block.rows) {
      if (row.trPr) collectIds(row.trPr, ids);
      for (const cell of row.cells) if (cell.tcPr) collectIds(cell.tcPr, ids);
    }
  }
  if (sectPr) collectIds(sectPr, ids);
  for (const el of extra) collectIds(el, ids);
  return Math.max(...ids) + 1;
}

/**
 * Record `<w:ins>` / `<w:del>` on a paragraph's mark: first in the mark's
 * `<w:rPr>` (CT_ParaRPr), `ins` before `del`; the rPr goes before `<w:sectPr>`
 * / `<w:pPrChange>` in the pPr (CT_PPr order).
 */
export function setParagraphMarkRevision(
  para: WmlParagraph,
  kind: "ins" | "del",
  attrs: readonly XmlAttr[],
): void {
  para.pPr ??= revisionElement("pPr", []);
  const pPrChildren = para.pPr.children as XmlNode[];
  let rPr = pPrChildren.find((c): c is XmlElement => isWml(c, "rPr"));
  if (!rPr) {
    rPr = revisionElement("rPr", []);
    const at = pPrChildren.findIndex((c) => isWml(c, "sectPr") || isWml(c, "pPrChange"));
    pPrChildren.splice(at < 0 ? pPrChildren.length : at, 0, rPr);
  }
  const list = rPr.children as XmlNode[];
  const existing = list.findIndex((c) => isWml(c, kind));
  if (existing >= 0) list.splice(existing, 1);
  const insAt = list.findIndex((c) => isWml(c, "ins"));
  list.splice(kind === "del" && insAt >= 0 ? insAt + 1 : 0, 0, revisionElement(kind, attrs));
}
