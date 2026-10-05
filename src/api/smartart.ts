/**
 * SmartArt graphics (ECMA-376 Part 1, §21.4 DrawingML diagrams).
 *
 * A diagram is an inline `w:drawing` whose graphic data (`a:graphicData`
 * with the diagram URI) points at four parts: the data model (the bullet
 * tree the user edits in the Text Pane), the layout definition, the style
 * definition (SmartArt Styles) and the colour definition (Change Colors).
 * Word lays the diagram out from the data model and the layout definition
 * when it opens the document.
 */

import {
  addPart,
  addRelationship,
  getPart,
  hasPart,
  partRelationships,
  relationshipById,
} from "../internal/opc/index.js";
import { REL_NS, type WmlParagraph, type WmlRun } from "../internal/wordprocessingml/index.js";
import { parseXml, type XmlElement } from "../internal/xml/index.js";
import {
  colorsDefinitionXml,
  colorsUniqueId,
  layoutDefinitionXml,
  layoutUniqueId,
  SMARTART_COLORS,
  SMARTART_LAYOUTS,
  SMARTART_STYLES,
  type SmartArtColors,
  type SmartArtLayout,
  type SmartArtStyle,
  styleDefinitionXml,
  styleUniqueId,
} from "../internal/dgm/definitions.js";
import type { Docx } from "./docx.js";

export {
  SMARTART_COLORS,
  SMARTART_LAYOUTS,
  SMARTART_STYLES,
  type SmartArtColors,
  type SmartArtLayout,
  type SmartArtStyle,
} from "../internal/dgm/definitions.js";

const DGM_NS = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const REL_TYPES = {
  data: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData",
  layout: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout",
  style: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle",
  colors: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors",
} as const;
const CONTENT_TYPES = {
  data: "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml",
  layout: "application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml",
  style: "application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml",
  colors: "application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml",
} as const;
const PART_PREFIX = {
  data: "data",
  layout: "layout",
  style: "quickStyle",
  colors: "colors",
} as const;
type PartKind = keyof typeof REL_TYPES;
const PART_KINDS: readonly PartKind[] = ["data", "layout", "style", "colors"];
// The relationship attribute on `dgm:relIds` for each part (§21.4.3.x).
const REL_ATTR: Readonly<Record<PartKind, string>> = {
  data: "dm",
  layout: "lo",
  style: "qs",
  colors: "cs",
};

const EMU_PER_POINT = 12700;
const HEADER = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n`;

/** One bullet of the Text Pane: a shape's text and its sub-bullets. */
export interface SmartArtNode {
  readonly text: string;
  readonly children?: readonly SmartArtNode[];
}

export interface AddSmartArtOptions {
  readonly layout: SmartArtLayout;
  readonly nodes: readonly SmartArtNode[];
  /** Points. Word's default SmartArt frame is 6" × 3.5". */
  readonly width?: number;
  readonly height?: number;
  readonly colors?: SmartArtColors;
  readonly style?: SmartArtStyle;
  readonly altText?: string;
}

/** A SmartArt graphic in the document: its drawing and the part names it uses. */
export interface SmartArtRef {
  readonly drawing: XmlElement;
  readonly parts: Readonly<Record<PartKind, string>>;
}

export interface SmartArtInfo {
  readonly layout: SmartArtLayout | undefined;
  readonly colors: SmartArtColors | undefined;
  readonly style: SmartArtStyle | undefined;
  readonly nodes: SmartArtNode[];
  /** Frame size in points. */
  readonly width: number;
  readonly height: number;
}

const DEFAULT_WIDTH = 432;
const DEFAULT_HEIGHT = 252;

function escapeXml(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

/**
 * Model ids are GUIDs in braces (`ST_ModelId`); sequential ones keep the
 * output deterministic while staying unique within the diagram.
 */
function modelId(n: number): string {
  return `{${n.toString(16).toUpperCase().padStart(8, "0")}-0000-4000-8000-000000000000}`;
}

function textBody(text: string): string {
  const run =
    text === ""
      ? `<a:endParaRPr lang="en-US"/>`
      : `<a:r><a:rPr lang="en-US"/><a:t>${escapeXml(text)}</a:t></a:r>`;
  return `<dgm:t><a:bodyPr/><a:lstStyle/><a:p>${run}</a:p></dgm:t>`;
}

function dataModelXml(
  nodes: readonly SmartArtNode[],
  layout: SmartArtLayout,
  colors: SmartArtColors,
  style: SmartArtStyle,
): string {
  let next = 1;
  const docId = modelId(next++);
  const points: string[] = [
    `<dgm:pt modelId="${docId}" type="doc"><dgm:prSet loTypeId="${layoutUniqueId(layout)}" loCatId="${SMARTART_LAYOUTS[layout].category}" ` +
      `qsTypeId="${styleUniqueId(style)}" qsCatId="simple" csTypeId="${colorsUniqueId(colors)}" csCatId="${SMARTART_COLORS[colors].category}"/>` +
      `<dgm:spPr/>${textBody("")}</dgm:pt>`,
  ];
  const cxns: string[] = [];
  const walk = (parent: string, list: readonly SmartArtNode[]): void => {
    list.forEach((node, ord) => {
      const id = modelId(next++);
      const cxn = modelId(next++);
      const par = modelId(next++);
      const sib = modelId(next++);
      points.push(
        `<dgm:pt modelId="${id}"><dgm:prSet phldrT="[Text]"/><dgm:spPr/>${textBody(node.text)}</dgm:pt>`,
        `<dgm:pt modelId="${par}" type="parTrans" cxnId="${cxn}"><dgm:prSet/><dgm:spPr/>${textBody("")}</dgm:pt>`,
        `<dgm:pt modelId="${sib}" type="sibTrans" cxnId="${cxn}"><dgm:prSet/><dgm:spPr/>${textBody("")}</dgm:pt>`,
      );
      cxns.push(
        `<dgm:cxn modelId="${cxn}" srcId="${parent}" destId="${id}" srcOrd="${ord}" destOrd="0" parTransId="${par}" sibTransId="${sib}"/>`,
      );
      walk(id, node.children ?? []);
    });
  };
  walk(docId, nodes);
  return (
    HEADER +
    `<dgm:dataModel xmlns:dgm="${DGM_NS}" xmlns:a="${A_NS}">` +
    `<dgm:ptLst>${points.join("")}</dgm:ptLst><dgm:cxnLst>${cxns.join("")}</dgm:cxnLst>` +
    `<dgm:bg/><dgm:whole/></dgm:dataModel>`
  );
}

function allocatePart(doc: Docx, kind: PartKind): string {
  let n = 1;
  while (hasPart(doc.opc, `/word/diagrams/${PART_PREFIX[kind]}${n}.xml`)) n++;
  return `/word/diagrams/${PART_PREFIX[kind]}${n}.xml`;
}

const encode = (xml: string): Uint8Array => new TextEncoder().encode(xml);

function nextDocPrId(doc: Docx): number {
  let max = 0;
  const visit = (el: XmlElement): void => {
    if (el.name.uri === WP_NS && el.name.local === "docPr") {
      const id = Number(el.attrs.find((a) => a.name.local === "id")?.value);
      if (Number.isFinite(id)) max = Math.max(max, id);
    }
    for (const c of el.children) if (c.kind === "element") visit(c);
  };
  for (const block of doc.document.body.blocks) {
    if (block.kind !== "paragraph") continue;
    for (const inline of block.children) {
      if (inline.kind !== "run") continue;
      for (const piece of inline.pieces) if (piece.kind === "drawing") visit(piece.node);
    }
  }
  return max + 1;
}

/**
 * Insert a SmartArt graphic (Insert ▸ SmartArt) as an inline drawing at the
 * end of `paragraph` (or at inline index `index`). Returns the new run.
 */
export function addSmartArt(
  doc: Docx,
  paragraph: WmlParagraph,
  options: AddSmartArtOptions,
  index: number = paragraph.children.length,
): WmlRun {
  const colors = options.colors ?? "accent1_2";
  const style = options.style ?? "simple1";
  const xml: Record<PartKind, string> = {
    data: dataModelXml(options.nodes, options.layout, colors, style),
    layout: layoutDefinitionXml(options.layout),
    style: styleDefinitionXml(style),
    colors: colorsDefinitionXml(colors),
  };
  const rels = partRelationships(doc.opc, doc.partName);
  const relIds = {} as Record<PartKind, string>;
  for (const kind of PART_KINDS) {
    const name = allocatePart(doc, kind);
    addPart(doc.opc, { name, contentType: CONTENT_TYPES[kind], data: encode(xml[kind]) });
    relIds[kind] = addRelationship(rels, {
      type: REL_TYPES[kind],
      target: name.slice("/word/".length),
    }).id;
  }
  const cx = Math.round((options.width ?? DEFAULT_WIDTH) * EMU_PER_POINT);
  const cy = Math.round((options.height ?? DEFAULT_HEIGHT) * EMU_PER_POINT);
  const id = nextDocPrId(doc);
  const descr = options.altText ? ` descr="${escapeXml(options.altText)}"` : "";
  const drawing = parseXml(
    `<w:drawing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="${WP_NS}" ` +
      `xmlns:a="${A_NS}" xmlns:dgm="${DGM_NS}" xmlns:r="${REL_NS}">` +
      `<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/>` +
      `<wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="${id}" name="Diagram ${id}"${descr}/>` +
      `<wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="${DGM_NS}">` +
      `<dgm:relIds r:dm="${relIds.data}" r:lo="${relIds.layout}" r:qs="${relIds.style}" r:cs="${relIds.colors}"/>` +
      `</a:graphicData></a:graphic></wp:inline></w:drawing>`,
  ).root;
  const run: WmlRun = { kind: "run", pieces: [{ kind: "drawing", node: drawing }], extras: [] };
  paragraph.children.splice(index, 0, run);
  doc.dirty = true;
  return run;
}

function findRelIds(el: XmlElement): XmlElement | undefined {
  if (el.name.uri === DGM_NS && el.name.local === "relIds") return el;
  for (const c of el.children) {
    if (c.kind !== "element") continue;
    const found = findRelIds(c);
    if (found) return found;
  }
  return undefined;
}

/** The SmartArt graphic a run holds, if any. */
export function runSmartArt(doc: Docx, run: WmlRun): SmartArtRef | undefined {
  for (const piece of run.pieces) {
    if (piece.kind !== "drawing") continue;
    const relIdsEl = findRelIds(piece.node);
    if (!relIdsEl) continue;
    const rels = partRelationships(doc.opc, doc.partName);
    const parts = {} as Record<PartKind, string>;
    for (const kind of PART_KINDS) {
      const relId = relIdsEl.attrs.find(
        (a) => a.name.uri === REL_NS && a.name.local === REL_ATTR[kind],
      )?.value;
      const rel = relId === undefined ? undefined : relationshipById(rels, relId);
      if (!rel) return undefined;
      parts[kind] = `/word/${rel.target.replace(/^\.\//, "")}`;
    }
    return { drawing: piece.node, parts };
  }
  return undefined;
}

/** Every SmartArt graphic in the body, in document order. */
export function smartArts(doc: Docx): SmartArtRef[] {
  const out: SmartArtRef[] = [];
  for (const block of doc.document.body.blocks) {
    if (block.kind !== "paragraph") continue;
    for (const inline of block.children) {
      if (inline.kind !== "run") continue;
      const ref = runSmartArt(doc, inline);
      if (ref) out.push(ref);
    }
  }
  return out;
}

function readPart(doc: Docx, name: string): XmlElement | undefined {
  const part = getPart(doc.opc, name);
  return part ? parseXml(new TextDecoder("utf-8").decode(part.data)).root : undefined;
}

/**
 * Rewrite a definition part. The raw-XML inspector may hold a parsed copy of
 * it; that copy is dropped so it cannot overwrite the new bytes on save.
 */
function writePart(doc: Docx, name: string, xml: string): void {
  const part = getPart(doc.opc, name);
  if (!part) throw new Error(`SmartArt part ${name} is missing`);
  part.data = encode(xml);
  doc.rawParts.delete(name);
  doc.rawPartsDirty.delete(name);
  doc.dirty = true;
}

function dgmChildren(el: XmlElement | undefined, local: string): XmlElement[] {
  if (!el) return [];
  return el.children.filter(
    (c): c is XmlElement => c.kind === "element" && c.name.uri === DGM_NS && c.name.local === local,
  );
}

const attr = (el: XmlElement, local: string): string | undefined =>
  el.attrs.find((a) => a.name.local === local && !a.isNamespaceDecl)?.value;

function plainText(el: XmlElement): string {
  // One line per a:p; runs concatenated.
  const paragraphs: string[] = [];
  const collect = (node: XmlElement, into: string[]): void => {
    for (const c of node.children) {
      if (c.kind !== "element") continue;
      if (c.name.uri === A_NS && c.name.local === "t") {
        into.push(c.children.map((t) => (t.kind === "text" ? t.value : "")).join(""));
      } else collect(c, into);
    }
  };
  for (const p of el.children) {
    if (p.kind !== "element" || p.name.local !== "p") continue;
    const parts: string[] = [];
    collect(p, parts);
    paragraphs.push(parts.join(""));
  }
  return paragraphs.join("\n");
}

function lookup<K extends string>(
  table: Readonly<Record<K, unknown>>,
  find: (key: K) => boolean,
): K | undefined {
  return (Object.keys(table) as K[]).find(find);
}

/** Read a diagram's layout, colours, style, Text Pane bullets and frame size. */
export function getSmartArt(doc: Docx, ref: SmartArtRef): SmartArtInfo {
  const data = readPart(doc, ref.parts.data);
  const pts = dgmChildren(dgmChildren(data, "ptLst")[0], "pt");
  const cxns = dgmChildren(dgmChildren(data, "cxnLst")[0], "cxn");
  const byId = new Map(pts.map((p) => [attr(p, "modelId") ?? "", p]));
  const childrenOf = new Map<string, Array<{ ord: number; id: string }>>();
  for (const cxn of cxns) {
    const type = attr(cxn, "type") ?? "parOf";
    if (type !== "parOf") continue;
    const src = attr(cxn, "srcId") ?? "";
    const list = childrenOf.get(src) ?? [];
    list.push({ ord: Number(attr(cxn, "srcOrd") ?? 0), id: attr(cxn, "destId") ?? "" });
    childrenOf.set(src, list);
  }
  const docPt = pts.find((p) => attr(p, "type") === "doc");
  const build = (parent: string): SmartArtNode[] =>
    (childrenOf.get(parent) ?? [])
      .toSorted((a, b) => a.ord - b.ord)
      .flatMap(({ id }) => {
        const pt = byId.get(id);
        const type = pt && (attr(pt, "type") ?? "node");
        if (!pt || type !== "node") return [];
        const t = dgmChildren(pt, "t")[0];
        const kids = build(id);
        return [{ text: t ? plainText(t) : "", ...(kids.length ? { children: kids } : {}) }];
      });
  const uniqueIdOf = (name: string): string => {
    const root = readPart(doc, name);
    return root ? (attr(root, "uniqueId") ?? "") : "";
  };
  const layoutId = uniqueIdOf(ref.parts.layout);
  const colorsId = uniqueIdOf(ref.parts.colors);
  const styleId = uniqueIdOf(ref.parts.style);
  const extent = findExtent(ref.drawing);
  return {
    layout: lookup(SMARTART_LAYOUTS, (k) => layoutUniqueId(k) === layoutId),
    colors: lookup(SMARTART_COLORS, (k) => colorsUniqueId(k) === colorsId),
    style: lookup(SMARTART_STYLES, (k) => styleUniqueId(k) === styleId),
    nodes: docPt ? build(attr(docPt, "modelId") ?? "") : [],
    width: extent.cx / EMU_PER_POINT,
    height: extent.cy / EMU_PER_POINT,
  };
}

function findExtent(el: XmlElement): { cx: number; cy: number } {
  if (el.name.uri === WP_NS && el.name.local === "extent") {
    return { cx: Number(attr(el, "cx") ?? 0), cy: Number(attr(el, "cy") ?? 0) };
  }
  for (const c of el.children) {
    if (c.kind !== "element") continue;
    const found = findExtent(c);
    if (found.cx || found.cy) return found;
  }
  return { cx: 0, cy: 0 };
}

/**
 * Text Pane: replace the diagram's bullets. The data model is rewritten, so
 * per-shape formatting a loaded diagram carried is not kept.
 */
export function setSmartArtNodes(
  doc: Docx,
  ref: SmartArtRef,
  nodes: readonly SmartArtNode[],
): void {
  const info = getSmartArt(doc, ref);
  writePart(
    doc,
    ref.parts.data,
    dataModelXml(
      nodes,
      info.layout ?? "basicBlockList",
      info.colors ?? "accent1_2",
      info.style ?? "simple1",
    ),
  );
}

/** SmartArt Design ▸ Layouts. */
export function setSmartArtLayout(doc: Docx, ref: SmartArtRef, layout: SmartArtLayout): void {
  const info = getSmartArt(doc, ref);
  writePart(doc, ref.parts.layout, layoutDefinitionXml(layout));
  writePart(
    doc,
    ref.parts.data,
    dataModelXml(info.nodes, layout, info.colors ?? "accent1_2", info.style ?? "simple1"),
  );
}

/** SmartArt Design ▸ Change Colors. */
export function setSmartArtColors(doc: Docx, ref: SmartArtRef, colors: SmartArtColors): void {
  const info = getSmartArt(doc, ref);
  writePart(doc, ref.parts.colors, colorsDefinitionXml(colors));
  writePart(
    doc,
    ref.parts.data,
    dataModelXml(info.nodes, info.layout ?? "basicBlockList", colors, info.style ?? "simple1"),
  );
}

/** SmartArt Design ▸ SmartArt Styles. */
export function setSmartArtStyle(doc: Docx, ref: SmartArtRef, style: SmartArtStyle): void {
  const info = getSmartArt(doc, ref);
  writePart(doc, ref.parts.style, styleDefinitionXml(style));
  writePart(
    doc,
    ref.parts.data,
    dataModelXml(info.nodes, info.layout ?? "basicBlockList", info.colors ?? "accent1_2", style),
  );
}

/** Resize the diagram's frame (points). */
export function setSmartArtSize(ref: SmartArtRef, width: number, height: number): void {
  const visit = (el: XmlElement): void => {
    if (el.name.uri === WP_NS && el.name.local === "extent") {
      const attrs = el.attrs as XmlElement["attrs"][number][];
      for (let i = 0; i < attrs.length; i++) {
        const a = attrs[i];
        if (a?.name.local === "cx")
          attrs[i] = { ...a, value: String(Math.round(width * EMU_PER_POINT)) };
        if (a?.name.local === "cy")
          attrs[i] = { ...a, value: String(Math.round(height * EMU_PER_POINT)) };
      }
    }
    for (const c of el.children) if (c.kind === "element") visit(c);
  };
  visit(ref.drawing);
}
