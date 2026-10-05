/**
 * Pictures, charts and their arrangement: everything Word's Picture Format
 * tab, Arrange group and Insert ▸ Chart write into a `<w:drawing>`.
 *
 * Drawings are addressed by their index in {@link imageDrawings} order (body
 * paragraphs, then table cells, in document order). Every setter writes only
 * ECMA-376 markup: no `a14:` / `wp14:` / `wpg:` / `cx:` extensions.
 */

import {
  addPart,
  addRelationship,
  getPart,
  hasPart,
  partRelationships,
  relationshipById,
  relationshipsByType,
  resolveInternalTarget,
  setContentTypeDefault,
} from "../internal/opc/index.js";
import {
  type ChartSpec,
  chartSheetRows,
  chartSpaceXml,
  readChartSpace,
  validateChartSpec,
} from "../internal/drawingml/chart.js";
import { readImagePixelSize, type ImagePixelSize } from "../internal/drawingml/image-size.js";
import {
  type AnchorOptions,
  applyAnchorOptions,
  applyPosition,
  applyWrap,
  BASE_RELATIVE_HEIGHT,
  maxDocPrId,
  placementOf,
  readAnchor,
  readWrap,
} from "../internal/drawingml/layout.js";
import {
  applyAdjustments,
  applyCrop,
  applyEffects,
  applyGeometry,
  applyOutline,
  applyTransform,
  blipOf,
  picOf,
  readAdjustments,
  readCrop,
  readEffects,
  readGeometry,
  readOutline,
  readTransform,
  setBlipRelationship,
} from "../internal/drawingml/picture.js";
import type {
  DrawingInfo,
  HorizontalPosition,
  PictureColorAdjustments,
  PictureCrop,
  PictureEffects,
  PictureImage,
  PictureInfo,
  PictureOutline,
  VerticalPosition,
  WrapStyle,
} from "../internal/drawingml/types.js";
import { writeWorkbook } from "../internal/drawingml/xlsx.js";
import {
  A_NS,
  attr,
  boolAttr,
  C_NS,
  child,
  descendant,
  DRAWING_PREFIXES,
  ensureDrawingNamespaces,
  escapeXml,
  fragment,
  nsAttr,
  numAttr,
  PIC_NS,
  placeChild,
  R_NS,
  setAttr,
  WP_NS,
} from "../internal/drawingml/xml.js";
import {
  extensionForImageContentType,
  sniffImageContentType,
  WML_NS,
  WML_RELATIONSHIPS,
  type WmlParagraph,
  type WmlRun,
} from "../internal/wordprocessingml/index.js";
import { parseXml, type XmlElement, type XmlNode } from "../internal/xml/index.js";
import { type Docx, imageDrawings, setImageSizeEmu } from "./docx.js";

export type {
  AnchorOptions,
  ChartSpec,
  DrawingInfo,
  HorizontalPosition,
  ImagePixelSize,
  PictureColorAdjustments,
  PictureCrop,
  PictureEffects,
  PictureImage,
  PictureInfo,
  PictureOutline,
  VerticalPosition,
  WrapStyle,
};
export type {
  BevelPreset,
  CompoundLine,
  DrawingAnchor,
  DrawingKind,
  HorizontalAlign,
  HorizontalRelativeTo,
  LineDash,
  Picture3dRotation,
  PictureBevel,
  PictureGlow,
  PictureReflection,
  PictureShadow,
  RectAlignment,
  VerticalAlign,
  VerticalRelativeTo,
  WrapDistance,
  WrapPoint,
  WrapSide,
} from "../internal/drawingml/types.js";
export type {
  ChartGrouping,
  ChartKind,
  ChartSeries,
  LegendPosition,
} from "../internal/drawingml/chart.js";

const CHART_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
const PACKAGE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package";
const CHART_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
const XLSX_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const EMU_PER_INCH = 914400;
const DOC_PR_ORDER = ["hlinkClick", "hlinkHover", "extLst"] as const;

// --- locating ----------------------------------------------------------------

function drawingAt(doc: Docx, index: number): XmlElement {
  const drawing = imageDrawings(doc)[index];
  if (!drawing) throw new Error(`No drawing at index ${index}.`);
  return drawing;
}

function pictureAt(doc: Docx, index: number): { drawing: XmlElement; pic: XmlElement } {
  const drawing = drawingAt(doc, index);
  const pic = picOf(drawing);
  if (!pic) throw new Error(`The drawing at index ${index} is not a picture.`);
  return { drawing, pic };
}

function edit(doc: Docx): void {
  doc.dirty = true;
}

/** Resolve a relationship target of the main document part to a part name. */
function documentTarget(doc: Docx, relId: string): string | undefined {
  const rel = relationshipById(partRelationships(doc.opc, doc.partName), relId);
  if (!rel || rel.targetMode === "External") return undefined;
  return resolveInternalTarget(doc.partName, rel.target);
}

function imageOf(doc: Docx, pic: XmlElement): PictureImage | undefined {
  const relId = nsAttr(blipOf(pic), R_NS, "embed");
  const partName = relId ? documentTarget(doc, relId) : undefined;
  const part = partName ? getPart(doc.opc, partName) : undefined;
  return part ? { partName: part.name, contentType: part.contentType, data: part.data } : undefined;
}

function chartPartOf(doc: Docx, drawing: XmlElement): string | undefined {
  const ref = descendant(drawing, C_NS, "chart");
  const relId = nsAttr(ref, R_NS, "id");
  return relId ? documentTarget(doc, relId) : undefined;
}

// --- reading -------------------------------------------------------------------

/**
 * Read a `<w:drawing>` (from {@link imageDrawings}) into plain data: kind,
 * size, docPr, wrapping and position, and — for pictures — crop, shape,
 * border, effects and color adjustments.
 */
export function readDrawing(doc: Docx, drawing: XmlElement): DrawingInfo {
  const placement = placementOf(drawing);
  const extent = child(placement, WP_NS, "extent");
  const docPr = child(placement, WP_NS, "docPr");
  const pic = picOf(drawing);
  const chartPartName = chartPartOf(doc, drawing);
  const kind = pic ? "picture" : chartPartName ? "chart" : "other";
  const hlinkId = nsAttr(child(docPr, A_NS, "hlinkClick"), R_NS, "id");
  const hlink = hlinkId
    ? relationshipById(partRelationships(doc.opc, doc.partName), hlinkId)
    : undefined;
  const locks =
    descendant(drawing, A_NS, "graphicFrameLocks") ?? descendant(drawing, A_NS, "picLocks");
  const anchor = readAnchor(drawing);
  return {
    kind,
    id: numAttr(docPr, "id") ?? 0,
    name: attr(docPr, "name") ?? "",
    description: attr(docPr, "descr") ?? "",
    title: attr(docPr, "title") ?? "",
    hidden: boolAttr(docPr, "hidden") ?? false,
    widthEmu: numAttr(extent, "cx") ?? 0,
    heightEmu: numAttr(extent, "cy") ?? 0,
    lockAspect: boolAttr(locks, "noChangeAspect") ?? false,
    ...(hlink?.targetMode === "External" ? { hyperlink: hlink.target } : {}),
    wrap: readWrap(drawing),
    ...(anchor ? { anchor } : {}),
    ...readTransform(pic),
    ...(pic ? { picture: readPicture(doc, pic) } : {}),
    ...(chartPartName ? { chartPartName } : {}),
  };
}

function readPicture(doc: Docx, pic: XmlElement): PictureInfo {
  const image = imageOf(doc, pic);
  const outline = readOutline(pic);
  return {
    ...(image ? { image } : {}),
    crop: readCrop(pic),
    geometry: readGeometry(pic),
    ...(outline ? { outline } : {}),
    effects: readEffects(pic),
    adjustments: readAdjustments(pic),
  };
}

/** The pixel size and resolution of PNG / JPEG / GIF / BMP bytes. */
export function imagePixelSize(bytes: Uint8Array): ImagePixelSize | undefined {
  return readImagePixelSize(bytes);
}

/** An image's size at its own resolution, in EMU (what Word inserts at 100 %). */
export function imageNaturalSizeEmu(
  bytes: Uint8Array,
): { widthEmu: number; heightEmu: number } | undefined {
  const size = readImagePixelSize(bytes);
  if (!size) return undefined;
  return {
    widthEmu: Math.round((size.width / size.dpiX) * EMU_PER_INCH),
    heightEmu: Math.round((size.height / size.dpiY) * EMU_PER_INCH),
  };
}

// --- arrangement -------------------------------------------------------------

function floatingHeights(doc: Docx): Array<{ drawing: XmlElement; height: number }> {
  return imageDrawings(doc).flatMap((drawing) => {
    const a = readAnchor(drawing);
    return a ? [{ drawing, height: a.relativeHeight }] : [];
  });
}

function nextRelativeHeight(doc: Docx): number {
  const heights = floatingHeights(doc).map((f) => f.height);
  return heights.length ? Math.max(...heights) + 1 : BASE_RELATIVE_HEIGHT;
}

/**
 * Set Word's Wrap Text style. A picture leaving the text line becomes a
 * floating object at `position` (default: the column's left edge and the
 * paragraph's top), drawn above the other floating objects.
 */
export function setDrawingWrap(
  doc: Docx,
  index: number,
  wrap: WrapStyle,
  position: { horizontal?: HorizontalPosition; vertical?: VerticalPosition } = {},
): void {
  const drawing = drawingAt(doc, index);
  applyWrap(drawing, wrap, {
    horizontal: position.horizontal ?? { relativeTo: "column", offsetEmu: 0 },
    vertical: position.vertical ?? { relativeTo: "paragraph", offsetEmu: 0 },
    relativeHeight: nextRelativeHeight(doc),
  });
  edit(doc);
}

/** Move a floating object: an alignment or an offset against a reference frame, per axis. */
export function setDrawingPosition(
  doc: Docx,
  index: number,
  position: { horizontal?: HorizontalPosition; vertical?: VerticalPosition },
): void {
  applyPosition(drawingAt(doc, index), position);
  edit(doc);
}

/** Wrap side and polygon, distance from text, overlap / lock anchor / layout in table cell. */
export function setDrawingAnchorOptions(doc: Docx, index: number, options: AnchorOptions): void {
  applyAnchorOptions(drawingAt(doc, index), options);
  edit(doc);
}

export type DrawingOrder =
  | "bringForward"
  | "bringToFront"
  | "bringInFrontOfText"
  | "sendBackward"
  | "sendToBack"
  | "sendBehindText";

/**
 * Bring Forward / Send Backward and their variants. Floating objects stack by
 * `relativeHeight`; the in/behind-text variants change the wrap style the way
 * Word's menu does.
 */
export function arrangeDrawing(doc: Docx, index: number, op: DrawingOrder): void {
  const drawing = drawingAt(doc, index);
  if (op === "bringInFrontOfText" || op === "sendBehindText") {
    setDrawingWrap(doc, index, op === "bringInFrontOfText" ? "inFrontOfText" : "behindText");
    return;
  }
  if (!readAnchor(drawing)) throw new Error("An object in line with text has no stacking order.");
  const stack = floatingHeights(doc).toSorted((a, b) => a.height - b.height);
  const at = stack.findIndex((f) => f.drawing === drawing);
  const target = {
    bringForward: at + 1,
    bringToFront: stack.length - 1,
    sendBackward: at - 1,
    sendToBack: 0,
  }[op];
  const clamped = Math.max(0, Math.min(stack.length - 1, target));
  const [moved] = stack.splice(at, 1);
  if (!moved) return;
  stack.splice(clamped, 0, moved);
  // Renumber the whole stack so ties (common in Word files) cannot hide the move.
  for (const [i, f] of stack.entries()) {
    applyAnchorOptions(f.drawing, { relativeHeight: BASE_RELATIVE_HEIGHT + i * 2 });
  }
  edit(doc);
}

/** Rotate (degrees, clockwise) and flip a picture (`a:xfrm rot` / `flipH` / `flipV`). */
export function setDrawingTransform(
  doc: Docx,
  index: number,
  transform: { rotation?: number; flipH?: boolean; flipV?: boolean },
): void {
  const { drawing, pic } = pictureAt(doc, index);
  const info = readDrawing(doc, drawing);
  applyTransform(drawing, pic, { cx: info.widthEmu, cy: info.heightEmu }, transform);
  edit(doc);
}

/** Selection Pane: rename (`wp:docPr name`). */
export function setDrawingName(doc: Docx, index: number, name: string): void {
  const docPr = child(placementOf(drawingAt(doc, index)), WP_NS, "docPr");
  if (!docPr) throw new Error("The drawing has no wp:docPr.");
  if (!name.trim()) throw new Error("An object name cannot be empty.");
  setAttr(docPr, "name", name);
  edit(doc);
}

/** Selection Pane: show / hide (`wp:docPr hidden`). */
export function setDrawingHidden(doc: Docx, index: number, hidden: boolean): void {
  const docPr = child(placementOf(drawingAt(doc, index)), WP_NS, "docPr");
  if (!docPr) throw new Error("The drawing has no wp:docPr.");
  setAttr(docPr, "hidden", hidden ? "1" : undefined);
  edit(doc);
}

/** Lock aspect ratio (`a:graphicFrameLocks` / `a:picLocks noChangeAspect`). */
export function setDrawingAspectLock(doc: Docx, index: number, locked: boolean): void {
  const drawing = drawingAt(doc, index);
  ensureDrawingNamespaces(drawing);
  const placement = placementOf(drawing);
  if (!placement) throw new Error("The drawing has neither wp:inline nor wp:anchor.");
  let frame = child(placement, WP_NS, "cNvGraphicFramePr");
  if (!frame) {
    frame = fragment(`<wp:cNvGraphicFramePr/>`);
    const kids = placement.children as XmlNode[];
    kids.splice(
      kids.indexOf(child(placement, A_NS, "graphic") ?? kids[kids.length - 1] ?? frame),
      0,
      frame,
    );
  }
  let locks = child(frame, A_NS, "graphicFrameLocks");
  if (!locks) {
    locks = fragment(`<a:graphicFrameLocks/>`);
    (frame.children as XmlNode[]).push(locks);
  }
  setAttr(locks, "noChangeAspect", locked ? "1" : undefined);
  const picLocks = descendant(drawing, A_NS, "picLocks");
  if (picLocks) setAttr(picLocks, "noChangeAspect", locked ? "1" : undefined);
  edit(doc);
}

const LINK_PROTOCOLS: ReadonlySet<string> = new Set(["http:", "https:", "mailto:"]);

/** Link a picture (`a:hlinkClick` on its docPr); `undefined` removes the link. */
export function setDrawingHyperlink(doc: Docx, index: number, url: string | undefined): void {
  const drawing = drawingAt(doc, index);
  ensureDrawingNamespaces(drawing);
  const docPr = child(placementOf(drawing), WP_NS, "docPr");
  if (!docPr) throw new Error("The drawing has no wp:docPr.");
  const targets = [docPr, descendant(drawing, PIC_NS, "cNvPr")].filter((e) => e !== undefined);
  for (const el of targets) {
    const kids = el.children as XmlNode[];
    const i = kids.findIndex(
      (k) => k.kind === "element" && k.name.uri === A_NS && k.name.local === "hlinkClick",
    );
    if (i >= 0) kids.splice(i, 1);
  }
  if (url !== undefined) {
    let protocol: string;
    try {
      protocol = new URL(url).protocol;
    } catch {
      throw new Error(`Invalid hyperlink URL: ${JSON.stringify(url)}.`);
    }
    if (!LINK_PROTOCOLS.has(protocol))
      throw new Error(`Unsupported hyperlink scheme ${JSON.stringify(protocol)}.`);
    const rel = addRelationship(partRelationships(doc.opc, doc.partName), {
      type: WML_RELATIONSHIPS.hyperlink,
      target: url,
      targetMode: "External",
    });
    for (const el of targets)
      placeChild(el, fragment(`<a:hlinkClick r:id="${rel.id}"/>`), DOC_PR_ORDER);
  }
  edit(doc);
}

/** Delete a drawing (its run goes too when nothing else is in it). */
export function removeDrawing(doc: Docx, index: number): void {
  const drawing = drawingAt(doc, index);
  const visit = (p: WmlParagraph): boolean => {
    for (const [ci, c] of p.children.entries()) {
      if (c.kind !== "run") continue;
      const pi = c.pieces.findIndex((piece) => piece.kind === "drawing" && piece.node === drawing);
      if (pi < 0) continue;
      c.pieces.splice(pi, 1);
      if (c.pieces.length === 0) p.children.splice(ci, 1);
      return true;
    }
    return false;
  };
  for (const block of doc.document.body.blocks) {
    if (block.kind === "paragraph" && visit(block)) break;
    if (
      block.kind === "table" &&
      block.rows.some((r) => r.cells.some((cell) => cell.paragraphs.some(visit)))
    )
      break;
  }
  edit(doc);
}

// --- picture format ----------------------------------------------------------

/** Crop (`a:srcRect`), in percent of the image per edge. */
export function setPictureCrop(doc: Docx, index: number, crop: PictureCrop): void {
  const { drawing, pic } = pictureAt(doc, index);
  applyCrop(drawing, pic, crop);
  edit(doc);
}

/** Crop to Shape: the picture's preset geometry (ST_ShapeType, e.g. `ellipse`, `roundRect`). */
export function setPictureGeometry(doc: Docx, index: number, preset: string): void {
  const { drawing, pic } = pictureAt(doc, index);
  applyGeometry(drawing, pic, preset);
  edit(doc);
}

/** Picture Border; `undefined` is No Outline. */
export function setPictureOutline(
  doc: Docx,
  index: number,
  outline: PictureOutline | undefined,
): void {
  const { drawing, pic } = pictureAt(doc, index);
  applyOutline(drawing, pic, outline);
  edit(doc);
}

/** Picture Effects (shadow, reflection, glow, soft edges, bevel, 3-D rotation), replacing the current set. */
export function setPictureEffects(doc: Docx, index: number, effects: PictureEffects): void {
  const { drawing, pic } = pictureAt(doc, index);
  applyEffects(drawing, pic, effects);
  edit(doc);
}

/** Corrections, Color and Transparency (blip effects), replacing the current set. */
export function setPictureColorAdjustments(
  doc: Docx,
  index: number,
  adjustments: PictureColorAdjustments,
): void {
  const { drawing, pic } = pictureAt(doc, index);
  applyAdjustments(drawing, pic, adjustments);
  edit(doc);
}

function addMediaPart(doc: Docx, bytes: Uint8Array, contentType: string | undefined): string {
  const type = contentType ?? sniffImageContentType(bytes);
  if (!type) throw new Error("Could not detect the image type; pass its content type explicitly.");
  const ext = extensionForImageContentType(type);
  let n = 1;
  while (hasPart(doc.opc, `/word/media/image${n}.${ext}`)) n++;
  const partName = `/word/media/image${n}.${ext}`;
  addPart(doc.opc, { name: partName, contentType: type, data: bytes });
  setContentTypeDefault(doc.opc.contentTypes, ext, type);
  return addRelationship(partRelationships(doc.opc, doc.partName), {
    type: WML_RELATIONSHIPS.image,
    target: partName.slice("/word/".length),
  }).id;
}

/**
 * Change Picture: point this picture at new image bytes (a new media part),
 * keeping its size, position and formatting. Other drawings that shared the
 * old image keep it.
 */
export function changePicture(
  doc: Docx,
  index: number,
  bytes: Uint8Array,
  contentType?: string,
): void {
  const { pic } = pictureAt(doc, index);
  setBlipRelationship(pic, addMediaPart(doc, bytes, contentType));
  edit(doc);
}

/**
 * Reset Picture: drop crop, shape, border, effects, color adjustments and
 * rotation. With `size`, also return to the image's own size (Reset Picture
 * & Size).
 */
export function resetPicture(doc: Docx, index: number, options: { size?: boolean } = {}): void {
  const { drawing, pic } = pictureAt(doc, index);
  applyCrop(drawing, pic, { left: 0, top: 0, right: 0, bottom: 0 });
  applyGeometry(drawing, pic, "rect");
  const spPr = child(pic, PIC_NS, "spPr");
  const ln = child(spPr, A_NS, "ln");
  if (spPr && ln) (spPr.children as XmlNode[]).splice(spPr.children.indexOf(ln), 1);
  applyEffects(drawing, pic, {});
  applyAdjustments(drawing, pic, {});
  const info = readDrawing(doc, drawing);
  applyTransform(
    drawing,
    pic,
    { cx: info.widthEmu, cy: info.heightEmu },
    { rotation: 0, flipH: false, flipV: false },
  );
  if (options.size) {
    const natural = info.picture?.image ? imageNaturalSizeEmu(info.picture.image.data) : undefined;
    if (natural) setImageSizeEmu(doc, index, natural.widthEmu, natural.heightEmu);
  }
  edit(doc);
}

// --- charts ----------------------------------------------------------------------

function writeChartParts(doc: Docx, chartPart: string, spec: ChartSpec): void {
  const rels = partRelationships(doc.opc, chartPart);
  let workbook = relationshipsByType(rels, PACKAGE_REL)[0];
  if (!workbook) {
    let n = 1;
    while (hasPart(doc.opc, `/word/embeddings/Microsoft_Excel_Worksheet${n}.xlsx`)) n++;
    const name = `/word/embeddings/Microsoft_Excel_Worksheet${n}.xlsx`;
    setContentTypeDefault(doc.opc.contentTypes, "xlsx", XLSX_CONTENT_TYPE);
    addPart(doc.opc, { name, contentType: XLSX_CONTENT_TYPE, data: new Uint8Array() });
    workbook = addRelationship(rels, {
      type: PACKAGE_REL,
      target: `../embeddings/Microsoft_Excel_Worksheet${n}.xlsx`,
    });
  }
  const workbookPart = getPart(doc.opc, resolveInternalTarget(chartPart, workbook.target));
  if (workbookPart) workbookPart.data = writeWorkbook(chartSheetRows(spec));
  const part = getPart(doc.opc, chartPart);
  if (!part) throw new Error(`Chart part ${chartPart} is missing.`);
  part.data = new TextEncoder().encode(chartSpaceXml(spec, workbook.id));
  // The raw-XML inspector caches parsed parts; drop a stale copy of this one.
  doc.rawParts.delete(chartPart);
  doc.rawPartsDirty.delete(chartPart);
}

/**
 * Build a chart: a `/word/charts/chartN.xml` part with the data cached in it,
 * an embedded workbook holding the same data (so Word's Edit Data opens it),
 * and the inline `<w:drawing>` run that shows it. Place the run with the
 * paragraph APIs.
 */
export function addChartRun(
  doc: Docx,
  spec: ChartSpec,
  options: { widthEmu: number; heightEmu: number; name?: string; altText?: string },
): WmlRun {
  validateChartSpec(spec);
  let n = 1;
  while (hasPart(doc.opc, `/word/charts/chart${n}.xml`)) n++;
  const chartPart = `/word/charts/chart${n}.xml`;
  addPart(doc.opc, { name: chartPart, contentType: CHART_CONTENT_TYPE, data: new Uint8Array() });
  writeChartParts(doc, chartPart, spec);
  const rel = addRelationship(partRelationships(doc.opc, doc.partName), {
    type: CHART_REL,
    target: `charts/chart${n}.xml`,
  });
  const id = maxDocPrId(imageDrawings(doc)) + 1;
  const decls = Object.entries(DRAWING_PREFIXES)
    .map(([prefix, uri]) => `xmlns:${prefix}="${uri}"`)
    .join(" ");
  const descr = options.altText ? ` descr="${escapeXml(options.altText)}"` : "";
  const drawing = parseXml(
    `<w:drawing xmlns:w="${WML_NS}" ${decls}><wp:inline distT="0" distB="0" distL="0" distR="0">` +
      `<wp:extent cx="${Math.round(options.widthEmu)}" cy="${Math.round(options.heightEmu)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>` +
      `<wp:docPr id="${id}" name="${escapeXml(options.name ?? `Chart ${id}`)}"${descr}/><wp:cNvGraphicFramePr/>` +
      `<a:graphic><a:graphicData uri="${C_NS}"><c:chart r:id="${rel.id}"/></a:graphicData></a:graphic>` +
      `</wp:inline></w:drawing>`,
  ).root;
  edit(doc);
  return { kind: "run", pieces: [{ kind: "drawing", node: drawing }], extras: [] };
}

/** Read a chart drawing's data and elements back from its chart part. */
export function readChart(doc: Docx, drawing: XmlElement): ChartSpec | undefined {
  const partName = chartPartOf(doc, drawing);
  const part = partName ? getPart(doc.opc, partName) : undefined;
  return part ? readChartSpace(new TextDecoder().decode(part.data)) : undefined;
}

/**
 * Rewrite the chart at `index` from `spec` (Edit Data, Chart Elements, Change
 * Colors, Change Chart Type): the chart part and its embedded workbook. The
 * part is regenerated, so formatting the spec does not describe is not kept.
 */
export function setChart(doc: Docx, index: number, spec: ChartSpec): void {
  const partName = chartPartOf(doc, drawingAt(doc, index));
  if (!partName) throw new Error(`The drawing at index ${index} is not a chart.`);
  validateChartSpec(spec);
  writeChartParts(doc, partName, spec);
  edit(doc);
}
