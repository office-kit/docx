/**
 * The page layout of each section, for the canvas: which body blocks a
 * section holds and its typed `<w:sectPr>` properties, plus the page
 * background and watermark. Pure reads; nothing here mutates the document.
 */

import {
  type Docx,
  getPageColor,
  getSectionProperties,
  getWatermark,
  type PageColor,
  sectionCount,
  getParagraphProp,
  type SectionProperties,
  type WatermarkInfo,
} from "@office-kit/docx";

export interface ResolvedSection {
  readonly index: number;
  /** First body block of the section (inclusive). */
  readonly startBlock: number;
  /** Last body block of the section (inclusive); the section-break paragraph for all but the last. */
  readonly endBlock: number;
  readonly properties: SectionProperties;
}

/** Every section, in order, with its block range and properties. */
export function resolveSectionLayout(doc: Docx): ResolvedSection[] {
  const count = sectionCount(doc);
  const blocks = doc.document.body.blocks.length;
  const starts: number[] = [0];
  doc.document.body.blocks.forEach((block, b) => {
    // A paragraph carrying a sectPr ends its section; the next block starts the next one.
    if (block.kind === "paragraph" && getParagraphProp(block, "sectPr").present) starts.push(b + 1);
  });
  return Array.from({ length: count }, (_, index) => ({
    index,
    startBlock: starts[index] ?? blocks,
    endBlock: (starts[index + 1] ?? blocks) - 1,
    properties: getSectionProperties(doc, index),
  }));
}

/** Page background and watermark, as Word shows them behind every page. */
export interface PageBackground {
  /** `undefined` when the page is plain white. */
  readonly color?: PageColor;
  readonly watermark?: WatermarkInfo;
}

export function resolvePageBackground(doc: Docx): PageBackground {
  const color = getPageColor(doc);
  const watermark = getWatermark(doc);
  return { ...(color ? { color } : {}), ...(watermark ? { watermark } : {}) };
}
