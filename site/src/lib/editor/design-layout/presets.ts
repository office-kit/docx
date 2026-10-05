/**
 * The preset lists of Word's Layout and Design menus (Margins, Size,
 * Columns, Watermark), in Word's order. Lengths are twips.
 */

import type { PageMargins, SectionColumns } from "@office-kit/docx";
import type { MessageKey } from "../i18n/index.svelte";

export interface MarginPreset {
  readonly key: MessageKey;
  readonly margins: PageMargins;
  /** Mirrored margins also turn on `w:mirrorMargins`. */
  readonly mirror: boolean;
}

const HEADER_FOOTER = { header: 720, footer: 720, gutter: 0 } as const;

export const MARGIN_PRESETS: readonly MarginPreset[] = [
  {
    key: "lay.margins.normal",
    mirror: false,
    margins: { top: 1440, bottom: 1440, left: 1440, right: 1440, ...HEADER_FOOTER },
  },
  {
    key: "lay.margins.narrow",
    mirror: false,
    margins: { top: 720, bottom: 720, left: 720, right: 720, ...HEADER_FOOTER },
  },
  {
    key: "lay.margins.moderate",
    mirror: false,
    margins: { top: 1440, bottom: 1440, left: 1080, right: 1080, ...HEADER_FOOTER },
  },
  {
    key: "lay.margins.wide",
    mirror: false,
    margins: { top: 1440, bottom: 1440, left: 2880, right: 2880, ...HEADER_FOOTER },
  },
  {
    key: "lay.margins.mirrored",
    mirror: true,
    margins: { top: 1440, bottom: 1440, left: 1800, right: 1440, ...HEADER_FOOTER },
  },
  {
    key: "lay.margins.office2003",
    mirror: false,
    margins: { top: 1440, bottom: 1440, left: 1800, right: 1800, ...HEADER_FOOTER },
  },
];

export interface PaperPreset {
  readonly name: string;
  readonly widthTwips: number;
  readonly heightTwips: number;
  /** Word's printer paper code (`w:pgSz/@w:code`). */
  readonly code: number;
  /** Show dimensions in mm (ISO / JIS sizes) rather than inches. */
  readonly metric: boolean;
}

// Word's Size menu. Codes are the Windows DMPAPER_* values Word writes.
export const PAPER_PRESETS: readonly PaperPreset[] = [
  { name: "Letter", widthTwips: 12240, heightTwips: 15840, code: 1, metric: false },
  { name: "Legal", widthTwips: 12240, heightTwips: 20160, code: 5, metric: false },
  { name: "Executive", widthTwips: 10440, heightTwips: 15120, code: 7, metric: false },
  { name: "Tabloid", widthTwips: 15840, heightTwips: 24480, code: 3, metric: false },
  { name: "A3", widthTwips: 16838, heightTwips: 23811, code: 8, metric: true },
  { name: "A4", widthTwips: 11906, heightTwips: 16838, code: 9, metric: true },
  { name: "A5", widthTwips: 8391, heightTwips: 11906, code: 11, metric: true },
  { name: "B4 (JIS)", widthTwips: 14570, heightTwips: 20636, code: 12, metric: true },
  { name: "B5 (JIS)", widthTwips: 10318, heightTwips: 14570, code: 13, metric: true },
  { name: "Envelope #10", widthTwips: 5940, heightTwips: 13680, code: 20, metric: false },
  { name: "Envelope DL", widthTwips: 6236, heightTwips: 12474, code: 27, metric: true },
  { name: "Envelope C5", widthTwips: 9184, heightTwips: 12983, code: 28, metric: true },
  { name: "Envelope Monarch", widthTwips: 5580, heightTwips: 10800, code: 37, metric: false },
  { name: "Japanese Postcard", widthTwips: 5669, heightTwips: 8391, code: 43, metric: true },
];

// Metric sizes are rounded to whole twips, so a saved A4 can be a twip or two off.
const PAPER_TOLERANCE_TWIPS = 2;
const close = (a: number, b: number): boolean => Math.abs(a - b) <= PAPER_TOLERANCE_TWIPS;

/** Whether a page size matches a preset (either orientation, within a twip of rounding). */
export function matchesPaper(p: PaperPreset, width: number, height: number): boolean {
  return (
    (close(p.widthTwips, width) && close(p.heightTwips, height)) ||
    (close(p.widthTwips, height) && close(p.heightTwips, width))
  );
}

export type ColumnPresetId = "one" | "two" | "three" | "left" | "right";

// Word's Left / Right presets: the narrow column is 11/36 of the text width
// left after the gap (1.83" of 6" on Letter with 1" margins).
const NARROW_COLUMN_FRACTION = 11 / 36;
const COLUMN_GAP = 720;

/** A Columns preset for a section whose text area is `textWidth` twips wide. */
export function columnPreset(id: ColumnPresetId, textWidth: number): SectionColumns {
  if (id === "one" || id === "two" || id === "three") {
    const count = id === "one" ? 1 : id === "two" ? 2 : 3;
    return { count, spaceTwips: COLUMN_GAP, separator: false };
  }
  const available = textWidth - COLUMN_GAP;
  const narrow = Math.round(available * NARROW_COLUMN_FRACTION);
  const wide = available - narrow;
  const [first, second] = id === "left" ? [narrow, wide] : [wide, narrow];
  return {
    count: 2,
    spaceTwips: COLUMN_GAP,
    separator: false,
    columns: [{ widthTwips: first, spaceTwips: COLUMN_GAP }, { widthTwips: second }],
  };
}

/** Which preset a section's columns are, for the menu's check mark. */
export function columnPresetOf(columns: SectionColumns): ColumnPresetId | undefined {
  if (!columns.columns) {
    if (columns.count === 1) return "one";
    if (columns.count === 2) return "two";
    if (columns.count === 3) return "three";
    return undefined;
  }
  const [a, b] = columns.columns;
  if (columns.count !== 2 || !a || !b) return undefined;
  if (a.widthTwips < b.widthTwips) return "left";
  if (a.widthTwips > b.widthTwips) return "right";
  return undefined;
}

export interface WatermarkPreset {
  readonly key: MessageKey;
  readonly layout: "diagonal" | "horizontal";
}

/** Design ▸ Watermark gallery, in Word's groups (Confidential, Disclaimers, Urgent). */
export const WATERMARK_PRESETS: ReadonlyArray<{
  readonly group: MessageKey;
  readonly items: readonly WatermarkPreset[];
}> = [
  {
    group: "dsn.wm.groupConfidential",
    items: [
      { key: "dsn.wm.confidential", layout: "diagonal" },
      { key: "dsn.wm.confidential", layout: "horizontal" },
      { key: "dsn.wm.doNotCopy", layout: "diagonal" },
      { key: "dsn.wm.doNotCopy", layout: "horizontal" },
    ],
  },
  {
    group: "dsn.wm.groupDisclaimers",
    items: [
      { key: "dsn.wm.draft", layout: "diagonal" },
      { key: "dsn.wm.draft", layout: "horizontal" },
      { key: "dsn.wm.sample", layout: "diagonal" },
      { key: "dsn.wm.sample", layout: "horizontal" },
    ],
  },
  {
    group: "dsn.wm.groupUrgent",
    items: [
      { key: "dsn.wm.asap", layout: "diagonal" },
      { key: "dsn.wm.asap", layout: "horizontal" },
      { key: "dsn.wm.urgent", layout: "diagonal" },
      { key: "dsn.wm.urgent", layout: "horizontal" },
    ],
  },
];
