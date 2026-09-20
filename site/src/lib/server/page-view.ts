// Reads a Docx back into a plain, serialisable description of its first page
// so the landing page can draw it as HTML at prerender time.
//
// `@office-kit/docx-preview` cannot do this job: it renders into a live DOM,
// and there is none during prerender. This reader covers only what the hero
// document uses — headings, formatted runs, lists, and a bordered table —
// and is not a general renderer.

import {
  type Docx,
  getParagraphNumbering,
  getParagraphStyle,
  getRunFormat,
  numberingPart,
  stylesPart,
  type WmlParagraph,
  type WmlTable,
} from "@office-kit/docx";

type XmlElement = NonNullable<WmlParagraph["pPr"]>;

export type PageRun = { text: string; bold: boolean; italic: boolean };

export type PageBlock =
  | { kind: "heading"; level: number; text: string; sizePt?: number; color?: string; bold: boolean }
  | { kind: "paragraph"; runs: PageRun[] }
  | { kind: "list"; ordered: boolean; items: PageRun[][] }
  | { kind: "table"; borderColor?: string; rows: Array<Array<{ text: string; fill?: string }>> };

const HEADING_STYLE = /^Heading([1-9])$/;
const HALF_POINTS_PER_POINT = 2;

const child = (el: XmlElement | undefined, local: string): XmlElement | undefined =>
  el?.children.find((c): c is XmlElement => c.kind === "element" && c.name.local === local);

const attr = (el: XmlElement | undefined, local: string): string | undefined =>
  el?.attrs.find((a) => a.name.local === local)?.value;

const hexColor = (value: string | undefined): string | undefined =>
  value && value !== "auto" ? `#${value}` : undefined;

function runsOf(paragraph: WmlParagraph): PageRun[] {
  const runs: PageRun[] = [];
  for (const inline of paragraph.children) {
    if (inline.kind !== "run") continue;
    const text = inline.pieces.map((p) => (p.kind === "text" ? p.value : "")).join("");
    if (!text) continue;
    const format = getRunFormat(inline);
    runs.push({ text, bold: format.bold === true, italic: format.italic === true });
  }
  return runs;
}

const textOf = (paragraph: WmlParagraph): string =>
  runsOf(paragraph)
    .map((r) => r.text)
    .join("");

function headingBlock(doc: Docx, styleId: string, level: number, text: string): PageBlock {
  const style = stylesPart(doc)?.styles.find((s) => attr(s, "styleId") === styleId);
  const rPr = child(style, "rPr");
  const halfPoints = Number(attr(child(rPr, "sz"), "val"));
  return {
    kind: "heading",
    level,
    text,
    sizePt: Number.isFinite(halfPoints) ? halfPoints / HALF_POINTS_PER_POINT : undefined,
    color: hexColor(attr(child(rPr, "color"), "val")),
    bold: child(rPr, "b") !== undefined,
  };
}

/** numId → whether the list's first level counts (1. 2. 3.) rather than bullets. */
function orderedLists(doc: Docx): Map<number, boolean> {
  const part = numberingPart(doc);
  const formats = new Map<string, string | undefined>();
  for (const abstractNum of part?.abstractNums ?? []) {
    formats.set(
      attr(abstractNum, "abstractNumId") ?? "",
      attr(child(child(abstractNum, "lvl"), "numFmt"), "val"),
    );
  }
  const ordered = new Map<number, boolean>();
  for (const num of part?.nums ?? []) {
    const format = formats.get(attr(child(num, "abstractNumId"), "val") ?? "");
    ordered.set(Number(attr(num, "numId")), format !== undefined && format !== "bullet");
  }
  return ordered;
}

function tableBlock(table: WmlTable): PageBlock {
  const top = child(child(table.tblPr, "tblBorders"), "top");
  return {
    kind: "table",
    borderColor: attr(top, "val") === "single" ? hexColor(attr(top, "color")) : undefined,
    rows: table.rows.map((row) =>
      row.cells.map((cell) => ({
        text: cell.paragraphs.map(textOf).join("\n"),
        fill: hexColor(attr(child(cell.tcPr, "shd"), "fill")),
      })),
    ),
  };
}

export function readPage(doc: Docx): PageBlock[] {
  const ordered = orderedLists(doc);
  const blocks: PageBlock[] = [];
  for (const block of doc.document.body.blocks) {
    if (block.kind === "table") {
      blocks.push(tableBlock(block));
      continue;
    }
    if (block.kind !== "paragraph") continue;

    const styleId = getParagraphStyle(block);
    const level = styleId ? HEADING_STYLE.exec(styleId)?.[1] : undefined;
    if (styleId && level) {
      blocks.push(headingBlock(doc, styleId, Number(level), textOf(block)));
      continue;
    }

    const numbering = getParagraphNumbering(block);
    if (numbering) {
      const isOrdered = ordered.get(numbering.numId) === true;
      const last = blocks.at(-1);
      // Word has no list container: consecutive numbered paragraphs are one list.
      if (last?.kind === "list" && last.ordered === isOrdered) last.items.push(runsOf(block));
      else blocks.push({ kind: "list", ordered: isOrdered, items: [runsOf(block)] });
      continue;
    }

    blocks.push({ kind: "paragraph", runs: runsOf(block) });
  }
  return blocks;
}
