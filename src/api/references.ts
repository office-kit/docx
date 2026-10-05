/**
 * Word's References tab: tables of contents and figures, footnotes and
 * endnotes, citations and bibliography (ECMA-376 Part 1 §22.6), captions,
 * index entries and the index, and tables of authorities.
 *
 * Every generated table is a field whose result this module computes the way
 * Word does on "Update Table", so a document saved from here shows the right
 * entries before Word ever opens it. Page numbers come from a caller-supplied
 * {@link PageNumberProvider} (a laid-out view knows the real pages); without
 * one they are estimated from the page and section breaks in the document.
 */

import type { XmlElement } from "../internal/xml/index.js";
import {
  addPart,
  addRelationship,
  getPart,
  hasPart,
  partRelationships,
  relationshipsByType,
} from "../internal/opc/index.js";
import {
  WML_NS,
  type WmlBlock,
  type WmlInline,
  type WmlParagraph,
} from "../internal/wordprocessingml/index.js";
import {
  fieldSwitch,
  formatFieldNumber,
  formatNumber,
  GENERAL_FORMAT_SWITCH,
  type FieldInstruction,
  type NumberingFormat,
  quoteFieldArg,
  quotedFieldArg,
} from "../internal/wordprocessingml/field-code.js";
import {
  blockFieldSkeletonXml,
  childElement,
  elementFromXml,
  elementXml,
  escapeXml,
  estimatePages,
  fieldRunsXml,
  type FlowField,
  inlinesFromXml,
  type ParagraphRef,
  paragraphFlow,
  paragraphsFromXml,
  removeField,
  resultlessFieldRunsXml,
  scanFields,
  setBlockResult,
  setInlineResult,
  textRunXml,
  visibleInlinesText,
  visibleParagraphText,
  wAttr,
  withOpenContentControls,
} from "../internal/wordprocessingml/field-flow.js";
import {
  BIBLIOGRAPHY_NS,
  type BibliographySource,
  CITATION_STYLES,
  type CitationOptions,
  type CitationStyle,
  type CitedSource,
  formatBibliography,
  formatCitation,
  parseSources,
  type StyledText,
  writeSources,
} from "../internal/wordprocessingml/bibliography.js";
import {
  findChild,
  NOTE_PR_ORDER,
  P_PR_ORDER,
  removeChild,
  SECT_PR_ORDER,
  SETTINGS_ORDER,
  setOrderedChild,
} from "../internal/wordprocessingml/ordered.js";
import { type Docx, ensureHeadingStyles } from "./docx.js";
import {
  editSettings,
  ensureBuiltInStyles,
  ensureNotesPart,
  insertInlinesAt,
  markNotesDirty,
  type NoteKind,
  notesPartOf,
  paragraphStyleId,
  ensureBodySectPr,
  readSettings,
  styleOutlineLevel,
  textWidthTwips,
} from "./part-helpers.js";

export type { NoteKind } from "./part-helpers.js";
export {
  type BibliographySource,
  CITATION_STYLES,
  type CitationOptions,
  type CitationStyle,
  type PersonName,
  SOURCE_FIELDS,
  SOURCE_TYPES,
  type SourceField,
  type SourceType,
} from "../internal/wordprocessingml/bibliography.js";
export type { NumberingFormat } from "../internal/wordprocessingml/field-code.js";

/**
 * The page a paragraph is laid out on, or `undefined` when unknown (the
 * estimate from the document's own breaks is used then).
 */
export type PageNumberProvider = (paragraph: WmlParagraph) => number | undefined;

/** Leader characters of a right-aligned tab (ST_TabTlc, §17.18.85). */
export type TabLeader = "none" | "dot" | "hyphen" | "underscore" | "middleDot";

// --- Built-in styles ---------------------------------------------------------

const TOC_INDENT_TWIPS = 220;
const HANGING_INDENT_TWIPS = 720;

function tocStyleXml(level: number): string {
  const indent = level > 1 ? `<w:ind w:left="${(level - 1) * TOC_INDENT_TWIPS}"/>` : "";
  return `<w:style w:type="paragraph" w:styleId="TOC${level}"><w:name w:val="toc ${level}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:autoRedefine/><w:uiPriority w:val="39"/><w:unhideWhenUsed/><w:pPr><w:spacing w:after="100"/>${indent}</w:pPr></w:style>`;
}

function indexStyleXml(level: number): string {
  return `<w:style w:type="paragraph" w:styleId="Index${level}"><w:name w:val="index ${level}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:autoRedefine/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:ind w:left="${level * TOC_INDENT_TWIPS}" w:hanging="${TOC_INDENT_TWIPS}"/></w:pPr></w:style>`;
}

const MAX_LEVEL = 9;
const LEVELS = Array.from({ length: MAX_LEVEL }, (_, i) => i + 1);

const TOC_STYLES: Record<string, string> = Object.fromEntries(
  LEVELS.map((l) => [`TOC${l}`, tocStyleXml(l)]),
);
const TOC_HEADING_STYLE = {
  TOCHeading:
    '<w:style w:type="paragraph" w:styleId="TOCHeading"><w:name w:val="TOC Heading"/><w:basedOn w:val="Heading1"/><w:next w:val="Normal"/><w:uiPriority w:val="39"/><w:unhideWhenUsed/><w:qFormat/><w:pPr><w:outlineLvl w:val="9"/></w:pPr></w:style>',
};
const CAPTION_STYLES = {
  Caption:
    '<w:style w:type="paragraph" w:styleId="Caption"><w:name w:val="caption"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="35"/><w:unhideWhenUsed/><w:qFormat/><w:pPr><w:spacing w:after="200" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:i/><w:iCs/><w:color w:val="44546A"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:style>',
};
const TABLE_OF_FIGURES_STYLES = {
  TableofFigures:
    '<w:style w:type="paragraph" w:styleId="TableofFigures"><w:name w:val="table of figures"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:pPr><w:spacing w:after="0"/></w:pPr></w:style>',
};
const INDEX_STYLES: Record<string, string> = {
  ...Object.fromEntries(LEVELS.map((l) => [`Index${l}`, indexStyleXml(l)])),
  IndexHeading:
    '<w:style w:type="paragraph" w:styleId="IndexHeading"><w:name w:val="index heading"/><w:basedOn w:val="Normal"/><w:next w:val="Index1"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:pPr><w:spacing w:before="240" w:after="120"/></w:pPr><w:rPr><w:b/><w:bCs/></w:rPr></w:style>',
};
const TOA_STYLES = {
  TableofAuthorities:
    '<w:style w:type="paragraph" w:styleId="TableofAuthorities"><w:name w:val="table of authorities"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:pPr><w:spacing w:after="0"/><w:ind w:left="220" w:hanging="220"/></w:pPr></w:style>',
  TOAHeading:
    '<w:style w:type="paragraph" w:styleId="TOAHeading"><w:name w:val="toa heading"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:pPr><w:spacing w:before="120"/></w:pPr><w:rPr><w:b/><w:bCs/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:style>',
};
const BIBLIOGRAPHY_STYLES = {
  Bibliography:
    '<w:style w:type="paragraph" w:styleId="Bibliography"><w:name w:val="Bibliography"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="37"/><w:unhideWhenUsed/></w:style>',
};
const NOTE_STYLES = {
  footnote: {
    FootnoteText:
      '<w:style w:type="paragraph" w:styleId="FootnoteText"><w:name w:val="footnote text"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style>',
    FootnoteReference:
      '<w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>',
  },
  endnote: {
    EndnoteText:
      '<w:style w:type="paragraph" w:styleId="EndnoteText"><w:name w:val="endnote text"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style>',
    EndnoteReference:
      '<w:style w:type="character" w:styleId="EndnoteReference"><w:name w:val="endnote reference"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>',
  },
} as const;

// --- Shared helpers ----------------------------------------------------------

/** Run `fn` over the body with Word's content controls opened, then mark the document dirty. */
function withFlow<T>(doc: Docx, fn: (flow: ParagraphRef[]) => T): T {
  const result = withOpenContentControls(doc.document.body, () =>
    fn(paragraphFlow(doc.document.body)),
  );
  doc.dirty = true;
  return result;
}

/** Page lookup: the provider's answer, else the estimate. */
function pageLookup(doc: Docx, provider: PageNumberProvider | undefined) {
  const estimate = estimatePages(doc.document.body);
  return (paragraph: WmlParagraph): number => provider?.(paragraph) ?? estimate.get(paragraph) ?? 1;
}

function insertBlocks(doc: Docx, at: number, blocks: readonly WmlBlock[]): void {
  const list = doc.document.body.blocks;
  // Keep the body's trailing section paragraph order intact: never insert
  // past the end.
  list.splice(Math.max(0, Math.min(at, list.length)), 0, ...blocks);
  doc.dirty = true;
}

function tabsXml(leader: TabLeader, position: number): string {
  return `<w:tabs><w:tab w:val="right" w:leader="${leader}" w:pos="${position}"/></w:tabs>`;
}

/** The right-tab leader of a generated table's first paragraph (what Update keeps). */
function existingLeader(paragraph: WmlParagraph): TabLeader {
  const tab = childElement(childElement(paragraph.pPr, "tabs"), "tab");
  const leader = wAttr(tab, "leader");
  if (!tab) return "dot";
  return leader === "dot" ||
    leader === "hyphen" ||
    leader === "underscore" ||
    leader === "middleDot"
    ? leader
    : "none";
}

interface BookmarkIndex {
  readonly names: Set<string>;
  nextId: number;
}

function bookmarkIndex(flow: readonly ParagraphRef[]): BookmarkIndex {
  const names = new Set<string>();
  let maxId = -1;
  for (const { paragraph } of flow) {
    for (const inline of paragraph.children) {
      if (inline.kind !== "raw" || inline.node.name.local !== "bookmarkStart") continue;
      const name = wAttr(inline.node, "name");
      if (name) names.add(name);
      maxId = Math.max(maxId, Number(wAttr(inline.node, "id") ?? -1));
    }
  }
  return { names, nextId: maxId + 1 };
}

// Word names TOC bookmarks `_Toc` + nine digits.
const TOC_BOOKMARK_BASE = 100_000_000;

/**
 * The hidden `_Toc…` bookmark wrapping a heading, added when missing: TOC
 * entries link to it (`\h`) and take their page number from it (PAGEREF).
 */
function ensureTocBookmark(paragraph: WmlParagraph, index: BookmarkIndex): string {
  for (const inline of paragraph.children) {
    if (inline.kind !== "raw" || inline.node.name.local !== "bookmarkStart") continue;
    const name = wAttr(inline.node, "name");
    if (name?.startsWith("_Toc")) return name;
  }
  let n = TOC_BOOKMARK_BASE + index.names.size;
  while (index.names.has(`_Toc${n}`)) n++;
  const name = `_Toc${n}`;
  index.names.add(name);
  const id = index.nextId++;
  const [start] = inlinesFromXml(`<w:bookmarkStart w:id="${id}" w:name="${name}"/>`);
  const [end] = inlinesFromXml(`<w:bookmarkEnd w:id="${id}"/>`);
  if (start && end) paragraph.children = [start, ...paragraph.children, end];
  return name;
}

function rangeArg(arg: string | undefined, fallback: [number, number]): [number, number] {
  const m = arg ? /^(\d)\s*-\s*(\d)$/.exec(arg) : null;
  return m ? [Number(m[1]), Number(m[2])] : fallback;
}

// --- SEQ numbering -----------------------------------------------------------

/**
 * The number every SEQ field shows, in document order (§17.16.5.56): `\r n`
 * resets to n, `\c` repeats the last number, and `\s n` restarts after each
 * heading of level n or higher.
 */
function sequenceNumbers(
  doc: Docx,
  flow: readonly ParagraphRef[],
  fields: readonly FlowField[],
): Map<FlowField, number> {
  const out = new Map<FlowField, number>();
  const byParagraph = new Map<WmlParagraph, FlowField[]>();
  const resetLevel = new Map<string, number>();
  for (const f of fields) {
    if (f.parsed.type !== "SEQ") continue;
    const list = byParagraph.get(f.begin.ref.paragraph) ?? [];
    list.push(f);
    byParagraph.set(f.begin.ref.paragraph, list);
    const s = fieldSwitch(f.parsed, "s")?.arg;
    if (s !== undefined) resetLevel.set(seqId(f), Number(s));
  }
  const counters = new Map<string, number>();
  for (const { paragraph } of flow) {
    if (resetLevel.size > 0) {
      const level = styleOutlineLevel(doc, paragraphStyleId(paragraph));
      if (level !== undefined) {
        for (const [id, at] of resetLevel) if (level <= at) counters.set(id, 0);
      }
    }
    for (const f of byParagraph.get(paragraph) ?? []) {
      const id = seqId(f);
      const reset = fieldSwitch(f.parsed, "r")?.arg;
      let value = counters.get(id) ?? 0;
      if (reset !== undefined && Number.isFinite(Number(reset))) value = Number(reset);
      else if (!fieldSwitch(f.parsed, "c")) value += 1;
      counters.set(id, value);
      out.set(f, value);
    }
  }
  return out;
}

function seqId(field: FlowField): string {
  return (field.parsed.args[0] ?? "").toLowerCase();
}

// --- Table of contents / figures ----------------------------------------------

export interface TableOfContentsOptions {
  /** Heading levels shown (`\o`). Default 1-3. */
  readonly levels?: { readonly from: number; readonly to: number };
  /** Include paragraphs by their applied outline level (`\u`). Default true. */
  readonly useOutlineLevels?: boolean;
  /** Entries link to their headings (`\h`). Default true. */
  readonly hyperlinks?: boolean;
  /** Hide tab leaders and page numbers in Web Layout (`\z`). Default true. */
  readonly hideInWebLayout?: boolean;
  /** Show page numbers. Default true; false writes `\n`. */
  readonly pageNumbers?: boolean;
  /** Right-align page numbers on a tab. Default true; false writes `\p " "`. */
  readonly rightAlignPageNumbers?: boolean;
  readonly tabLeader?: TabLeader;
  /**
   * Build a table of figures from captions with this label (`\c "Figure"`)
   * instead of a table of contents from headings.
   */
  readonly captionLabel?: string;
  /** With `captionLabel`: show "Figure 1:" before each entry. Default true (`\c`; false is `\a`). */
  readonly includeLabelAndNumber?: boolean;
  /** A title paragraph ("Contents") in the TOC Heading style before the table. */
  readonly title?: string;
  readonly pageOf?: PageNumberProvider;
}

/** The TOC field instruction Word writes for these options. */
export function tableOfContentsInstruction(options: TableOfContentsOptions = {}): string {
  const parts = ["TOC"];
  if (options.captionLabel !== undefined) {
    if (options.hyperlinks ?? true) parts.push("\\h");
    if (options.hideInWebLayout ?? true) parts.push("\\z");
    parts.push(
      options.includeLabelAndNumber === false ? "\\a" : "\\c",
      quotedFieldArg(options.captionLabel),
    );
  } else {
    const from = options.levels?.from ?? 1;
    const to = options.levels?.to ?? 3;
    if (!Number.isInteger(from) || !Number.isInteger(to) || from < 1 || to < from || to > 9) {
      throw new Error(`Invalid TOC level range ${from}-${to}; expected 1 <= from <= to <= 9.`);
    }
    parts.push("\\o", `"${from}-${to}"`);
    if (options.hyperlinks ?? true) parts.push("\\h");
    if (options.hideInWebLayout ?? true) parts.push("\\z");
    if (options.useOutlineLevels ?? true) parts.push("\\u");
  }
  if (options.pageNumbers === false) parts.push("\\n");
  if (options.rightAlignPageNumbers === false) parts.push("\\p", '" "');
  return parts.join(" ");
}

interface TocEntry {
  readonly level: number;
  readonly text: string;
  readonly paragraph: WmlParagraph;
}

function headingEntries(doc: Docx, flow: readonly ParagraphRef[], parsed: FieldInstruction): TocEntry[] {
  const outline = fieldSwitch(parsed, "o");
  const useLevels = !!fieldSwitch(parsed, "u");
  const custom = fieldSwitch(parsed, "t")?.arg;
  if (!outline && !useLevels && !custom) return [];
  const [from, to] = rangeArg(outline?.arg, [1, MAX_LEVEL]);
  // `\t "Style,1,Other,2"`: extra styles by name, with the level they map to.
  const styleLevels = new Map<string, number>();
  if (custom) {
    const parts = custom.split(/[,;]/).map((s) => s.trim());
    for (let i = 0; i + 1 < parts.length; i += 2) {
      styleLevels.set((parts[i] as string).toLowerCase(), Number(parts[i + 1]));
    }
  }
  const styleNames = new Map<string, string>();
  const out: TocEntry[] = [];
  for (const { paragraph } of flow) {
    const styleId = paragraphStyleId(paragraph);
    if (styleId?.startsWith("TOC")) continue;
    const direct = wAttr(childElement(paragraph.pPr, "outlineLvl"), "val");
    // Outline level 9 is "body text": Add Text ▸ Do Not Show in Table of Contents.
    if (direct === "9") continue;
    let level: number | undefined;
    if (useLevels && direct !== undefined) level = Number(direct) + 1;
    if (level === undefined && outline) level = styleOutlineLevel(doc, styleId);
    if (level === undefined && styleId && styleLevels.size > 0) {
      let name = styleNames.get(styleId);
      if (name === undefined) {
        const el = doc.stylesCache?.styles.find((s) => wAttr(s, "styleId") === styleId);
        name = (wAttr(childElement(el, "name"), "val") ?? styleId).toLowerCase();
        styleNames.set(styleId, name);
      }
      level = styleLevels.get(name) ?? styleLevels.get(styleId.toLowerCase());
    }
    if (level === undefined || level < from || level > to) continue;
    const text = visibleParagraphText(paragraph).replace(/\t/g, " ").trim();
    if (text) out.push({ level, text, paragraph });
  }
  return out;
}

function figureEntries(
  fields: readonly FlowField[],
  numbers: ReadonlyMap<FlowField, number>,
  label: string,
  withLabel: boolean,
): TocEntry[] {
  const out: TocEntry[] = [];
  const seen = new Set<WmlParagraph>();
  for (const f of fields) {
    if (f.parsed.type !== "SEQ" || (f.parsed.args[0] ?? "").toLowerCase() !== label.toLowerCase())
      continue;
    const paragraph = f.begin.ref.paragraph;
    if (seen.has(paragraph) || f.begin.ref !== f.end.ref) continue;
    seen.add(paragraph);
    const before = visibleInlinesText(paragraph.children.slice(0, f.begin.inline));
    const after = visibleInlinesText(paragraph.children.slice(f.end.inline + 1));
    const number = formatFieldNumber(numbers.get(f) ?? 1, f.parsed);
    const text = withLabel
      ? `${before}${number}${after}`
      : after.replace(/^[\s:.\-–—]+/, "");
    out.push({ level: 1, text: text.replace(/\t/g, " ").trim(), paragraph });
  }
  return out;
}

interface TocLayout {
  readonly leader: TabLeader;
  readonly width: number;
}

function tocResult(
  doc: Docx,
  flow: readonly ParagraphRef[],
  field: FlowField,
  fields: readonly FlowField[],
  numbers: ReadonlyMap<FlowField, number>,
  pageOf: (p: WmlParagraph) => number,
  bookmarks: ReadonlyMap<WmlParagraph, string>,
  layout: TocLayout,
): WmlParagraph[] {
  const parsed = field.parsed;
  const caption = fieldSwitch(parsed, "c") ?? fieldSwitch(parsed, "a");
  const figures = caption?.arg !== undefined;
  const entries = figures
    ? figureEntries(fields, numbers, caption.arg as string, caption.name.toLowerCase() === "c")
    : headingEntries(doc, flow, parsed);
  if (entries.length === 0) {
    const message = figures
      ? "No table of figures entries found."
      : "No table of contents entries found.";
    return paragraphsFromXml(`<w:p>${textRunXml(message, "<w:b/><w:bCs/><w:noProof/>")}</w:p>`);
  }
  const hyperlinks = !!fieldSwitch(parsed, "h");
  const omit = fieldSwitch(parsed, "n");
  const [omitFrom, omitTo] = omit ? rangeArg(omit.arg, [1, MAX_LEVEL]) : [0, -1];
  const separator = fieldSwitch(parsed, "p")?.arg;
  const xml = entries.map((entry) => {
    const style = figures ? "TableofFigures" : `TOC${entry.level}`;
    const anchor = bookmarks.get(entry.paragraph);
    const showPage = entry.level < omitFrom || entry.level > omitTo;
    const tabs = showPage && separator === undefined ? tabsXml(layout.leader, layout.width) : "";
    let content = textRunXml(entry.text, "<w:noProof/>");
    if (showPage && anchor) {
      const sep = separator === undefined ? "<w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:tab/></w:r>" : textRunXml(separator, "<w:noProof/><w:webHidden/>");
      content += sep;
      content += fieldRunsXml(
        `PAGEREF ${anchor} \\h`,
        textRunXml(String(pageOf(entry.paragraph)), "<w:noProof/><w:webHidden/>"),
        "<w:noProof/><w:webHidden/>",
      );
    }
    if (hyperlinks && anchor) {
      content = `<w:hyperlink w:anchor="${escapeXml(anchor)}" w:history="1">${content}</w:hyperlink>`;
    }
    return `<w:p><w:pPr><w:pStyle w:val="${style}"/>${tabs}</w:pPr>${content}</w:p>`;
  });
  return paragraphsFromXml(xml.join(""));
}

/**
 * Insert a table of contents (or, with `captionLabel`, a table of figures)
 * before body block `at`, with its entries computed. Automatic Table 1 and 2
 * of Word's gallery differ only in the title ("Contents" / "Table of
 * Contents").
 *
 * Word wraps its gallery tables in a `<w:sdt>` building-block control; this
 * writes the TOC field directly, which Word reads and updates the same way.
 */
export function insertTableOfContents(
  doc: Docx,
  at: number,
  options: TableOfContentsOptions = {},
): void {
  const figures = options.captionLabel !== undefined;
  ensureBuiltInStyles(doc, figures ? TABLE_OF_FIGURES_STYLES : TOC_STYLES);
  const blocks: WmlBlock[] = [];
  if (options.title !== undefined) {
    ensureHeadingStyles(doc, 1);
    ensureBuiltInStyles(doc, TOC_HEADING_STYLE);
    blocks.push(
      ...paragraphsFromXml(
        `<w:p><w:pPr><w:pStyle w:val="TOCHeading"/></w:pPr>${textRunXml(options.title)}</w:p>`,
      ),
    );
  }
  const leader = options.tabLeader ?? "dot";
  const instruction = tableOfContentsInstruction(options);
  blocks.push(
    ...paragraphsFromXml(
      blockFieldSkeletonXml(instruction, tabsXml(leader, textWidthTwips(doc))),
    ),
  );
  insertBlocks(doc, at, blocks);
  updateTables(doc, { types: ["TOC"], ...(options.pageOf ? { pageOf: options.pageOf } : {}) });
}

/**
 * Remove every table of contents (TOC fields without `\c` / `\a`), with the
 * TOC Heading title paragraph just before each. Returns how many were removed.
 */
export function removeTableOfContents(doc: Docx): number {
  return withFlow(doc, (flow) => {
    const fields = scanFields(flow).filter(
      (f) =>
        f.depth === 0 &&
        f.parsed.type === "TOC" &&
        !fieldSwitch(f.parsed, "c") &&
        !fieldSwitch(f.parsed, "a"),
    );
    for (const f of fields.toReversed()) {
      const list = f.begin.ref.list as WmlParagraph[];
      const titleIndex = f.begin.ref.index - 1;
      removeField(f);
      const title = list[titleIndex];
      if (title?.kind === "paragraph" && paragraphStyleId(title) === "TOCHeading") {
        list.splice(titleIndex, 1);
      }
    }
    return fields.length;
  });
}

/** Add Text level: 1-9 applies Heading N; `"none"` keeps the paragraph out of tables of contents. */
export type TocLevel = number | "none";

/**
 * Word's References ▸ Add Text: make a paragraph a TOC entry of a level by
 * applying the Heading style, or exclude it with outline level "body text"
 * (`w:outlineLvl w:val="9"`, which `\u` honours).
 */
export function setTocLevel(doc: Docx, paragraph: WmlParagraph, level: TocLevel): void {
  const pPr = paragraph.pPr ?? elementFromXml("<w:pPr/>");
  paragraph.pPr = pPr;
  if (level === "none") {
    setOrderedChild(pPr, elementFromXml('<w:outlineLvl w:val="9"/>'), P_PR_ORDER);
  } else {
    if (!Number.isInteger(level) || level < 1 || level > MAX_LEVEL) {
      throw new Error(`TOC level must be 1-9, got ${level}.`);
    }
    ensureHeadingStyles(doc, level);
    removeChild(pPr, "outlineLvl");
    setOrderedChild(pPr, elementFromXml(`<w:pStyle w:val="Heading${level}"/>`), P_PR_ORDER);
  }
  doc.dirty = true;
}

/** The TOC level of a paragraph as Add Text shows it (checked item). */
export function tocLevel(doc: Docx, paragraph: WmlParagraph): TocLevel {
  const direct = wAttr(childElement(paragraph.pPr, "outlineLvl"), "val");
  if (direct === "9") return "none";
  if (direct !== undefined) return Number(direct) + 1;
  return styleOutlineLevel(doc, paragraphStyleId(paragraph)) ?? "none";
}

// --- Update Table --------------------------------------------------------------

/** The generated tables {@link updateTables} recomputes. */
export type TableFieldType = "TOC" | "INDEX" | "TOA" | "BIBLIOGRAPHY" | "CITATION";

const TABLE_TYPES: readonly TableFieldType[] = ["TOC", "INDEX", "TOA", "BIBLIOGRAPHY", "CITATION"];

export interface UpdateTablesOptions {
  /** Which kinds to update. Default: all. */
  readonly types?: readonly TableFieldType[];
  readonly pageOf?: PageNumberProvider;
}

/**
 * Word's "Update Table" for the generated tables: recompute the result of
 * every TOC (contents and figures), INDEX, TOA, BIBLIOGRAPHY and CITATION
 * field from the document. Returns the number of fields updated.
 */
export function updateTables(doc: Docx, options: UpdateTablesOptions = {}): number {
  const types = new Set(options.types ?? TABLE_TYPES);
  return withFlow(doc, (initialFlow) => {
    // Headings get their `_Toc` bookmarks first: adding them shifts inline
    // indexes, so fields are scanned afterwards.
    const tocBookmarks = new Map<WmlParagraph, string>();
    if (types.has("TOC")) {
      const fields = scanFields(initialFlow);
      const tocs = fields.filter((f) => f.depth === 0 && f.parsed.type === "TOC");
      if (tocs.length > 0) {
        const numbers = sequenceNumbers(doc, initialFlow, fields);
        const index = bookmarkIndex(initialFlow);
        for (const toc of tocs) {
          const caption = fieldSwitch(toc.parsed, "c") ?? fieldSwitch(toc.parsed, "a");
          const entries =
            caption?.arg !== undefined
              ? figureEntries(fields, numbers, caption.arg, true)
              : headingEntries(doc, initialFlow, toc.parsed);
          for (const e of entries) {
            if (!tocBookmarks.has(e.paragraph)) {
              tocBookmarks.set(e.paragraph, ensureTocBookmark(e.paragraph, index));
            }
          }
        }
      }
    }
    const flow = paragraphFlow(doc.document.body);
    const fields = scanFields(flow);
    const pageOf = pageLookup(doc, options.pageOf);
    const numbers = sequenceNumbers(doc, flow, fields);
    const bibliography = types.has("BIBLIOGRAPHY") || types.has("CITATION") ? readBibliography(doc) : undefined;
    const citationNumbers = citationOrder(fields);
    const targets = fields.filter(
      (f) => f.depth === 0 && types.has(f.parsed.type as TableFieldType),
    );
    // Last to first, so replacing one result never moves an earlier field.
    for (const f of targets.toReversed()) {
      switch (f.parsed.type) {
        case "TOC": {
          ensureBuiltInStyles(doc, fieldSwitch(f.parsed, "c") || fieldSwitch(f.parsed, "a") ? TABLE_OF_FIGURES_STYLES : TOC_STYLES);
          const layout = { leader: existingLeader(f.begin.ref.paragraph), width: textWidthTwips(doc) };
          setBlockResult(f, tocResult(doc, flow, f, fields, numbers, pageOf, tocBookmarks, layout));
          break;
        }
        case "INDEX":
          ensureBuiltInStyles(doc, INDEX_STYLES);
          setBlockResult(f, indexResult(doc, f, fields, pageOf));
          break;
        case "TOA":
          ensureBuiltInStyles(doc, TOA_STYLES);
          setBlockResult(f, toaResult(doc, f, fields, pageOf));
          break;
        case "BIBLIOGRAPHY":
          ensureBuiltInStyles(doc, BIBLIOGRAPHY_STYLES);
          setBlockResult(f, bibliographyResult(bibliography, citationNumbers));
          break;
        case "CITATION":
          setInlineResult(f, citationResult(bibliography, f.parsed, citationNumbers));
          break;
      }
    }
    return targets.length;
  });
}

// --- Captions ---------------------------------------------------------------------

/** Word's built-in caption labels. */
export const CAPTION_LABELS = ["Figure", "Table", "Equation"] as const;

export interface CaptionOptions {
  /** The label ("Figure", "Table", "Equation", or one added with {@link addCaptionLabel}). */
  readonly label: string;
  /** Text after the number, as typed in Word's Caption box (e.g. ": Results"). */
  readonly text?: string;
  /** Leave the label out of the caption ("Exclude label from caption"). */
  readonly excludeLabel?: boolean;
  /** Numbering format of the SEQ field. Default decimal. */
  readonly numberFormat?: NumberingFormat;
  /**
   * Include the chapter number: the heading level that starts a chapter and
   * the separator Word puts between chapter and caption number.
   */
  readonly chapter?: { readonly headingLevel: number; readonly separator: string };
}

/**
 * Insert a caption paragraph (Caption style) before or after body block
 * `blockIndex`: label, a `SEQ` field numbered in document order, and the
 * caption text. Later captions with the same label are renumbered. Returns
 * the caption paragraph.
 *
 * The chapter number (`STYLEREF n \s`) is the count of headings of that level
 * up to the caption, which matches Word when chapter headings are numbered
 * 1, 2, 3 … from the start of the document.
 */
export function insertCaption(
  doc: Docx,
  blockIndex: number,
  position: "above" | "below",
  options: CaptionOptions,
): WmlParagraph {
  const label = options.label.trim();
  if (!label) throw new Error("A caption needs a label.");
  ensureBuiltInStyles(doc, CAPTION_STYLES);
  const formatSwitch = GENERAL_FORMAT_SWITCH[options.numberFormat ?? "decimal"] ?? "ARABIC";
  const chapter = options.chapter;
  const seq = `SEQ ${quoteFieldArg(label)} \\* ${formatSwitch}${chapter ? ` \\s ${chapter.headingLevel}` : ""}`;
  let xml = options.excludeLabel ? "" : textRunXml(`${label} `);
  if (chapter) {
    xml += fieldRunsXml(`STYLEREF ${chapter.headingLevel} \\s`, textRunXml("1"));
    xml += textRunXml(chapter.separator);
  }
  xml += fieldRunsXml(seq, textRunXml("1"));
  if (options.text) xml += textRunXml(options.text);
  const [paragraph] = paragraphsFromXml(`<w:p><w:pPr><w:pStyle w:val="Caption"/></w:pPr>${xml}</w:p>`);
  if (!paragraph) throw new Error("Failed to build the caption paragraph.");
  insertBlocks(doc, position === "above" ? blockIndex : blockIndex + 1, [paragraph]);
  withFlow(doc, (flow) => {
    const fields = scanFields(flow);
    const numbers = sequenceNumbers(doc, flow, fields);
    const chapters = chapterNumbers(doc, flow, chapter?.headingLevel);
    for (const f of fields.toReversed()) {
      const sameLabel =
        f.parsed.type === "SEQ" && (f.parsed.args[0] ?? "").toLowerCase() === label.toLowerCase();
      if (sameLabel && f.begin.ref === f.end.ref) {
        setInlineResult(f, inlinesFromXml(textRunXml(formatFieldNumber(numbers.get(f) ?? 1, f.parsed))));
      } else if (f.parsed.type === "STYLEREF" && f.begin.ref.paragraph === paragraph && chapter) {
        setInlineResult(f, inlinesFromXml(textRunXml(String(chapters.get(paragraph) ?? 1))));
      }
    }
  });
  return paragraph;
}

/** For each paragraph, the number of headings of level ≤ `level` at or before it. */
function chapterNumbers(
  doc: Docx,
  flow: readonly ParagraphRef[],
  level: number | undefined,
): Map<WmlParagraph, number> {
  const out = new Map<WmlParagraph, number>();
  if (level === undefined) return out;
  let count = 0;
  for (const { paragraph } of flow) {
    const l = styleOutlineLevel(doc, paragraphStyleId(paragraph));
    if (l !== undefined && l <= level) count++;
    out.set(paragraph, Math.max(count, 1));
  }
  return out;
}

/** Caption labels: Word's built-in ones plus those the document defines in `w:captions`. */
export function captionLabels(doc: Docx): string[] {
  const captions = findChild(readSettings(doc), "captions");
  const own = (captions?.children ?? [])
    .filter((c): c is XmlElement => c.kind === "element" && c.name.local === "caption")
    .map((c) => wAttr(c, "name") ?? "")
    .filter(Boolean);
  return [...new Set([...CAPTION_LABELS, ...own])];
}

/** Add a caption label (Caption ▸ New Label) to the document's `w:captions` settings. */
export function addCaptionLabel(doc: Docx, name: string): void {
  const label = name.trim();
  if (!label) throw new Error("A caption label cannot be empty.");
  if (captionLabels(doc).some((l) => l.toLowerCase() === label.toLowerCase())) return;
  editSettings(doc, (root) => {
    const existing = findChild(root, "captions");
    const kept = (existing?.children ?? [])
      .filter((c): c is XmlElement => c.kind === "element")
      .map((c) => elementXml(c))
      .join("");
    const caption = `<w:caption w:name="${escapeXml(label)}" w:pos="below"/>`;
    setOrderedChild(root, elementFromXml(`<w:captions>${kept}${caption}</w:captions>`), SETTINGS_ORDER);
  });
}

// --- Index ----------------------------------------------------------------------

export interface IndexEntryOptions {
  /** Main entry text. */
  readonly main: string;
  readonly subentry?: string;
  /** "See …" cross-reference text shown instead of a page number (`\t`). */
  readonly crossReference?: string;
  /** Show a page range ending at this bookmark (`\r`). */
  readonly pageRangeBookmark?: string;
  readonly bold?: boolean;
  readonly italic?: boolean;
}

/** The XE instruction for an index entry. Colons inside entry text are escaped as `\:`. */
export function indexEntryInstruction(entry: IndexEntryOptions): string {
  const escapeColon = (s: string) => s.replace(/:/g, "\\:");
  const text = [entry.main, entry.subentry]
    .filter((s): s is string => !!s && s.trim() !== "")
    .map((s) => escapeColon(s.trim()))
    .join(":");
  if (!text) throw new Error("An index entry needs main entry text.");
  const parts = ["XE", quotedFieldArg(text)];
  if (entry.crossReference) parts.push("\\t", quotedFieldArg(entry.crossReference));
  if (entry.pageRangeBookmark) parts.push("\\r", quoteFieldArg(entry.pageRangeBookmark));
  if (entry.bold) parts.push("\\b");
  if (entry.italic) parts.push("\\i");
  return parts.join(" ");
}

/**
 * Mark an index entry (an XE field, which has no visible result) at a
 * character offset of a paragraph. Returns the inline index of the field.
 */
export function markIndexEntry(
  doc: Docx,
  paragraph: WmlParagraph,
  offset: number,
  entry: IndexEntryOptions,
): number {
  const at = insertInlinesAt(paragraph, offset, inlinesFromXml(resultlessFieldRunsXml(indexEntryInstruction(entry))));
  doc.dirty = true;
  return at;
}

/**
 * Mark All: an XE entry after the first occurrence of `text` in every
 * paragraph that contains it (matching case, as Word does). Returns how many
 * entries were marked.
 */
export function markAllIndexEntries(doc: Docx, text: string, entry: IndexEntryOptions): number {
  if (!text) return 0;
  return withFlow(doc, (flow) => {
    let count = 0;
    for (const { paragraph } of flow) {
      if (paragraphStyleId(paragraph)?.startsWith("Index")) continue;
      // Offsets count run text only, which is what insertInlinesAt expects.
      let runText = "";
      for (const c of paragraph.children) {
        if (c.kind !== "run") continue;
        for (const p of c.pieces) {
          if (p.kind === "text") runText += p.value;
          else if (p.kind === "tab" || p.kind === "break" || p.kind === "noBreakHyphen" || p.kind === "softHyphen") runText += " ";
        }
      }
      const found = runText.indexOf(text);
      if (found < 0) continue;
      insertInlinesAt(paragraph, found + text.length, inlinesFromXml(resultlessFieldRunsXml(indexEntryInstruction(entry))));
      count++;
    }
    return count;
  });
}

export interface IndexOptions {
  /** Subentries on their own indented lines (default) or run in after the main entry. */
  readonly type?: "indented" | "runIn";
  /** Number of columns (1-4). Default 2. */
  readonly columns?: number;
  /** Right-align page numbers on a tab. Default false. */
  readonly rightAlignPageNumbers?: boolean;
  readonly tabLeader?: TabLeader;
  /** Letter headings ("A", "B" …) before each group. */
  readonly headings?: boolean;
  /** Language (LCID) used for sorting, written as `\z`. Default 1033 (English US). */
  readonly languageId?: number;
  readonly pageOf?: PageNumberProvider;
}

const DEFAULT_LCID = 1033;
const COLUMN_GAP_TWIPS = 720;
const MAX_INDEX_COLUMNS = 4;

/** The INDEX instruction for these options. */
export function indexInstruction(options: IndexOptions = {}): string {
  const parts = ["INDEX"];
  if (options.headings) parts.push("\\h", '"A"');
  if (options.rightAlignPageNumbers) parts.push("\\e", '"\t"');
  if (options.type === "runIn") parts.push("\\r");
  const columns = options.columns ?? 2;
  if (!Number.isInteger(columns) || columns < 1 || columns > MAX_INDEX_COLUMNS) {
    throw new Error(`Index columns must be 1-${MAX_INDEX_COLUMNS}, got ${columns}.`);
  }
  parts.push("\\c", `"${columns}"`);
  parts.push("\\z", `"${options.languageId ?? DEFAULT_LCID}"`);
  return parts.join(" ");
}

/**
 * Insert an index before body block `at`, with its entries computed from the
 * document's XE fields. With more than one column the index sits in its own
 * continuous section, as Word lays it out.
 */
export function insertIndex(doc: Docx, at: number, options: IndexOptions = {}): void {
  ensureBuiltInStyles(doc, INDEX_STYLES);
  const columns = options.columns ?? 2;
  const instruction = indexInstruction(options);
  const leader = options.tabLeader ?? "dot";
  const tabs = options.rightAlignPageNumbers ? tabsXml(leader, columnWidth(doc, columns)) : "";
  const blocks = paragraphsFromXml(
    blockFieldSkeletonXml(instruction, tabs),
  );
  if (columns > 1) wrapInColumnSection(doc, at, blocks, columns);
  insertBlocks(doc, at, blocks);
  updateTables(doc, { types: ["INDEX"], ...(options.pageOf ? { pageOf: options.pageOf } : {}) });
}

function columnWidth(doc: Docx, columns: number): number {
  return Math.floor((textWidthTwips(doc) - (columns - 1) * COLUMN_GAP_TWIPS) / columns);
}

/**
 * Put `blocks` in a continuous section of `columns` columns: the section the
 * blocks are inserted into ends just before them (a copy of its properties on
 * a new empty paragraph), the column section ends with the last block, and
 * the original section continues after it without a page break.
 */
function wrapInColumnSection(doc: Docx, at: number, blocks: WmlParagraph[], columns: number): void {
  ensureBodySectPr(doc);
  const following = followingSectPr(doc, at);
  const copy = (el: XmlElement | undefined): XmlElement =>
    el ? structuredClone(el) : elementFromXml("<w:sectPr/>");
  const before = copy(following);
  const [opener] = paragraphsFromXml("<w:p/>");
  if (opener) {
    opener.pPr = elementFromXml("<w:pPr/>");
    setOrderedChild(opener.pPr, before, P_PR_ORDER);
    blocks.unshift(opener);
  }
  const section = copy(following);
  removeChild(section, "headerReference");
  removeChild(section, "footerReference");
  setOrderedChild(section, elementFromXml('<w:type w:val="continuous"/>'), SECT_PR_ORDER);
  setOrderedChild(section, elementFromXml(`<w:cols w:num="${columns}" w:space="${COLUMN_GAP_TWIPS}"/>`), SECT_PR_ORDER);
  const last = blocks.at(-1);
  if (last) {
    last.pPr ??= elementFromXml("<w:pPr/>");
    setOrderedChild(last.pPr, section, P_PR_ORDER);
  }
  if (following) setOrderedChild(following, elementFromXml('<w:type w:val="continuous"/>'), SECT_PR_ORDER);
}

/** The sectPr of the section containing body block `at` (the next section break at or after it). */
function followingSectPr(doc: Docx, at: number): XmlElement | undefined {
  const blocks = doc.document.body.blocks;
  for (let i = Math.max(at, 0); i < blocks.length; i++) {
    const b = blocks[i] as WmlBlock;
    if (b.kind === "paragraph") {
      const s = childElement(b.pPr, "sectPr");
      if (s) return s;
    }
  }
  return doc.document.body.sectPr;
}

interface IndexPage {
  readonly text: string;
  readonly sortKey: number;
  readonly bold: boolean;
  readonly italic: boolean;
}

interface IndexNode {
  readonly name: string;
  readonly pages: IndexPage[];
  readonly see: string[];
  readonly children: Map<string, IndexNode>;
}

function bookmarkPages(flow: readonly ParagraphRef[], pageOf: (p: WmlParagraph) => number) {
  const starts = new Map<string, { id: string; page: number }>();
  const ends = new Map<string, number>();
  for (const { paragraph } of flow) {
    for (const inline of paragraph.children) {
      if (inline.kind !== "raw") continue;
      const local = inline.node.name.local;
      const id = wAttr(inline.node, "id") ?? "";
      if (local === "bookmarkStart") {
        starts.set(wAttr(inline.node, "name") ?? "", { id, page: pageOf(paragraph) });
      } else if (local === "bookmarkEnd") {
        ends.set(id, pageOf(paragraph));
      }
    }
  }
  return (name: string): [number, number] | undefined => {
    const s = starts.get(name);
    if (!s) return undefined;
    return [s.page, ends.get(s.id) ?? s.page];
  };
}

function splitEntry(text: string): string[] {
  // `:` separates levels; `\:` is a literal colon.
  return text
    .split(/(?<!\\):/)
    .map((s) => s.replace(/\\:/g, ":").trim())
    .filter(Boolean);
}

function indexResult(
  doc: Docx,
  field: FlowField,
  fields: readonly FlowField[],
  pageOf: (p: WmlParagraph) => number,
): WmlParagraph[] {
  const flow = paragraphFlow(doc.document.body);
  const rangeOf = bookmarkPages(flow, pageOf);
  const root = new Map<string, IndexNode>();
  const rangeSep = fieldSwitch(field.parsed, "g")?.arg ?? "–";
  for (const f of fields) {
    if (f.parsed.type !== "XE") continue;
    const levels = splitEntry(f.parsed.args[0] ?? "");
    if (levels.length === 0) continue;
    let map = root;
    let node: IndexNode | undefined;
    for (const level of levels) {
      const key = level.toLowerCase();
      node = map.get(key);
      if (!node) {
        node = { name: level, pages: [], see: [], children: new Map() };
        map.set(key, node);
      }
      map = node.children;
    }
    if (!node) continue;
    const see = fieldSwitch(f.parsed, "t")?.arg;
    if (see) {
      if (!node.see.includes(see)) node.see.push(see);
      continue;
    }
    const range = fieldSwitch(f.parsed, "r")?.arg;
    const span = range ? rangeOf(range) : undefined;
    const page = pageOf(f.begin.ref.paragraph);
    const text = span && span[1] > span[0] ? `${span[0]}${rangeSep}${span[1]}` : String(span?.[0] ?? page);
    if (node.pages.some((p) => p.text === text)) continue;
    node.pages.push({
      text,
      sortKey: span?.[0] ?? page,
      bold: !!fieldSwitch(f.parsed, "b"),
      italic: !!fieldSwitch(f.parsed, "i"),
    });
  }
  if (root.size === 0) {
    return paragraphsFromXml(`<w:p>${textRunXml("No index entries found.", "<w:b/><w:bCs/><w:noProof/>")}</w:p>`);
  }
  const entrySep = fieldSwitch(field.parsed, "e")?.arg ?? ", ";
  const pageSep = fieldSwitch(field.parsed, "l")?.arg ?? ", ";
  const runIn = !!fieldSwitch(field.parsed, "r");
  const headings = fieldSwitch(field.parsed, "h");
  const tabs = entrySep === "\t" ? `${tabsXml(existingLeaderOr(field, "dot"), columnWidth(doc, Number(fieldSwitch(field.parsed, "c")?.arg ?? 1) || 1))}` : "";
  const sorted = (m: Map<string, IndexNode>) =>
    [...m.values()].toSorted((a, b) => a.name.localeCompare(b.name, undefined, { sensitivity: "base" }));
  const pagesXml = (node: IndexNode): string => {
    const pages = node.pages.toSorted((a, b) => a.sortKey - b.sortKey);
    const parts = pages.map((p) => {
      const rPr = `${p.bold ? "<w:b/><w:bCs/>" : ""}${p.italic ? "<w:i/><w:iCs/>" : ""}`;
      return textRunXml(p.text, rPr);
    });
    let xml = parts.join(textRunXml(pageSep));
    if (node.see.length) {
      const see = node.see.map((s) => `${textRunXml("See ", "<w:i/><w:iCs/>")}${textRunXml(s)}`).join(textRunXml("; "));
      xml += (xml ? textRunXml(". ") : "") + see;
    }
    return xml ? textRunXml(node.pages.length || entrySep === "\t" ? entrySep : ". ") + xml : "";
  };
  const out: string[] = [];
  let letter = "";
  for (const node of sorted(root)) {
    if (headings) {
      const first = node.name[0]?.toUpperCase() ?? "";
      if (first !== letter) {
        letter = first;
        out.push(`<w:p><w:pPr><w:pStyle w:val="IndexHeading"/></w:pPr>${textRunXml(letter)}</w:p>`);
      }
    }
    if (runIn) {
      const subs = sorted(node.children)
        .map((c) => `${textRunXml(c.name)}${pagesXml(c)}`)
        .join(textRunXml("; "));
      const sep = subs ? textRunXml(node.pages.length ? "; " : ": ") : "";
      out.push(`<w:p><w:pPr><w:pStyle w:val="Index1"/>${tabs}</w:pPr>${textRunXml(node.name)}${pagesXml(node)}${sep}${subs}</w:p>`);
      continue;
    }
    const emit = (n: IndexNode, level: number): void => {
      out.push(`<w:p><w:pPr><w:pStyle w:val="Index${level}"/>${tabs}</w:pPr>${textRunXml(n.name)}${pagesXml(n)}</w:p>`);
      for (const c of sorted(n.children)) emit(c, Math.min(level + 1, MAX_LEVEL));
    };
    emit(node, 1);
  }
  return paragraphsFromXml(out.join(""));
}

function existingLeaderOr(field: FlowField, fallback: TabLeader): TabLeader {
  const tab = childElement(childElement(field.begin.ref.paragraph.pPr, "tabs"), "tab");
  return tab ? existingLeader(field.begin.ref.paragraph) : fallback;
}

// --- Table of authorities ---------------------------------------------------------------

/** Word's table-of-authorities categories (TA `\c`), 1-based. */
export const AUTHORITY_CATEGORIES = [
  "Cases",
  "Statutes",
  "Other Authorities",
  "Rules",
  "Treatises",
  "Regulations",
  "Constitutional Provisions",
] as const;

export interface CitationMarkOptions {
  /** The full citation as it appears in the table (`\l`). */
  readonly longCitation: string;
  /** The short form later references use (`\s`). Defaults to the long citation. */
  readonly shortCitation?: string;
  /** Category number, 1-based into {@link AUTHORITY_CATEGORIES}. */
  readonly category: number;
}

/** Mark a citation for the table of authorities (a TA field) at a character offset. */
export function markAuthorityCitation(
  doc: Docx,
  paragraph: WmlParagraph,
  offset: number,
  options: CitationMarkOptions,
): number {
  const long = options.longCitation.trim();
  if (!long) throw new Error("A citation needs its long form.");
  if (!Number.isInteger(options.category) || options.category < 1 || options.category > 16) {
    throw new Error(`TA category must be 1-16, got ${options.category}.`);
  }
  const short = options.shortCitation?.trim() || long;
  const instruction = `TA \\l ${quotedFieldArg(long)} \\s ${quotedFieldArg(short)} \\c ${options.category}`;
  const at = insertInlinesAt(paragraph, offset, inlinesFromXml(resultlessFieldRunsXml(instruction)));
  doc.dirty = true;
  return at;
}

export interface TableOfAuthoritiesOptions {
  /** Category to list, 1-based; `"all"` inserts one table per category in use. */
  readonly category: number | "all";
  /** "Use passim": five or more pages print as "passim". Default true. */
  readonly passim?: boolean;
  readonly tabLeader?: TabLeader;
  readonly pageOf?: PageNumberProvider;
}

// Word replaces the page list with "passim" from five references on.
const PASSIM_THRESHOLD = 5;

/** Insert a table of authorities (TOA field per category, with heading) before body block `at`. */
export function insertTableOfAuthorities(doc: Docx, at: number, options: TableOfAuthoritiesOptions): void {
  ensureBuiltInStyles(doc, TOA_STYLES);
  let categories: number[];
  if (options.category === "all") {
    const used = withFlow(doc, (flow) =>
      scanFields(flow)
        .filter((f) => f.parsed.type === "TA")
        .map((f) => Number(fieldSwitch(f.parsed, "c")?.arg ?? 0))
        .filter((c) => c >= 1),
    );
    categories = [...new Set(used)].toSorted((a, b) => a - b);
    if (categories.length === 0) categories = [1];
  } else {
    categories = [options.category];
  }
  const leader = options.tabLeader ?? "dot";
  const xml = categories
    .map((c) => {
      const instruction = `TOA \\h \\c "${c}"${options.passim ?? true ? " \\p" : ""}`;
      return blockFieldSkeletonXml(instruction, tabsXml(leader, textWidthTwips(doc)));
    })
    .join("");
  insertBlocks(doc, at, paragraphsFromXml(xml));
  updateTables(doc, { types: ["TOA"], ...(options.pageOf ? { pageOf: options.pageOf } : {}) });
}

function toaResult(
  doc: Docx,
  field: FlowField,
  fields: readonly FlowField[],
  pageOf: (p: WmlParagraph) => number,
): WmlParagraph[] {
  const category = Number(fieldSwitch(field.parsed, "c")?.arg ?? 1);
  const passim = !!fieldSwitch(field.parsed, "p");
  // Long citations by their short form, then the pages each authority is cited on.
  const longOf = new Map<string, { long: string; category: number }>();
  for (const f of fields) {
    if (f.parsed.type !== "TA") continue;
    const long = fieldSwitch(f.parsed, "l")?.arg;
    const short = fieldSwitch(f.parsed, "s")?.arg ?? long;
    if (long && short) {
      longOf.set(short.toLowerCase(), { long, category: Number(fieldSwitch(f.parsed, "c")?.arg ?? 0) });
    }
  }
  const pages = new Map<string, Set<number>>();
  for (const f of fields) {
    if (f.parsed.type !== "TA") continue;
    const key = (fieldSwitch(f.parsed, "s")?.arg ?? fieldSwitch(f.parsed, "l")?.arg ?? "").toLowerCase();
    const authority = longOf.get(key);
    if (!authority || authority.category !== category) continue;
    const set = pages.get(key) ?? new Set<number>();
    set.add(pageOf(f.begin.ref.paragraph));
    pages.set(key, set);
  }
  const out: string[] = [];
  if (fieldSwitch(field.parsed, "h")) {
    const name = AUTHORITY_CATEGORIES[category - 1] ?? `Category ${category}`;
    out.push(`<w:p><w:pPr><w:pStyle w:val="TOAHeading"/></w:pPr>${textRunXml(name)}</w:p>`);
  }
  const entries = [...pages.entries()].toSorted((a, b) =>
    (longOf.get(a[0])?.long ?? "").localeCompare(longOf.get(b[0])?.long ?? "", undefined, { sensitivity: "base" }),
  );
  if (entries.length === 0) {
    out.push(`<w:p>${textRunXml("No table of authorities entries found.", "<w:b/><w:bCs/><w:noProof/>")}</w:p>`);
    return paragraphsFromXml(out.join(""));
  }
  const tabs = tabsXml(existingLeaderOr(field, "dot"), textWidthTwips(doc));
  for (const [key, set] of entries) {
    const list = [...set].toSorted((a, b) => a - b);
    const pageText = passim && list.length >= PASSIM_THRESHOLD ? "passim" : list.join(", ");
    out.push(`<w:p><w:pPr><w:pStyle w:val="TableofAuthorities"/>${tabs}</w:pPr>${textRunXml(longOf.get(key)?.long ?? key)}<w:r><w:tab/></w:r>${textRunXml(pageText)}</w:p>`);
  }
  return paragraphsFromXml(out.join(""));
}

// --- Footnotes and endnotes ---------------------------------------------------------------

/** Where notes are placed (ST_FtnPos / ST_EdnPos). */
export type NotePosition = "pageBottom" | "beneathText" | "sectEnd" | "docEnd";
/** When numbering restarts (ST_RestartNumber). */
export type NoteRestart = "continuous" | "eachSect" | "eachPage";

/** Footnote / endnote numbering and placement (`w:footnotePr` / `w:endnotePr`). */
export interface NoteProperties {
  readonly position?: NotePosition;
  readonly numberFormat?: NumberingFormat;
  readonly startAt?: number;
  readonly restart?: NoteRestart;
}

const NOTE_DEFAULTS: Readonly<Record<NoteKind, Required<NoteProperties>>> = {
  footnote: { position: "pageBottom", numberFormat: "decimal", startAt: 1, restart: "continuous" },
  endnote: { position: "docEnd", numberFormat: "lowerRoman", startAt: 1, restart: "continuous" },
};

const NOTE_POSITIONS: Readonly<Record<NoteKind, ReadonlySet<string>>> = {
  footnote: new Set(["pageBottom", "beneathText", "sectEnd", "docEnd"]),
  endnote: new Set(["sectEnd", "docEnd"]),
};

function notePrLocal(kind: NoteKind): "footnotePr" | "endnotePr" {
  return kind === "footnote" ? "footnotePr" : "endnotePr";
}

function readNotePr(el: XmlElement | undefined): NoteProperties {
  if (!el) return {};
  const pos = wAttr(findChild(el, "pos"), "val");
  const fmt = wAttr(findChild(el, "numFmt"), "val");
  const start = wAttr(findChild(el, "numStart"), "val");
  const restart = wAttr(findChild(el, "numRestart"), "val");
  return {
    ...(pos ? { position: pos as NotePosition } : {}),
    ...(fmt ? { numberFormat: fmt as NumberingFormat } : {}),
    ...(start ? { startAt: Number(start) } : {}),
    ...(restart ? { restart: restart as NoteRestart } : {}),
  };
}

/** The section properties in order: every paragraph-level sectPr, then the body's. */
function sectionList(doc: Docx): XmlElement[] {
  const out: XmlElement[] = [];
  for (const block of doc.document.body.blocks) {
    if (block.kind !== "paragraph") continue;
    const s = childElement(block.pPr, "sectPr");
    if (s) out.push(s);
  }
  if (doc.document.body.sectPr) out.push(doc.document.body.sectPr);
  return out;
}

/** The effective note properties of a section (its own, over the document's, over Word's defaults). */
export function noteProperties(doc: Docx, kind: NoteKind, section?: number): Required<NoteProperties> {
  const docLevel = readNotePr(findChild(readSettings(doc), notePrLocal(kind)));
  const sections = sectionList(doc);
  const sect = sections[section ?? sections.length - 1];
  const own = readNotePr(childElement(sect, notePrLocal(kind)));
  return { ...NOTE_DEFAULTS[kind], ...docLevel, ...own };
}

function notePrXml(kind: NoteKind, props: NoteProperties, extra = ""): string {
  const local = notePrLocal(kind);
  const parts: string[] = [];
  if (props.position) parts.push(`<w:pos w:val="${props.position}"/>`);
  if (props.numberFormat) parts.push(`<w:numFmt w:val="${props.numberFormat}"/>`);
  if (props.startAt !== undefined) parts.push(`<w:numStart w:val="${props.startAt}"/>`);
  if (props.restart) parts.push(`<w:numRestart w:val="${props.restart}"/>`);
  return `<w:${local}>${parts.join("")}${extra}</w:${local}>`;
}

/**
 * The Footnote and Endnote dialog's Apply: write numbering and placement to
 * the document settings and every section ("Whole document"), or to one
 * section's properties (0-based index, "This section").
 */
export function setNoteProperties(
  doc: Docx,
  kind: NoteKind,
  props: NoteProperties,
  scope: "document" | { readonly section: number } = "document",
): void {
  if (props.position && !NOTE_POSITIONS[kind].has(props.position)) {
    throw new Error(`${props.position} is not a valid ${kind} position.`);
  }
  if (props.startAt !== undefined && (!Number.isInteger(props.startAt) || props.startAt < 0)) {
    throw new Error(`Start at must be a whole number, got ${props.startAt}.`);
  }
  ensureBodySectPr(doc);
  const sections = sectionList(doc);
  if (scope === "document") {
    editSettings(doc, (root) => {
      const existing = findChild(root, notePrLocal(kind));
      const separators = (existing?.children ?? []).filter(
        (c): c is XmlElement => c.kind === "element" && (c.name.local === "footnote" || c.name.local === "endnote"),
      );
      const merged = { ...readNotePr(existing), ...props };
      const el = elementFromXml(notePrXml(kind, merged));
      (el.children as XmlElement[]).push(...separators);
      setOrderedChild(root, el, SETTINGS_ORDER);
    });
    for (const s of sections) applySectionNotePr(s, kind, props);
  } else {
    const s = sections[scope.section];
    if (!s) throw new Error(`No section ${scope.section}.`);
    applySectionNotePr(s, kind, props);
  }
  doc.dirty = true;
}

function applySectionNotePr(sectPr: XmlElement, kind: NoteKind, props: NoteProperties): void {
  const merged = { ...readNotePr(childElement(sectPr, notePrLocal(kind))), ...props };
  const el = elementFromXml(notePrXml(kind, merged));
  // Keep each child in CT_FtnProps order.
  const ordered = elementFromXml(`<w:${notePrLocal(kind)}/>`);
  for (const c of el.children) if (c.kind === "element") setOrderedChild(ordered, c, NOTE_PR_ORDER);
  setOrderedChild(sectPr, ordered, SECT_PR_ORDER);
}

/** A note reference in the body, in document order, with the mark it displays. */
export interface NoteMark {
  readonly kind: NoteKind;
  readonly id: number;
  /** The displayed reference mark ("1", "iv", "*", or a custom mark). */
  readonly mark: string;
  readonly paragraph: WmlParagraph;
}

function referenceOf(inline: WmlInline, kind: NoteKind): { id: number; custom?: string } | undefined {
  const local = kind === "footnote" ? "footnoteReference" : "endnoteReference";
  if (inline.kind === "run") {
    const piece = inline.pieces.find((p) => p.kind === "raw" && p.node.name.local === local);
    if (!piece || piece.kind !== "raw") return undefined;
    const id = Number(wAttr(piece.node, "id"));
    const customFlag = wAttr(piece.node, "customMarkFollows");
    if (customFlag === "1" || customFlag === "true" || customFlag === "on") {
      const custom = inline.pieces
        .map((p) => (p.kind === "text" ? p.value : p.kind === "symbol" ? String.fromCharCode(Number.parseInt(p.char, 16)) : ""))
        .join("");
      return { id, custom };
    }
    return { id };
  }
  if (inline.node.name.local === "r") {
    const el = childElement(inline.node, local);
    if (el) return { id: Number(wAttr(el, "id")) };
  }
  return undefined;
}

/**
 * The reference marks of every footnote or endnote, in document order,
 * numbered by the effective `numFmt` / `numStart` / `numRestart` of the
 * section each reference sits in. Custom marks keep their text and do not
 * advance the numbering. `eachPage` restarts use `pageOf` (estimated pages
 * otherwise).
 */
export function noteMarks(doc: Docx, kind: NoteKind, pageOf?: PageNumberProvider): NoteMark[] {
  const pages = pageLookup(doc, pageOf);
  const out: NoteMark[] = [];
  const sections = sectionList(doc);
  let section = 0;
  let count = 0;
  let lastPage = 0;
  let props = noteProperties(doc, kind, 0);
  for (const ref of paragraphFlow(doc.document.body)) {
    for (const inline of ref.paragraph.children) {
      const r = referenceOf(inline, kind);
      if (!r) continue;
      if (props.restart === "eachPage") {
        const page = pages(ref.paragraph);
        if (page !== lastPage) count = 0;
        lastPage = page;
      }
      if (r.custom !== undefined) {
        out.push({ kind, id: r.id, mark: r.custom, paragraph: ref.paragraph });
        continue;
      }
      out.push({ kind, id: r.id, mark: formatNumber(props.startAt + count, props.numberFormat), paragraph: ref.paragraph });
      count++;
    }
    if (ref.list === doc.document.body.blocks && childElement(ref.paragraph.pPr, "sectPr")) {
      section = Math.min(section + 1, sections.length - 1);
      props = noteProperties(doc, kind, section);
      if (props.restart === "eachSect") count = 0;
    }
  }
  return out;
}

export interface InsertNoteOptions {
  /** The note's text (Word leaves it empty and moves the caret into the note). */
  readonly text?: string;
  /** A custom reference mark instead of automatic numbering. */
  readonly customMark?: string;
}

/**
 * Insert a footnote or endnote reference at a character offset of a body
 * paragraph, and the note itself (with Word's built-in Footnote Text /
 * Footnote Reference styles). Returns the note id.
 */
export function insertNote(
  doc: Docx,
  paragraph: WmlParagraph,
  offset: number,
  kind: NoteKind,
  options: InsertNoteOptions = {},
): number {
  const part = ensureNotesPart(doc, kind);
  ensureBuiltInStyles(doc, NOTE_STYLES[kind]);
  let id = 1;
  const used = new Set(part.footnotes.map((n) => Number(wAttr(n, "id"))));
  while (used.has(id)) id++;
  const refStyle = kind === "footnote" ? "FootnoteReference" : "EndnoteReference";
  const textStyle = kind === "footnote" ? "FootnoteText" : "EndnoteText";
  const custom = options.customMark?.trim();
  const refRunInNote = custom
    ? textRunXml(custom, `<w:rStyle w:val="${refStyle}"/>`)
    : `<w:r><w:rPr><w:rStyle w:val="${refStyle}"/></w:rPr><w:${kind}Ref/></w:r>`;
  const body = textRunXml(` ${options.text ?? ""}`);
  part.footnotes.push(
    elementFromXml(`<w:${kind} w:id="${id}"><w:p><w:pPr><w:pStyle w:val="${textStyle}"/></w:pPr>${refRunInNote}${body}</w:p></w:${kind}>`),
  );
  markNotesDirty(doc, kind);
  const reference = custom
    ? `<w:r><w:rPr><w:rStyle w:val="${refStyle}"/></w:rPr><w:${kind}Reference w:customMarkFollows="1" w:id="${id}"/><w:t>${escapeXml(custom)}</w:t></w:r>`
    : `<w:r><w:rPr><w:rStyle w:val="${refStyle}"/></w:rPr><w:${kind}Reference w:id="${id}"/></w:r>`;
  insertInlinesAt(paragraph, offset, inlinesFromXml(reference));
  doc.dirty = true;
  return id;
}

export type NoteConversion = "footnotesToEndnotes" | "endnotesToFootnotes" | "swap";

/**
 * Convert Notes: move every footnote to the endnotes (or the reverse, or
 * swap both), renumbering ids in the target part and retargeting the body's
 * references. Returns the number of notes converted.
 */
export function convertNotes(doc: Docx, conversion: NoteConversion): number {
  const plan: Array<[NoteKind, NoteKind]> =
    conversion === "footnotesToEndnotes"
      ? [["footnote", "endnote"]]
      : conversion === "endnotesToFootnotes"
        ? [["endnote", "footnote"]]
        : [
            ["footnote", "endnote"],
            ["endnote", "footnote"],
          ];
  // Collect both sides before moving anything, so a swap does not move a
  // note twice.
  const moves = plan.map(([from, to]) => {
    const part = notesPartOf(doc, from);
    const notes = (part?.footnotes ?? []).filter((n) => !wAttr(n, "type") || wAttr(n, "type") === "normal");
    return { from, to, notes };
  });
  const remap = new Map<string, { kind: NoteKind; id: number }>();
  for (const { from, to, notes } of moves) {
    if (notes.length === 0) continue;
    const source = notesPartOf(doc, from);
    const target = ensureNotesPart(doc, to);
    ensureBuiltInStyles(doc, NOTE_STYLES[to]);
    const used = new Set(target.footnotes.map((n) => Number(wAttr(n, "id"))));
    for (const note of notes) {
      let id = 1;
      while (used.has(id)) id++;
      used.add(id);
      remap.set(`${from}:${wAttr(note, "id")}`, { kind: to, id });
      target.footnotes.push(retargetNote(note, from, to, id));
    }
    if (source) source.footnotes = source.footnotes.filter((n) => !notes.includes(n));
    markNotesDirty(doc, from);
    markNotesDirty(doc, to);
  }
  if (remap.size === 0) return 0;
  withFlow(doc, (flow) => {
    for (const { paragraph } of flow) {
      paragraph.children = paragraph.children.map((inline) => retargetReference(inline, remap));
    }
  });
  return remap.size;
}

const CAPITAL = (s: string) => s[0]?.toUpperCase() + s.slice(1);

/**
 * A copy of a note as the other kind: element names (`w:footnote` →
 * `w:endnote`, `w:footnoteRef` → `w:endnoteRef`) and the built-in note styles
 * change, everything else (including foreign-namespace attributes) is kept.
 */
function retargetNote(note: XmlElement, from: NoteKind, to: NoteKind, id: number): XmlElement {
  const styleFrom = new Map([
    [`${CAPITAL(from)}Text`, `${CAPITAL(to)}Text`],
    [`${CAPITAL(from)}Reference`, `${CAPITAL(to)}Reference`],
  ]);
  const rename = new Map([
    [from, to],
    [`${from}Ref`, `${to}Ref`],
  ]);
  const transform = (el: XmlElement, top: boolean): XmlElement => {
    const local = el.name.uri === WML_NS ? (rename.get(el.name.local) ?? el.name.local) : el.name.local;
    const isStyleRef = el.name.local === "pStyle" || el.name.local === "rStyle";
    const attrs = el.attrs.map((a) => {
      if (top && a.name.local === "id") return { ...a, value: String(id) };
      if (isStyleRef && a.name.local === "val") return { ...a, value: styleFrom.get(a.value) ?? a.value };
      return a;
    });
    return {
      ...el,
      name: { ...el.name, local },
      attrs,
      children: el.children.map((c) => (c.kind === "element" ? transform(c, false) : c)),
    };
  };
  return transform(note, true);
}

function retargetReference(
  inline: WmlInline,
  remap: ReadonlyMap<string, { kind: NoteKind; id: number }>,
): WmlInline {
  for (const kind of ["footnote", "endnote"] as const) {
    const ref = referenceOf(inline, kind);
    if (!ref) continue;
    const target = remap.get(`${kind}:${ref.id}`);
    if (!target) return inline;
    const T = CAPITAL(target.kind);
    const custom = ref.custom !== undefined;
    const xml = custom
      ? `<w:r><w:rPr><w:rStyle w:val="${T}Reference"/></w:rPr><w:${target.kind}Reference w:customMarkFollows="1" w:id="${target.id}"/><w:t>${escapeXml(ref.custom ?? "")}</w:t></w:r>`
      : `<w:r><w:rPr><w:rStyle w:val="${T}Reference"/></w:rPr><w:${target.kind}Reference w:id="${target.id}"/></w:r>`;
    return inlinesFromXml(xml)[0] ?? inline;
  }
  return inline;
}

// --- Citations and bibliography ------------------------------------------------------------

const CUSTOM_XML_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml";
const CUSTOM_XML_PROPS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps";
const CUSTOM_XML_PROPS_CT = "application/vnd.openxmlformats-officedocument.customXmlProperties+xml";
const CUSTOM_XML_CT = "application/xml";

/** The part name of the document's `<b:Sources>` custom XML part, if any. */
function bibliographyPartName(doc: Docx): string | undefined {
  const rels = partRelationships(doc.opc, doc.partName);
  for (const rel of relationshipsByType(rels, CUSTOM_XML_REL)) {
    if (rel.targetMode === "External") continue;
    const name = resolveFromWord(rel.target);
    const part = getPart(doc.opc, name);
    if (!part) continue;
    const head = new TextDecoder("utf-8").decode(part.data.subarray(0, 512));
    if (head.includes(BIBLIOGRAPHY_NS)) return name;
  }
  return undefined;
}

function resolveFromWord(target: string): string {
  if (target.startsWith("/")) return target;
  const segments = ["word", ...target.split("/")];
  const out: string[] = [];
  for (const s of segments) {
    if (s === "..") out.pop();
    else if (s !== "." && s !== "") out.push(s);
  }
  return `/${out.join("/")}`;
}

interface BibliographyState {
  readonly style: CitationStyle;
  readonly sources: readonly BibliographySource[];
}

function readBibliography(doc: Docx): BibliographyState {
  const name = bibliographyPartName(doc);
  const part = name ? getPart(doc.opc, name) : undefined;
  if (!part) return { style: "APA", sources: [] };
  return parseSources(new TextDecoder("utf-8").decode(part.data));
}

/** The document's bibliography sources (Manage Sources ▸ Current List). */
export function bibliographySources(doc: Docx): BibliographySource[] {
  return [...readBibliography(doc).sources];
}

/** The citation style shown in References ▸ Style. Default APA. */
export function bibliographyStyle(doc: Docx): CitationStyle {
  return readBibliography(doc).style;
}

function writeBibliography(doc: Docx, state: BibliographyState): void {
  const tags = new Set<string>();
  for (const s of state.sources) {
    if (!s.tag.trim()) throw new Error("Every source needs a tag.");
    if (tags.has(s.tag)) throw new Error(`Duplicate source tag ${JSON.stringify(s.tag)}.`);
    tags.add(s.tag);
  }
  const data = new TextEncoder().encode(writeSources({ style: state.style, sources: [...state.sources] }));
  const existing = bibliographyPartName(doc);
  const part = existing ? getPart(doc.opc, existing) : undefined;
  if (part) {
    part.data = data;
    doc.dirty = true;
    return;
  }
  let n = 1;
  while (hasPart(doc.opc, `/customXml/item${n}.xml`) || hasPart(doc.opc, `/customXml/itemProps${n}.xml`)) n++;
  const itemName = `/customXml/item${n}.xml`;
  const propsName = `/customXml/itemProps${n}.xml`;
  addPart(doc.opc, { name: itemName, contentType: CUSTOM_XML_CT, data });
  addPart(doc.opc, {
    name: propsName,
    contentType: CUSTOM_XML_PROPS_CT,
    data: new TextEncoder().encode(
      `<?xml version="1.0" encoding="UTF-8" standalone="no"?><ds:datastoreItem ds:itemID="{${newGuid()}}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"><ds:schemaRefs><ds:schemaRef ds:uri="${BIBLIOGRAPHY_NS}"/></ds:schemaRefs></ds:datastoreItem>`,
    ),
  });
  addRelationship(partRelationships(doc.opc, itemName), { type: CUSTOM_XML_PROPS_REL, target: `itemProps${n}.xml` });
  addRelationship(partRelationships(doc.opc, doc.partName), { type: CUSTOM_XML_REL, target: `../customXml/item${n}.xml` });
  doc.dirty = true;
}

function newGuid(): string {
  return crypto.randomUUID().toUpperCase();
}

/**
 * Replace the document's source list (Manage Sources ▸ Current List), and
 * refresh every citation and bibliography. Tags must be unique.
 */
export function setBibliographySources(doc: Docx, sources: readonly BibliographySource[]): void {
  writeBibliography(doc, { style: bibliographyStyle(doc), sources });
  updateTables(doc, { types: ["CITATION", "BIBLIOGRAPHY"] });
}

/** Switch the citation style (References ▸ Style) and reformat every citation and bibliography. */
export function setBibliographyStyle(doc: Docx, style: CitationStyle): void {
  if (!Object.hasOwn(CITATION_STYLES, style)) throw new Error(`Unknown citation style ${style}.`);
  writeBibliography(doc, { style, sources: bibliographySources(doc) });
  updateTables(doc, { types: ["CITATION", "BIBLIOGRAPHY"] });
}

/** Word's tag for a new source: three letters of the first author (or title) and a two-digit year, made unique. */
export function suggestSourceTag(existing: readonly BibliographySource[], source: Omit<BibliographySource, "tag">): string {
  const lead = source.corporateAuthor ?? source.authors?.[0]?.last ?? source.fields?.Title ?? "Src";
  const stem = lead.replace(/[^\p{L}\p{N}]/gu, "").slice(0, 3) || "Src";
  const year = (source.fields?.Year ?? "").replace(/\D/g, "").slice(-2);
  const base = `${stem}${year}`;
  const taken = new Set(existing.map((s) => s.tag));
  if (!taken.has(base)) return base;
  let i = 1;
  while (taken.has(`${base}${i}`)) i++;
  return `${base}${i}`;
}

/** Word's language id for citations (`\l`); English (United States). */
const CITATION_LCID = 1033;

/** The CITATION instruction for one or more sources (`\m` adds further sources). */
export function citationInstruction(tags: readonly string[], options: CitationOptions = {}): string {
  const [first, ...more] = tags;
  if (!first) throw new Error("A citation needs a source tag.");
  const parts = ["CITATION", quoteFieldArg(first), "\\l", String(CITATION_LCID)];
  if (options.pages) parts.push("\\p", quoteFieldArg(options.pages));
  if (options.volume) parts.push("\\v", quoteFieldArg(options.volume));
  if (options.prefix) parts.push("\\f", quoteFieldArg(options.prefix));
  if (options.suffix) parts.push("\\s", quoteFieldArg(options.suffix));
  if (options.suppressAuthor) parts.push("\\n");
  if (options.suppressYear) parts.push("\\y");
  if (options.suppressTitle) parts.push("\\t");
  for (const tag of more) parts.push("\\m", quoteFieldArg(tag));
  return parts.join(" ");
}

/**
 * Insert a citation (CITATION field with its formatted result) at a
 * character offset of a paragraph. The source must be in the document's
 * source list.
 */
export function insertCitation(
  doc: Docx,
  paragraph: WmlParagraph,
  offset: number,
  tag: string,
  options: CitationOptions = {},
): void {
  if (!bibliographySources(doc).some((s) => s.tag === tag)) {
    throw new Error(`No source with tag ${JSON.stringify(tag)} in the document.`);
  }
  insertInlinesAt(paragraph, offset, inlinesFromXml(fieldRunsXml(citationInstruction([tag], options), textRunXml(`(${tag})`))));
  doc.dirty = true;
  updateTables(doc, { types: ["CITATION", "BIBLIOGRAPHY"] });
}

/** Word's Bibliography gallery titles. */
export const BIBLIOGRAPHY_TITLES = ["Bibliography", "References", "Works Cited"] as const;

/**
 * Insert a bibliography before body block `at`: a Heading 1 title (omit for
 * Word's "Insert Bibliography") and the BIBLIOGRAPHY field with its entries.
 */
export function insertBibliography(doc: Docx, at: number, options: { readonly title?: string } = {}): void {
  ensureBuiltInStyles(doc, BIBLIOGRAPHY_STYLES);
  let xml = "";
  if (options.title) {
    ensureHeadingStyles(doc, 1);
    xml += `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>${textRunXml(options.title)}</w:p>`;
  }
  xml += blockFieldSkeletonXml("BIBLIOGRAPHY");
  insertBlocks(doc, at, paragraphsFromXml(xml));
  updateTables(doc, { types: ["BIBLIOGRAPHY"] });
}

/** Tag → number by first citation, for numeric styles. */
function citationOrder(fields: readonly FlowField[]): Map<string, number> {
  const out = new Map<string, number>();
  for (const f of fields) {
    if (f.parsed.type !== "CITATION") continue;
    const tags = [f.parsed.args[0], ...f.parsed.switches.filter((s) => s.name.toLowerCase() === "m").map((s) => s.arg)];
    for (const tag of tags) if (tag && !out.has(tag)) out.set(tag, out.size + 1);
  }
  return out;
}

function citationOptionsOf(parsed: FieldInstruction): CitationOptions {
  const arg = (n: string) => fieldSwitch(parsed, n)?.arg;
  const pages = arg("p");
  const volume = arg("v");
  const prefix = arg("f");
  const suffix = arg("s");
  return {
    ...(pages ? { pages } : {}),
    ...(volume ? { volume } : {}),
    ...(prefix ? { prefix } : {}),
    ...(suffix ? { suffix } : {}),
    ...(fieldSwitch(parsed, "n") ? { suppressAuthor: true } : {}),
    ...(fieldSwitch(parsed, "y") ? { suppressYear: true } : {}),
    ...(fieldSwitch(parsed, "t") ? { suppressTitle: true } : {}),
  };
}

function styledRunsXml(parts: readonly StyledText[]): string {
  return parts.map((p) => textRunXml(p.text, p.italic ? "<w:i/><w:iCs/>" : "")).join("");
}

function citationResult(
  state: BibliographyState | undefined,
  parsed: FieldInstruction,
  numbers: ReadonlyMap<string, number>,
): WmlInline[] {
  const bySource = new Map((state?.sources ?? []).map((s) => [s.tag, s]));
  const tags = [parsed.args[0], ...parsed.switches.filter((s) => s.name.toLowerCase() === "m").map((s) => s.arg)];
  const options = citationOptionsOf(parsed);
  const cited: CitedSource[] = [];
  for (const tag of tags) {
    if (!tag) continue;
    const source = bySource.get(tag) ?? { tag, type: "Misc" as const };
    cited.push({ source, number: numbers.get(tag) ?? 1, options: tag === tags[0] ? options : {} });
  }
  const text = formatCitation(cited, state?.style ?? "APA");
  return inlinesFromXml(textRunXml(text, "<w:noProof/>"));
}

function bibliographyResult(
  state: BibliographyState | undefined,
  numbers: ReadonlyMap<string, number>,
): WmlParagraph[] {
  const sources = state?.sources ?? [];
  if (sources.length === 0) {
    return paragraphsFromXml(`<w:p>${textRunXml("There are no sources in the current document.", "<w:b/><w:bCs/><w:noProof/>")}</w:p>`);
  }
  const entries = formatBibliography(sources, state?.style ?? "APA", numbers);
  return paragraphsFromXml(
    entries
      .map(
        (parts) =>
          `<w:p><w:pPr><w:pStyle w:val="Bibliography"/><w:ind w:left="${HANGING_INDENT_TWIPS}" w:hanging="${HANGING_INDENT_TWIPS}"/></w:pPr>${styledRunsXml(parts)}</w:p>`,
      )
      .join(""),
  );
}
