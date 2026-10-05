/**
 * The paginator: distributes measured blocks over pages, columns and regions
 * the way Word's Print Layout does. It is pure — heights are measured by the
 * caller (the canvas measures rendered DOM; tests inject numbers) — so the
 * page-breaking rules can be unit-tested and the browser only pays for
 * measuring.
 *
 * Rules implemented (ECMA-376 Part 1):
 * - section starts `nextPage` / `oddPage` / `evenPage` / `continuous` /
 *   `nextColumn` (§17.6.22), with a blank page inserted to reach an odd or
 *   even page number, and the columns of a region closed by a continuous
 *   break balanced, as Word does;
 * - `w:pageBreakBefore` (§17.3.1.23) and `w:br w:type="page|column"` inside a
 *   paragraph (§17.3.3.1), given as break offsets;
 * - `w:keepNext` (§17.3.1.15), `w:keepLines` (§17.3.1.14) and
 *   `w:widowControl` (§17.3.1.44) when a paragraph is split by its lines;
 * - table rows split across pages with header rows (`w:tblHeader`,
 *   §17.4.49) repeated; a row itself is never split, which also honours
 *   `w:cantSplit` (§17.4.6);
 * - footnotes take room at the bottom of the page holding their reference;
 * - vertical page alignment (`w:vAlign`, §17.6.23).
 *
 * Everything is in one length unit chosen by the caller (CSS px on the canvas).
 */

export type SectionStart = "nextPage" | "continuous" | "evenPage" | "oddPage" | "nextColumn";
export type PageKind = "first" | "even" | "default";
export type VerticalAlign = "top" | "center" | "both" | "bottom";

export interface PaginatorSection {
  readonly start: SectionStart;
  /** Number of text columns (`w:cols w:num`). */
  readonly columns: number;
  /** `w:titlePg`: the section's first page uses the first-page header/footer. */
  readonly titlePage: boolean;
  /** `w:pgNumType w:start`: page numbering restarts at this number. */
  readonly pageNumberStart?: number | undefined;
  readonly vAlign: VerticalAlign;
  /**
   * Height available to body text on a page of this kind: the space between
   * the margins, less what a header or footer taller than its margin takes.
   */
  bodyHeight(kind: PageKind): number;
}

export interface LayoutRow {
  readonly height: number;
  /** Repeated at the top of each page the table continues on (only leading rows count). */
  readonly header?: boolean;
}

export interface LayoutNote {
  /** Identifies the note to the caller, e.g. `footnote:3`. */
  readonly key: string;
  readonly height: number;
}

export interface LayoutBlock {
  readonly height: number;
  readonly section: number;
  readonly pageBreakBefore?: boolean;
  readonly keepNext?: boolean;
  readonly keepLines?: boolean;
  readonly widowControl?: boolean;
  /**
   * Bottoms of the block's line boxes, measured from the block's top, in
   * order. Read only when the block must be split, so the caller can measure
   * lazily.
   */
  readonly lines?: () => readonly number[];
  /** Breaks forced inside the block, as offsets from its top where the next content starts. */
  readonly breaks?: ReadonlyArray<{ readonly offset: number; readonly kind: "page" | "column" }>;
  /** Table rows; makes the block splittable between rows. */
  readonly rows?: readonly LayoutRow[];
  /** Footnotes referenced from this block. */
  readonly notes?: readonly LayoutNote[];
}

export interface PaginatorOptions {
  /** settings `w:evenAndOddHeaders`: even pages use the even header/footer. */
  readonly evenAndOddHeaders: boolean;
  /** Room the footnote separator takes above the first footnote of a page. */
  readonly noteSeparatorHeight: number;
}

export type Fragment =
  /** The whole block. */
  | { readonly kind: "block"; readonly block: number; readonly height: number }
  /** The part of a paragraph between two offsets from its top. */
  | { readonly kind: "slice"; readonly block: number; readonly from: number; readonly to: number }
  /** Rows `from`..`to` (exclusive) of a table, after `headerRows` repeated header rows. */
  | {
      readonly kind: "rows";
      readonly block: number;
      readonly from: number;
      readonly to: number;
      readonly headerRows: number;
      readonly height: number;
    };

export interface LaidOutColumn {
  readonly fragments: Fragment[];
  height: number;
  /** Space added between fragments for `w:vAlign="both"`. */
  gap: number;
}

export interface Region {
  readonly section: number;
  /** Offset from the top of the page's body area. */
  top: number;
  readonly columns: LaidOutColumn[];
  /** Offset added above the region's content for vertical alignment. */
  offset: number;
}

export interface LaidOutPage {
  readonly section: number;
  readonly kind: PageKind;
  /** The page number shown by PAGE (§17.16.5.47), honouring numbering restarts. */
  readonly number: number;
  /** 1-based index of the page within its section. */
  readonly numberInSection: number;
  /** A page inserted by an odd/even section break, with no content. */
  readonly blank: boolean;
  readonly regions: Region[];
  readonly notes: string[];
  /** Room footnotes take at the bottom of the body area (separator included). */
  noteHeight: number;
}

// Rounding slack so a block measured at 300.0001 still fits 300.
const EPSILON = 0.5;
const MIN_LINES_AT_BREAK = 2;

function fragmentHeight(f: Fragment): number {
  return f.kind === "slice" ? f.to - f.from : f.height;
}

function pageKind(
  section: PaginatorSection,
  numberInSection: number,
  pageNumber: number,
  options: PaginatorOptions,
): PageKind {
  if (section.titlePage && numberInSection === 1) return "first";
  if (options.evenAndOddHeaders && pageNumber % 2 === 0) return "even";
  return "default";
}

/** Re-pack a region's fragments into its columns at the smallest equal height. */
function balance(region: Region): void {
  const count = region.columns.length;
  if (count < 2) return;
  const all = region.columns.flatMap((c) => c.fragments);
  if (all.length < 2) return;
  const heights = all.map(fragmentHeight);
  const fits = (limit: number): boolean => {
    let col = 0;
    let used = 0;
    for (const h of heights) {
      if (used > 0 && used + h > limit + EPSILON) {
        col++;
        used = 0;
      }
      used += h;
    }
    return col < count;
  };
  let lo = Math.max(...heights);
  let hi = heights.reduce((a, b) => a + b, 0);
  // Binary search on the column height; a pixel of precision is plenty.
  while (hi - lo > 1) {
    const mid = (lo + hi) / 2;
    if (fits(mid)) hi = mid;
    else lo = mid;
  }
  for (const column of region.columns) {
    column.fragments.length = 0;
    column.height = 0;
  }
  let col = 0;
  for (const [i, fragment] of all.entries()) {
    const h = heights[i] ?? 0;
    let column = region.columns[col];
    if (column && column.height > 0 && column.height + h > hi + EPSILON && col < count - 1) {
      col++;
      column = region.columns[col];
    }
    if (!column) break;
    column.fragments.push(fragment);
    column.height += h;
  }
}

function regionHeight(region: Region): number {
  return Math.max(0, ...region.columns.map((c) => c.height));
}

/**
 * Lay blocks out into pages. `sections[block.section]` describes each block's
 * section; blocks must be in document order with non-decreasing sections.
 */
export function paginate(
  blocks: readonly LayoutBlock[],
  sections: readonly PaginatorSection[],
  options: PaginatorOptions,
): LaidOutPage[] {
  const pages: LaidOutPage[] = [];
  let page: LaidOutPage | undefined;
  let region: Region | undefined;
  let col = 0;
  let sectionIndex = -1;
  let numberInSection = 0;
  let lastNumber = 0;

  const sectionOf = (index: number): PaginatorSection => {
    const s = sections[index];
    if (!s) throw new Error(`No geometry for section ${index}.`);
    return s;
  };

  const column = (): LaidOutColumn => {
    const c = region?.columns[col];
    if (!c) throw new Error("Paginator has no current column.");
    return c;
  };

  const newRegion = (section: number, top: number): Region => {
    const count = Math.max(1, sectionOf(section).columns);
    const r: Region = {
      section,
      top,
      offset: 0,
      columns: Array.from({ length: count }, () => ({ fragments: [], height: 0, gap: 0 })),
    };
    page?.regions.push(r);
    col = 0;
    return r;
  };

  const openPage = (section: number, blank: boolean): void => {
    const geometry = sectionOf(section);
    numberInSection++;
    const number =
      numberInSection === 1 && geometry.pageNumberStart !== undefined
        ? geometry.pageNumberStart
        : lastNumber + 1;
    lastNumber = number;
    page = {
      section,
      kind: pageKind(geometry, numberInSection, number, options),
      number,
      numberInSection,
      blank,
      regions: [],
      notes: [],
      noteHeight: 0,
    };
    pages.push(page);
    region = blank ? undefined : newRegion(section, 0);
  };

  const bodyHeight = (): number => (page ? sectionOf(page.section).bodyHeight(page.kind) : 0);
  const available = (): number =>
    bodyHeight() - (page?.noteHeight ?? 0) - (region?.top ?? 0) - column().height;
  const columnEmpty = (): boolean => column().fragments.length === 0;
  const pageEmpty = (): boolean =>
    !page || page.regions.every((r) => r.columns.every((c) => c.fragments.length === 0));

  const advanceColumn = (): void => {
    if (region && col < region.columns.length - 1) {
      col++;
      return;
    }
    openPage(sectionIndex, false);
  };

  const notesHeight = (notes: readonly LayoutNote[] | undefined): number => {
    if (!notes?.length || !page) return 0;
    const own = notes.reduce((n, note) => n + note.height, 0);
    return own + (page.notes.length === 0 ? options.noteSeparatorHeight : 0);
  };

  const place = (fragment: Fragment, notes: readonly LayoutNote[] | undefined): void => {
    const c = column();
    c.fragments.push(fragment);
    c.height += fragmentHeight(fragment);
    if (page && notes?.length) {
      page.noteHeight += notesHeight(notes);
      page.notes.push(...notes.map((n) => n.key));
    }
  };

  const startSection = (next: number): void => {
    const geometry = sectionOf(next);
    const previous = sectionIndex;
    sectionIndex = next;
    if (!page || previous < 0) {
      numberInSection = 0;
      openPage(next, false);
      return;
    }
    if (geometry.start === "continuous" || geometry.start === "nextColumn") {
      // The section begins on the same page in a region of its own; Word
      // evens out the columns of the region it closes.
      if (region) balance(region);
      const top = region ? region.top + regionHeight(region) : 0;
      // The page it starts on counts as the section's first page.
      numberInSection = 1;
      region = newRegion(next, top);
      if (geometry.start === "nextColumn" && !pageEmpty()) advanceColumn();
      return;
    }
    numberInSection = 0;
    const firstNumber = geometry.pageNumberStart ?? lastNumber + 1;
    const wantOdd = geometry.start === "oddPage";
    if (
      (geometry.start === "oddPage" || geometry.start === "evenPage") &&
      firstNumber % 2 === (wantOdd ? 0 : 1)
    ) {
      // Skip a page so the section starts on the required side.
      openPage(next, true);
      numberInSection = 0;
    }
    openPage(next, false);
  };

  /** Place `[from, to)` of a paragraph, splitting it by lines where needed. */
  const placeSlice = (index: number, block: LayoutBlock, from: number, to: number): void => {
    let start = from;
    for (;;) {
      const height = to - start;
      const notes = start === 0 ? block.notes : undefined;
      const room = available() - notesHeight(notes);
      const whole = start === 0 && to === block.height;
      const fragment = (end: number): Fragment =>
        whole && end === to
          ? { kind: "block", block: index, height: block.height }
          : { kind: "slice", block: index, from: start, to: end };
      if (height <= room + EPSILON) {
        place(fragment(to), notes);
        return;
      }
      if (!block.keepLines || columnEmpty()) {
        const lines = (block.lines?.() ?? []).filter(
          (l) => l > start + EPSILON && l <= to + EPSILON,
        );
        let fit = 0;
        while (fit < lines.length && (lines[fit] ?? 0) - start <= room + EPSILON) fit++;
        if (block.widowControl && lines.length >= MIN_LINES_AT_BREAK * 2) {
          // No single line left behind at the bottom or carried to the top.
          if (lines.length - fit === 1) fit--;
          if (fit < MIN_LINES_AT_BREAK) fit = 0;
        } else if (block.widowControl && fit < lines.length && fit === 1) {
          fit = 0;
        }
        if (fit === 0 && columnEmpty()) fit = Math.max(1, Math.min(lines.length - 1, fit));
        const cut = lines[fit - 1];
        if (fit > 0 && fit < lines.length && cut !== undefined) {
          place(fragment(cut), notes);
          advanceColumn();
          start = cut;
          continue;
        }
      }
      if (!columnEmpty()) {
        advanceColumn();
        continue;
      }
      // Taller than an empty column and unsplittable: let it overflow, as Word does.
      place(fragment(to), notes);
      return;
    }
  };

  const placeTable = (index: number, block: LayoutBlock, rows: readonly LayoutRow[]): void => {
    if (block.height <= available() - notesHeight(block.notes) + EPSILON) {
      place({ kind: "block", block: index, height: block.height }, block.notes);
      return;
    }
    let headerRows = 0;
    while (rows[headerRows]?.header) headerRows++;
    // A table made only of header rows has nothing to repeat them for.
    if (headerRows === rows.length) headerRows = 0;
    const headerHeight = rows.slice(0, headerRows).reduce((n, r) => n + r.height, 0);
    let row = 0;
    while (row < rows.length) {
      const repeat = row > 0 && row >= headerRows ? headerRows : 0;
      const notes = row === 0 ? block.notes : undefined;
      const room = available() - notesHeight(notes);
      let height = repeat ? headerHeight : 0;
      let end = row;
      while (end < rows.length && height + (rows[end]?.height ?? 0) <= room + EPSILON) {
        height += rows[end]?.height ?? 0;
        end++;
      }
      // Header rows are never left alone at the bottom of a page.
      if (row === 0 && end > 0 && end <= headerRows && end < rows.length) end = 0;
      if (end === row) {
        if (!columnEmpty()) {
          advanceColumn();
          continue;
        }
        height = (repeat ? headerHeight : 0) + (rows[end]?.height ?? 0);
        end++;
      }
      place({ kind: "rows", block: index, from: row, to: end, headerRows: repeat, height }, notes);
      row = end;
      if (row < rows.length) advanceColumn();
    }
  };

  /** Height a keepNext chain starting at `index` needs on one page. */
  const keepChainHeight = (index: number, limit: number): number => {
    let height = 0;
    let i = index;
    for (; i < blocks.length && height <= limit; i++) {
      const b = blocks[i];
      if (!b || b.section !== blocks[index]?.section) return height;
      if (!b.keepNext) {
        const first = b.lines?.()[0];
        return height + (first ?? b.height);
      }
      height += b.height;
    }
    return height;
  };

  for (const [index, block] of blocks.entries()) {
    if (block.section !== sectionIndex) startSection(block.section);
    if (block.pageBreakBefore && !pageEmpty()) openPage(sectionIndex, false);
    if (block.keepNext && !columnEmpty()) {
      const full = bodyHeight() - (region?.top ?? 0);
      const chain = keepChainHeight(index, full);
      if (chain > available() + EPSILON && chain <= full) advanceColumn();
    }
    if (block.rows && block.rows.length > 0) {
      placeTable(index, block, block.rows);
      continue;
    }
    let from = 0;
    for (const forced of block.breaks ?? []) {
      if (forced.offset <= from || forced.offset > block.height) continue;
      placeSlice(index, block, from, forced.offset);
      if (forced.kind === "page") openPage(sectionIndex, false);
      else advanceColumn();
      from = forced.offset;
    }
    placeSlice(index, block, from, block.height);
  }
  if (pages.length === 0) {
    // An empty document still has one page (in its last section).
    sectionIndex = Math.max(0, sections.length - 1);
    openPage(sectionIndex, false);
  }
  alignPages(pages, sections);
  return pages;
}

function alignPages(pages: readonly LaidOutPage[], sections: readonly PaginatorSection[]): void {
  for (const page of pages) {
    const geometry = sections[page.section];
    const only = page.regions.length === 1 ? page.regions[0] : undefined;
    if (!geometry || !only || geometry.vAlign === "top") continue;
    const free = Math.max(0, geometry.bodyHeight(page.kind) - page.noteHeight - regionHeight(only));
    if (geometry.vAlign === "center") only.offset = free / 2;
    else if (geometry.vAlign === "bottom") only.offset = free;
    else {
      for (const column of only.columns) {
        const gaps = column.fragments.length - 1;
        const own = Math.max(0, geometry.bodyHeight(page.kind) - page.noteHeight - column.height);
        column.gap = gaps > 0 ? own / gaps : 0;
      }
    }
  }
}

/** Pages per section, for SECTIONPAGES. */
export function sectionPageCounts(pages: readonly LaidOutPage[]): Map<number, number> {
  const out = new Map<number, number>();
  for (const page of pages) out.set(page.section, (out.get(page.section) ?? 0) + 1);
  return out;
}
