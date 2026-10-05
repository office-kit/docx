/**
 * Print Layout on the canvas: renders the document once, measures every block
 * in an off-screen galley, lets the pure paginator (`paginate` in the editor
 * package) decide what goes on which page, then moves the rendered blocks into
 * page boxes with their headers, footers, footnotes and decorations.
 *
 * Rendered block elements are moved, never re-rendered, between layouts, so
 * the caret survives a re-pagination; a paragraph split across pages is shown
 * through clipped windows onto the same paragraph (the first window holds the
 * element itself, later ones a clone marked `data-wk-clone`).
 */

import {
  collectNoteReferences,
  createFlagResolver,
  documentSections,
  endnotesAtSectionEnd,
  type Fragment,
  type LaidOutPage,
  type LayoutBlock,
  type LayoutNote,
  type LayoutSettings,
  layoutSettings,
  type NoteOccurrence,
  numberNotes,
  type PageKind,
  paginate,
  type PaginatorSection,
  type ParagraphFlags,
  renderBlocksHtml,
  renderDocumentHtml,
  type SectionModel,
  sectionPageCounts,
  STORY_ATTR,
} from "@office-kit/docx-editor";
import { type Docx, storyBody, storyKey, type StoryRef } from "@office-kit/docx";
import { breakOffsets, forcedBreaks, lineBottoms, outerHeight } from "./measure";
import { decoratePage, fillPageFields, numberLines } from "./decorations";
import { twipsToPx } from "./units";

/** Room the footnote separator line takes above a page's first footnote (Word: ~1 line). */
const NOTE_SEPARATOR_PX = 14;
const PAGE_KINDS: readonly PageKind[] = ["default", "first", "even"];

export interface LayoutOptions {
  /** View ▸ Gridlines. */
  readonly showGridlines: boolean;
  /** The header/footer story being edited: the body is dimmed and header areas become editable. */
  readonly editingHeaderFooter: boolean;
}

interface Item {
  /** The block's master element (a `.wk-blk` wrapper around what the renderer produced). */
  readonly content: HTMLElement;
  /** Body block index; `undefined` for an endnotes area. */
  readonly block: number | undefined;
  readonly section: number;
  height: number;
  lines: number[] | undefined;
  readonly flags: ParagraphFlags | undefined;
  readonly table: HTMLTableElement | undefined;
  readonly rows: HTMLTableRowElement[];
  rowHeights: Array<{ height: number; header: boolean }>;
  /** A table's height beyond its rows (borders, spacing), measured whole in the galley. */
  tableExtra: number;
  notes: LayoutNote[];
  breaks: Array<{ offset: number; kind: "page" | "column" }>;
}

interface StoryElement {
  readonly el: HTMLElement;
  height: number;
}

/** Where things sit on a page, in px. */
export interface PageGeometry {
  readonly width: number;
  readonly height: number;
  readonly bodyTop: number;
  readonly bodyBottom: number;
  readonly left: number;
  readonly right: number;
  readonly textWidth: number;
  readonly headerTop: number;
  readonly footerBottom: number;
}

function sectionGeometry(section: SectionModel): Omit<PageGeometry, "bodyTop" | "bodyBottom"> {
  const m = section.margins;
  const width = twipsToPx(section.pageWidth);
  // The gutter is added to the left (inside) margin, as Word does without mirrorMargins.
  const left = twipsToPx(m.left + m.gutter);
  const right = twipsToPx(m.right);
  return {
    width,
    height: twipsToPx(section.pageHeight),
    left,
    right,
    textWidth: width - left - right,
    headerTop: twipsToPx(m.header),
    footerBottom: twipsToPx(m.footer),
  };
}

function el(tag: string, cls: string): HTMLElement {
  const e = document.createElement(tag);
  e.className = cls;
  return e;
}

/** Group the renderer's top-level elements by their `data-wk-block` index. */
function blockElements(html: string, count: number): HTMLElement[] {
  const tpl = document.createElement("template");
  tpl.innerHTML = html;
  const out: HTMLElement[] = Array.from({ length: count }, () => el("div", "wk-blk"));
  let current = 0;
  // Snapshot: appending a node elsewhere removes it from the live NodeList.
  for (const node of Array.from(tpl.content.childNodes)) {
    if (node instanceof HTMLElement) {
      const index = Number.parseInt(node.getAttribute("data-wk-block") ?? "", 10);
      if (Number.isInteger(index) && index >= 0 && index < count) current = index;
    }
    out[current]?.appendChild(node);
  }
  return out;
}

function storyContainer(doc: Docx, ref: StoryRef, cls: string): HTMLElement | undefined {
  const body = storyBody(doc, ref);
  if (!body) return undefined;
  const container = el("div", cls);
  container.setAttribute(STORY_ATTR, storyKey(ref));
  container.innerHTML = renderBlocksHtml(doc, body.blocks);
  return container;
}

export class PageLayout {
  pages: LaidOutPage[] = [];
  sections: SectionModel[] = [];
  private settings: LayoutSettings = { evenAndOddHeaders: false };
  private items: Item[] = [];
  private readonly headers = new Map<string, StoryElement>();
  private readonly notes = new Map<string, StoryElement>();
  private refs: NoteOccurrence[] = [];
  private noteMarks = new Map<string, string>();
  private readonly pageOf = new Map<number, number>();
  private sectionPages = new Map<number, number>();
  private doc: Docx | undefined;
  private options: LayoutOptions = { showGridlines: false, editingHeaderFooter: false };

  constructor(
    private readonly root: HTMLElement,
    private readonly galley: HTMLElement,
  ) {}

  /** settings `w:evenAndOddHeaders` of the laid-out document. */
  get evenAndOddHeaders(): boolean {
    return this.settings.evenAndOddHeaders;
  }

  /** Render, measure and paginate the whole document. */
  build(doc: Docx, options: LayoutOptions): void {
    this.doc = doc;
    this.options = options;
    this.sections = documentSections(doc);
    this.settings = layoutSettings(doc);
    this.headers.clear();
    this.notes.clear();
    this.galley.replaceChildren();
    const flags = createFlagResolver(doc);
    const blocks = doc.document.body.blocks;
    const contents = blockElements(renderDocumentHtml(doc), blocks.length);

    // One galley column per section at its first column's width, and one at
    // its text width for headers, footers and notes.
    const columnGalleys = this.sections.map((s) => {
      const g = el("div", "wk-galley-col");
      g.style.width = `${twipsToPx(s.columns[0]?.width ?? 0)}px`;
      this.galley.appendChild(g);
      return g;
    });
    const textGalleys = this.sections.map((s) => {
      const g = el("div", "wk-galley-col");
      g.style.width = `${sectionGeometry(s).textWidth}px`;
      this.galley.appendChild(g);
      return g;
    });

    this.refs = collectNoteReferences(doc);
    const footnotesByBlock = new Map<number, string[]>();
    const endnotesBySection = new Map<number, string[]>();
    const lastSection = this.sections.length - 1;
    const blockSection: number[] = [];
    for (const s of this.sections)
      for (let b = s.firstBlock; b <= s.lastBlock; b++) blockSection[b] = s.index;
    for (const ref of this.refs) {
      const section = blockSection[ref.block] ?? lastSection;
      if (ref.kind === "footnote") {
        const list = footnotesByBlock.get(ref.block) ?? [];
        list.push(ref.key);
        footnotesByBlock.set(ref.block, list);
        const note = storyContainer(doc, { kind: "footnote", id: ref.id }, "wk-note");
        if (note) {
          textGalleys[section]?.appendChild(note);
          this.notes.set(ref.key, { el: note, height: 0 });
        }
      } else {
        const target = endnotesAtSectionEnd(this.sections[section]) ? section : lastSection;
        const list = endnotesBySection.get(target) ?? [];
        list.push(ref.key);
        endnotesBySection.set(target, list);
      }
    }

    for (const section of this.sections) {
      for (const kind of ["header", "footer"] as const) {
        const parts = kind === "header" ? section.headers : section.footers;
        for (const pageKind of PAGE_KINDS) {
          const partName = parts[pageKind];
          if (!partName || this.headers.has(partName)) continue;
          const container = storyContainer(doc, { kind, partName }, "wk-hf-content");
          if (!container) continue;
          textGalleys[section.index]?.appendChild(container);
          this.headers.set(partName, { el: container, height: 0 });
        }
      }
    }

    this.items = [];
    for (const section of this.sections) {
      for (let b = section.firstBlock; b <= section.lastBlock; b++) {
        const content = contents[b];
        const node = blocks[b];
        if (!content || !node) continue;
        columnGalleys[section.index]?.appendChild(content);
        const table =
          node.kind === "table" ? (content.querySelector("table") ?? undefined) : undefined;
        const rows = table
          ? [...table.querySelectorAll<HTMLTableRowElement>(":scope > tbody > tr")]
          : [];
        const headerRows =
          node.kind === "table"
            ? node.rows.map(
                (r) =>
                  !!r.trPr?.children.some(
                    (c) => c.kind === "element" && c.name.local === "tblHeader",
                  ),
              )
            : [];
        this.items.push({
          content,
          block: b,
          section: section.index,
          height: 0,
          lines: undefined,
          flags: node.kind === "paragraph" ? flags(node) : undefined,
          table,
          rows,
          rowHeights: rows.map((_, i) => ({ height: 0, header: headerRows[i] ?? false })),
          tableExtra: 0,
          notes: (footnotesByBlock.get(b) ?? []).map((key) => ({ key, height: 0 })),
          breaks: [],
        });
      }
      const endnotes = endnotesBySection.get(section.index);
      if (endnotes?.length) {
        const area = el("div", "wk-blk wk-endnotes");
        area.appendChild(el("div", "wk-note-sep"));
        for (const key of endnotes) {
          const id = Number(key.slice(key.indexOf(":") + 1));
          const note = storyContainer(doc, { kind: "endnote", id }, "wk-note");
          if (note) area.appendChild(note);
        }
        columnGalleys[section.index]?.appendChild(area);
        this.items.push({
          content: area,
          block: undefined,
          section: section.index,
          height: 0,
          lines: undefined,
          flags: undefined,
          table: undefined,
          rows: [],
          rowHeights: [],
          tableExtra: 0,
          notes: [],
          breaks: [],
        });
      }
    }

    // One read pass after all the writes above.
    for (const story of this.headers.values()) story.height = outerHeight(story.el);
    for (const note of this.notes.values()) note.height = outerHeight(note.el);
    for (const item of this.items) {
      this.measureItem(item);
      if (item.table) {
        item.tableExtra = item.content.offsetHeight - item.height;
        item.height += item.tableExtra;
      }
      item.notes = item.notes.map((n) => ({
        key: n.key,
        height: this.notes.get(n.key)?.height ?? 0,
      }));
    }
    this.distribute(options);
  }

  private measureItem(item: Item): void {
    item.lines = undefined;
    item.rowHeights = item.rows.map((row, i) => ({
      height: row.offsetHeight,
      header: item.rowHeights[i]?.header ?? false,
    }));
    // A table may be split over several page boxes; its rows add up wherever they are.
    item.height = item.table
      ? item.rowHeights.reduce((n, r) => n + r.height, 0) + item.tableExtra
      : item.content.offsetHeight;
    const node =
      item.block === undefined || !this.doc ? undefined : this.doc.document.body.blocks[item.block];
    item.breaks = node?.kind === "paragraph" ? breakOffsets(item.content, forcedBreaks(node)) : [];
  }

  /**
   * Re-measure one block after typing changed it. Returns whether its height
   * changed, i.e. whether the pages must be laid out again.
   */
  remeasure(block: number): boolean {
    const item = this.items.find((i) => i.block === block);
    if (!item) return false;
    // A clipped (split) block is measured through its master element, whose
    // full height the clip does not change.
    const before = item.height;
    this.measureItem(item);
    return Math.abs(item.height - before) > 0.5 || item.breaks.length > 0;
  }

  /** The 0-based page holding a body block (its first fragment). */
  pageOfBlock(block: number): number {
    return this.pageOf.get(block) ?? 0;
  }

  /** Lay the measured items out again and rebuild the page boxes. */
  distribute(options: LayoutOptions): void {
    this.options = options;
    const geometry = this.sections.map((section) => this.paginatorSection(section));
    const blocks: LayoutBlock[] = this.items.map((item) => {
      const flags = item.flags;
      return {
        height: item.height,
        section: item.section,
        ...(flags?.pageBreakBefore ? { pageBreakBefore: true } : {}),
        ...(flags?.keepNext ? { keepNext: true } : {}),
        ...(flags?.keepLines ? { keepLines: true } : {}),
        ...(flags?.widowControl ? { widowControl: true } : {}),
        ...(item.table
          ? { rows: item.rowHeights }
          : { lines: () => (item.lines ??= lineBottoms(item.content)) }),
        ...(item.breaks.length ? { breaks: item.breaks } : {}),
        ...(item.notes.length ? { notes: item.notes } : {}),
      };
    });
    this.pages = paginate(blocks, geometry, {
      evenAndOddHeaders: this.settings.evenAndOddHeaders,
      noteSeparatorHeight: NOTE_SEPARATOR_PX,
    });

    this.pageOf.clear();
    this.pages.forEach((page, p) => {
      for (const region of page.regions)
        for (const column of region.columns)
          for (const f of column.fragments) {
            const block = this.items[f.block]?.block;
            if (block !== undefined && !this.pageOf.has(block)) this.pageOf.set(block, p);
          }
    });
    const blockSection = (block: number): number =>
      this.sections.find((s) => block >= s.firstBlock && block <= s.lastBlock)?.index ?? 0;
    this.noteMarks = numberNotes(this.refs, this.sections, blockSection, (b) =>
      this.pageOfBlock(b),
    );

    // Put every table's rows back in its master before re-splitting.
    for (const item of this.items) {
      const body = item.table?.tBodies[0];
      if (body) body.replaceChildren(...item.rows);
      item.content.style.marginBottom = "";
      item.content.style.marginTop = "";
    }
    const container = el("div", "wk-pages");
    this.sectionPages = sectionPageCounts(this.pages);
    const placed = new Set<Item>();
    container.append(...this.pages.map((page, i) => this.buildPage(page, i, placed)));
    this.root.replaceChildren(container);
    this.galley.replaceChildren();
    const suppressed = new Set(
      this.items.filter((i) => i.flags?.suppressLineNumbers).map((i) => i.block),
    );
    numberLines(this.root, this.pages, this.sections, (block) => suppressed.has(block));
  }

  private headerHeight(section: SectionModel, kind: "header" | "footer", page: PageKind): number {
    const part = (kind === "header" ? section.headers : section.footers)[page];
    return part ? (this.headers.get(part)?.height ?? 0) : 0;
  }

  /** Body box of a page: a header or footer taller than its margin pushes the body. */
  geometry(sectionIndex: number, kind: PageKind): PageGeometry {
    const section = this.sections[sectionIndex] ?? this.sections[0];
    if (!section) throw new Error("The document has no sections.");
    const base = sectionGeometry(section);
    const m = section.margins;
    const top = twipsToPx(m.top);
    const bottom = twipsToPx(m.bottom);
    const header = base.headerTop + this.headerHeight(section, "header", kind);
    const footer = base.footerBottom + this.headerHeight(section, "footer", kind);
    return {
      ...base,
      bodyTop: m.fixedTop ? top : Math.max(top, header),
      bodyBottom: m.fixedBottom ? bottom : Math.max(bottom, footer),
    };
  }

  private paginatorSection(section: SectionModel): PaginatorSection {
    return {
      start: section.start,
      columns: section.columns.length,
      titlePage: section.titlePage,
      pageNumberStart: section.pageNumbers.start,
      vAlign: section.vAlign,
      bodyHeight: (kind) => {
        const g = this.geometry(section.index, kind);
        return g.height - g.bodyTop - g.bodyBottom;
      },
    };
  }

  private fragmentElement(f: Fragment, firstUse: Set<Item>): HTMLElement | undefined {
    const item = this.items[f.block];
    if (!item) return undefined;
    if (f.kind === "block") {
      firstUse.add(item);
      return item.content;
    }
    if (f.kind === "slice") {
      const window = el("div", "wk-slice");
      window.style.height = `${f.to - f.from}px`;
      const master = !firstUse.has(item);
      const inner = master ? item.content : this.cloneOf(item.content);
      firstUse.add(item);
      inner.style.marginTop = `${-f.from}px`;
      window.appendChild(inner);
      return window;
    }
    // Rows: the first fragment keeps the master table; later ones get a shell.
    const master = !firstUse.has(item);
    firstUse.add(item);
    const holder = master ? item.content : this.cloneOf(item.content, true);
    const table = master ? item.table : (holder.querySelector("table") ?? undefined);
    const body = table?.tBodies[0];
    if (!body) return holder;
    const rows: HTMLElement[] = [];
    for (let r = 0; r < f.headerRows; r++) {
      const header = item.rows[r];
      if (!header) continue;
      const copy = this.cloneOf(header);
      copy.setAttribute("contenteditable", "false");
      rows.push(copy);
    }
    rows.push(...item.rows.slice(f.from, f.to));
    body.replaceChildren(...rows);
    return holder;
  }

  /** A copy of rendered content that must not be mistaken for the original. */
  private cloneOf(node: HTMLElement, shellOnly = false): HTMLElement {
    const copy = node.cloneNode(true);
    if (!(copy instanceof HTMLElement))
      throw new Error("Cloning an element returned a non-element.");
    if (shellOnly) copy.querySelector("tbody")?.replaceChildren();
    copy.setAttribute("data-wk-clone", "");
    return copy;
  }

  /** `firstUse`: items already placed in this distribution (their later fragments clone). */
  private buildPage(page: LaidOutPage, index: number, firstUse: Set<Item>): HTMLElement {
    const section = this.sections[page.section];
    if (!section) throw new Error(`Page ${index} has no section.`);
    const g = this.geometry(page.section, page.kind);
    const box = el("div", "wk-pagebox");
    box.dataset.page = String(index);
    box.style.width = `${g.width}px`;
    box.style.height = `${g.height}px`;
    for (const [name, value] of [
      ["--page-w", g.width],
      ["--page-h", g.height],
      ["--m-top", g.bodyTop],
      ["--m-bottom", g.bodyBottom],
      ["--m-left", g.left],
      ["--m-right", g.right],
    ] as const) {
      box.style.setProperty(name, `${value}px`);
    }
    if (this.settings.pageColor) box.style.backgroundColor = `#${this.settings.pageColor}`;

    // Header, body and footer share the page's one grid cell and are placed
    // by margins rather than absolute positioning: the page box stays the
    // nearest positioned ancestor, which floating objects (anchored shapes,
    // watermarks) position against through the --page-* / --m-* variables.
    const hfEditable = this.options.editingHeaderFooter ? "true" : "false";
    const area = (kind: "header" | "footer"): HTMLElement => {
      const zone = el("div", `wk-hf wk-${kind}`);
      zone.dataset.kind = kind;
      zone.dataset.pageKind = page.kind;
      zone.dataset.section = String(page.section);
      zone.setAttribute("contenteditable", hfEditable);
      zone.style.marginLeft = `${g.left}px`;
      zone.style.width = `${g.textWidth}px`;
      if (kind === "header") {
        zone.style.marginTop = `${g.headerTop}px`;
        zone.style.minHeight = `${Math.max(0, g.bodyTop - g.headerTop)}px`;
      } else {
        zone.style.marginBottom = `${g.footerBottom}px`;
        zone.style.minHeight = `${Math.max(0, g.bodyBottom - g.footerBottom)}px`;
      }
      const part = (kind === "header" ? section.headers : section.footers)[page.kind];
      const master = part ? this.headers.get(part) : undefined;
      if (master) zone.appendChild(this.cloneOf(master.el));
      return zone;
    };

    const body = el("div", "wk-body");
    body.style.marginTop = `${g.bodyTop}px`;
    body.style.marginLeft = `${g.left}px`;
    body.style.width = `${g.textWidth}px`;
    body.style.height = `${g.height - g.bodyTop - g.bodyBottom}px`;
    if (this.options.editingHeaderFooter) body.setAttribute("contenteditable", "false");

    const flow = el("div", "wk-regions");
    let y = 0;
    for (const region of page.regions) {
      const rs = this.sections[region.section] ?? section;
      const r = el("div", "wk-region");
      r.style.marginTop = `${region.top + region.offset - y}px`;
      const height = Math.max(0, ...region.columns.map((c) => c.height));
      r.style.height = `${height}px`;
      y = region.top + region.offset + height;
      region.columns.forEach((column, c) => {
        const spec = rs.columns[c];
        const col = el("div", "wk-col");
        col.style.width = `${twipsToPx(spec?.width ?? 0)}px`;
        for (const [i, f] of column.fragments.entries()) {
          const node = this.fragmentElement(f, firstUse);
          if (!node) continue;
          if (column.gap > 0 && i < column.fragments.length - 1)
            node.style.marginBottom = `${column.gap}px`;
          col.appendChild(node);
        }
        r.appendChild(col);
        if (c < region.columns.length - 1) {
          const gap = el("div", rs.separator ? "wk-colgap sep" : "wk-colgap");
          gap.setAttribute("contenteditable", "false");
          gap.style.width = `${twipsToPx(spec?.space ?? 0)}px`;
          gap.style.height = `${height}px`;
          r.appendChild(gap);
        }
      });
      flow.appendChild(r);
    }
    body.appendChild(flow);

    if (page.notes.length) {
      const notes = el("div", "wk-notes");
      notes.style.height = `${page.noteHeight}px`;
      const separator = el("div", "wk-note-sep");
      separator.setAttribute("contenteditable", "false");
      notes.appendChild(separator);
      for (const key of page.notes) {
        const note = this.notes.get(key);
        if (note) notes.appendChild(note.el);
      }
      body.appendChild(notes);
    }

    box.append(area("header"), body, area("footer"));
    decoratePage(box, {
      page,
      section,
      geometry: g,
      showGridlines: this.options.showGridlines,
      doc: this.doc,
    });
    fillPageFields(box, {
      page,
      section,
      totalPages: this.pages.length,
      sectionPages: this.sectionPages.get(page.section) ?? 0,
      noteMarks: this.noteMarks,
    });
    return box;
  }

  /** One continuous flow without page boxes (Web Layout, Draft, Outline, Read). */
  buildFlow(doc: Docx): void {
    this.doc = doc;
    this.pages = [];
    const flow = el("div", "wk-flow");
    const blocks = doc.document.body.blocks;
    flow.append(...blockElements(renderDocumentHtml(doc), blocks.length));
    this.items = [];
    this.root.replaceChildren(flow);
  }
}
