/**
 * Reference commands: bookmarks, hyperlinks, fields, and Word's References
 * tab — tables of contents and figures, footnotes and endnotes, citations and
 * bibliography, captions, index and table of authorities. Commands that place
 * something "at the caret" use the caret paragraph and its character offset.
 */

import {
  addBookmark,
  addCaptionLabel,
  addHyperlink,
  addInternalHyperlink,
  appendField,
  type BibliographySource,
  type CaptionOptions,
  type CitationMarkOptions,
  type CitationOptions,
  type CitationStyle,
  convertNotes,
  type Docx,
  type IndexEntryOptions,
  type IndexOptions,
  insertBibliography,
  insertCaption,
  insertCitation,
  insertIndex,
  insertNote,
  type InsertNoteOptions,
  insertTableOfAuthorities,
  insertTableOfContents,
  markAllIndexEntries,
  markAuthorityCitation,
  markIndexEntry,
  type NoteConversion,
  type NoteKind,
  type NoteProperties,
  type PageNumberProvider,
  removeTableOfContents,
  setBibliographySources,
  setBibliographyStyle,
  setNoteProperties,
  setNoteText,
  setTocLevel,
  type TableFieldType,
  type TableOfAuthoritiesOptions,
  type TableOfContentsOptions,
  type TocLevel,
  updateTables,
  type WmlParagraph,
  type WmlRun,
} from "@office-kit/docx";
import { absoluteOffset } from "../char-offset.js";
import { paragraphAt } from "../doc-access.js";
import type { EditorModel } from "../model.js";
import { comparePositions, type DocPosition, orderSelection } from "../selection.js";
import type { Command } from "./types.js";
import { caretBlockIndex, moveLastBlockAfter } from "./insert-util.js";

/** The caret paragraph, or undefined when the caret is not in a paragraph. */
function caretParagraph(model: EditorModel) {
  const pos = model.selection?.focus;
  return pos ? paragraphAt(model.doc, pos) : undefined;
}

/**
 * The caret paragraph and the character offset to insert at: the end of a
 * selected range (Word marks index entries and citations after the
 * selection), else the caret.
 */
function insertionPoint(
  model: EditorModel,
): { paragraph: WmlParagraph; offset: number } | undefined {
  const sel = model.selection;
  if (!sel) return undefined;
  const at = orderSelection(sel).end;
  const paragraph = paragraphAt(model.doc, at);
  return paragraph ? { paragraph, offset: absoluteOffset(paragraph, at) } : undefined;
}

// One character per piece, in `runTextLength` units, so offsets line up.
const PIECE_TEXT: Readonly<Record<string, string>> = {
  tab: "\t",
  break: "\n",
  noBreakHyphen: "-",
  softHyphen: "­",
};

/**
 * The selected text when the selection lies within one paragraph (what Mark
 * Entry and Mark Citation prefill), else "".
 */
export function selectedText(model: EditorModel): string {
  const sel = model.selection;
  if (!sel) return "";
  const { start, end, collapsed } = orderSelection(sel);
  const paragraph = paragraphAt(model.doc, start);
  if (collapsed || !paragraph || paragraph !== paragraphAt(model.doc, end)) return "";
  const text = paragraph.children
    .filter((c): c is WmlRun => c.kind === "run")
    .flatMap((run) =>
      run.pieces.map((p) => (p.kind === "text" ? p.value : (PIECE_TEXT[p.kind] ?? ""))),
    )
    .join("");
  return text.slice(absoluteOffset(paragraph, start), absoluteOffset(paragraph, end));
}

function requirePoint(model: EditorModel): { paragraph: WmlParagraph; offset: number } {
  const point = insertionPoint(model);
  if (!point) throw new Error("Place the insertion point in text first.");
  return point;
}

/**
 * Where a generated table (TOC, index …) goes: before the caret's top-level
 * block when the caret is at its very start, otherwise after it.
 */
function tableInsertIndex(model: EditorModel): number {
  const focus = model.selection?.focus;
  const block = caretBlockIndex(model.doc, focus?.block);
  if (!focus) return model.doc.document.body.blocks.length;
  const paragraph = paragraphAt(model.doc, focus);
  const atStart = !focus.cell && paragraph !== undefined && absoluteOffset(paragraph, focus) === 0;
  return atStart ? block : block + 1;
}

const hasCaret = (model: EditorModel): boolean => !!caretParagraph(model);

export const addBookmarkCommand: Command<{ name: string }> = {
  id: "references.bookmark",
  group: "references",
  label: "Bookmark",
  run(model, { name }) {
    const para = caretParagraph(model);
    if (para) addBookmark(model.doc, name, para);
  },
  isEnabled: hasCaret,
};

// Schemes a hyperlink may target. Anything else (notably `javascript:` /
// `data:`) is rejected: the target ends up as a clickable link in Word and in
// any HTML rendering of the document.
const ALLOWED_LINK_PROTOCOLS: ReadonlySet<string> = new Set(["http:", "https:", "mailto:"]);

function assertSafeUrl(url: string): void {
  let protocol: string;
  try {
    protocol = new URL(url).protocol;
  } catch {
    throw new Error(`Invalid hyperlink URL: ${JSON.stringify(url)}.`);
  }
  if (!ALLOWED_LINK_PROTOCOLS.has(protocol)) {
    throw new Error(`Unsupported hyperlink scheme ${JSON.stringify(protocol)}.`);
  }
}

export const insertHyperlinkCommand: Command<{ url: string; text: string; tooltip?: string }> = {
  id: "references.hyperlink",
  group: "references",
  label: "Hyperlink",
  run(model, { url, text, tooltip }) {
    assertSafeUrl(url);
    const at = caretBlockIndex(model.doc, model.selection?.focus.block);
    addHyperlink(model.doc, url, text, tooltip !== undefined ? { tooltip } : {});
    moveLastBlockAfter(model.doc, at);
  },
};

export const insertInternalLinkCommand: Command<{
  bookmark: string;
  text: string;
  tooltip?: string;
}> = {
  id: "references.internalLink",
  group: "references",
  label: "Link to bookmark",
  run(model, { bookmark, text, tooltip }) {
    const at = caretBlockIndex(model.doc, model.selection?.focus.block);
    addInternalHyperlink(model.doc, bookmark, text, tooltip !== undefined ? { tooltip } : {});
    moveLastBlockAfter(model.doc, at);
  },
};

export const insertFieldCommand: Command<{ instruction: string }> = {
  id: "references.field",
  group: "references",
  label: "Field",
  run(model, { instruction }) {
    const para = caretParagraph(model);
    if (para) appendField(model.doc, para, instruction);
  },
  isEnabled: hasCaret,
};

// --- Table of contents ---------------------------------------------------------

/** Table of Contents gallery / Custom Table of Contents. */
export const insertTocCommand: Command<TableOfContentsOptions> = {
  id: "references.toc",
  group: "references",
  label: "Table of Contents",
  run(model, options) {
    insertTableOfContents(model.doc, tableInsertIndex(model), options);
  },
};

export const removeTocCommand: Command<void, number> = {
  id: "references.removeToc",
  group: "references",
  label: "Remove Table of Contents",
  run: (model) => removeTableOfContents(model.doc),
};

/** Add Text: make the caret paragraph a TOC entry of a level, or exclude it. */
export const addTextCommand: Command<{ level: TocLevel }> = {
  id: "references.addText",
  group: "references",
  label: "Add Text",
  run(model, { level }) {
    const paragraph = caretParagraph(model);
    if (paragraph) setTocLevel(model.doc, paragraph, level);
  },
  isEnabled: hasCaret,
};

/** Update Table / Update Index / Update Table of Authorities. */
export const updateTablesCommand: Command<
  { types: readonly TableFieldType[]; pageOf?: PageNumberProvider },
  number
> = {
  id: "references.updateTables",
  group: "references",
  label: "Update Table",
  run: (model, { types, pageOf }) =>
    updateTables(model.doc, { types, ...(pageOf ? { pageOf } : {}) }),
};

// --- Footnotes ------------------------------------------------------------------

function noteCommand(
  kind: NoteKind,
  id: string,
  label: string,
): Command<InsertNoteOptions, number> {
  return {
    id,
    group: "references",
    label,
    run(model, options) {
      const { paragraph, offset } = requirePoint(model);
      return insertNote(model.doc, paragraph, offset, kind, options);
    },
    isEnabled: hasCaret,
  };
}

export const addFootnoteCommand = noteCommand("footnote", "references.footnote", "Insert Footnote");
export const addEndnoteCommand = noteCommand("endnote", "references.endnote", "Insert Endnote");

/** Footnote and Endnote dialog ▸ Apply. */
export const noteOptionsCommand: Command<{
  kind: NoteKind;
  properties: NoteProperties;
  scope: "document" | { section: number };
}> = {
  id: "references.noteOptions",
  group: "references",
  label: "Footnote and Endnote",
  run(model, { kind, properties, scope }) {
    setNoteProperties(model.doc, kind, properties, scope);
  },
};

/** Edit a note's text (Show Notes). */
export const noteTextCommand: Command<{ kind: NoteKind; id: number; text: string }> = {
  id: "references.noteText",
  group: "references",
  label: "Note Text",
  run(model, { kind, id, text }) {
    setNoteText(model.doc, kind, id, text);
  },
};

export const convertNotesCommand: Command<{ conversion: NoteConversion }, number> = {
  id: "references.convertNotes",
  group: "references",
  label: "Convert Notes",
  run: (model, { conversion }) => convertNotes(model.doc, conversion),
};

/** The run index of each note reference of `kind`, per paragraph position, in document order. */
function noteReferencePositions(doc: Docx, kind: NoteKind): DocPosition[] {
  const local = kind === "footnote" ? "footnoteReference" : "endnoteReference";
  const out: DocPosition[] = [];
  const visit = (paragraph: WmlParagraph, base: DocPosition): void => {
    const runs = paragraph.children.filter((c): c is WmlRun => c.kind === "run");
    runs.forEach((run, inline) => {
      if (run.pieces.some((p) => p.kind === "raw" && p.node.name.local === local)) {
        out.push({ ...base, inline, offset: 0 });
      }
    });
  };
  doc.document.body.blocks.forEach((block, i) => {
    if (block.kind === "paragraph") visit(block, { block: i });
    else if (block.kind === "table") {
      block.rows.forEach((row, r) =>
        row.cells.forEach((cell, c) =>
          cell.paragraphs.forEach((p, para) =>
            visit(p, { block: i, cell: { row: r, col: c }, para }),
          ),
        ),
      );
    }
  });
  return out;
}

/**
 * Next Footnote menu: the position of the next (or previous) footnote or
 * endnote reference from the caret, or undefined when there is none. Moving
 * the caret is not an edit, so this is a query the UI applies with
 * `model.setSelection`.
 */
export function findNoteReference(
  model: EditorModel,
  kind: NoteKind,
  direction: "next" | "previous",
): DocPosition | undefined {
  const positions = noteReferencePositions(model.doc, kind);
  const caret = model.selection?.focus ?? { block: 0 };
  return direction === "next"
    ? positions.find((p) => comparePositions(p, caret) > 0)
    : positions.findLast((p) => comparePositions(p, caret) < 0);
}

/** Whether the document has any reference of `kind` (Show Notes / Next Footnote enablement). */
export function hasNoteReferences(model: EditorModel, kind: NoteKind): boolean {
  return noteReferencePositions(model.doc, kind).length > 0;
}

// --- Citations & bibliography -------------------------------------------------------------

/** Manage Sources ▸ Current List (also how Insert Citation ▸ Add New Source adds one). */
export const setSourcesCommand: Command<{ sources: readonly BibliographySource[] }> = {
  id: "references.sources",
  group: "references",
  label: "Manage Sources",
  run(model, { sources }) {
    setBibliographySources(model.doc, sources);
  },
};

export const citationStyleCommand: Command<{ style: CitationStyle }> = {
  id: "references.citationStyle",
  group: "references",
  label: "Bibliography Style",
  run(model, { style }) {
    setBibliographyStyle(model.doc, style);
  },
};

export const insertCitationCommand: Command<{ tag: string; options?: CitationOptions }> = {
  id: "references.citation",
  group: "references",
  label: "Insert Citation",
  run(model, { tag, options }) {
    const { paragraph, offset } = requirePoint(model);
    insertCitation(model.doc, paragraph, offset, tag, options ?? {});
  },
  isEnabled: hasCaret,
};

export const insertBibliographyCommand: Command<{ title?: string }> = {
  id: "references.bibliography",
  group: "references",
  label: "Bibliography",
  run(model, { title }) {
    insertBibliography(model.doc, tableInsertIndex(model), title === undefined ? {} : { title });
  },
};

// --- Captions --------------------------------------------------------------------------

export const insertCaptionCommand: Command<{
  position: "above" | "below";
  options: CaptionOptions;
}> = {
  id: "references.caption",
  group: "references",
  label: "Insert Caption",
  run(model, { position, options }) {
    const block = caretBlockIndex(model.doc, model.selection?.focus.block);
    insertCaption(model.doc, block, position, options);
  },
};

export const addCaptionLabelCommand: Command<{ name: string }> = {
  id: "references.captionLabel",
  group: "references",
  label: "New Label",
  run(model, { name }) {
    addCaptionLabel(model.doc, name);
  },
};

// --- Index ------------------------------------------------------------------------------

/** Mark Entry ▸ Mark (or Mark All, which marks every occurrence of `markAll`). */
export const markIndexEntryCommand: Command<
  { entry: IndexEntryOptions; markAll?: string },
  number
> = {
  id: "references.markEntry",
  group: "references",
  label: "Mark Entry",
  run(model, { entry, markAll }) {
    if (markAll) return markAllIndexEntries(model.doc, markAll, entry);
    const { paragraph, offset } = requirePoint(model);
    markIndexEntry(model.doc, paragraph, offset, entry);
    return 1;
  },
  isEnabled: hasCaret,
};

export const insertIndexCommand: Command<IndexOptions> = {
  id: "references.index",
  group: "references",
  label: "Insert Index",
  run(model, options) {
    insertIndex(model.doc, tableInsertIndex(model), options);
  },
};

// --- Table of authorities ----------------------------------------------------------------

export const markCitationCommand: Command<CitationMarkOptions> = {
  id: "references.markCitation",
  group: "references",
  label: "Mark Citation",
  run(model, options) {
    const { paragraph, offset } = requirePoint(model);
    markAuthorityCitation(model.doc, paragraph, offset, options);
  },
  isEnabled: hasCaret,
};

export const insertToaCommand: Command<TableOfAuthoritiesOptions> = {
  id: "references.toa",
  group: "references",
  label: "Insert Table of Authorities",
  run(model, options) {
    insertTableOfAuthorities(model.doc, tableInsertIndex(model), options);
  },
};

export const referencesCommands = [
  addBookmarkCommand,
  insertHyperlinkCommand,
  insertInternalLinkCommand,
  insertFieldCommand,
  insertTocCommand,
  removeTocCommand,
  addTextCommand,
  updateTablesCommand,
  addFootnoteCommand,
  addEndnoteCommand,
  noteOptionsCommand,
  noteTextCommand,
  convertNotesCommand,
  setSourcesCommand,
  citationStyleCommand,
  insertCitationCommand,
  insertBibliographyCommand,
  insertCaptionCommand,
  addCaptionLabelCommand,
  markIndexEntryCommand,
  insertIndexCommand,
  markCitationCommand,
  insertToaCommand,
];
