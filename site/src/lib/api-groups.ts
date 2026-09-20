// Single source of truth for the public API listing.
//
// Consumed by:
//   - `/api` — renders the groups as sections on the docs site.
//   - `/llms-full.txt` — flattens the same groups into the API section
//     of the LLM-readable concatenation.
//   - `scripts/check-api-page.mjs` — diffs the entries here against the
//     live `@office-kit/docx` exports (plus a `previewToDOM` whitelist for
//     `@office-kit/docx-preview`). CI fails if a new export isn't added here.

export type ApiEntry = { name: string; sig?: string };
export type ApiGroup = { title: string; description: string; entries: ApiEntry[] };

export const apiGroups: ApiGroup[] = [
  {
    title: "Lifecycle",
    description: "Create, open, clone, and serialise a document.",
    entries: [
      { name: "createDocx", sig: "({ paragraphs? }?) => Docx" },
      { name: "openDocx", sig: "(bytes: Uint8Array) => Docx" },
      { name: "fromBlob", sig: "(blob: Blob) => Promise<Docx>" },
      { name: "toUint8Array", sig: "(doc: Docx) => Uint8Array" },
      { name: "toBlob", sig: "(doc: Docx) => Blob" },
      { name: "clone", sig: "(doc: Docx) => Docx" },
    ],
  },
  {
    title: "Paragraphs & blocks",
    description: "Append, insert, and remove body paragraphs, headings, and breaks.",
    entries: [
      { name: "appendParagraph" },
      { name: "insertParagraphAt" },
      { name: "removeParagraph" },
      { name: "appendHeading" },
      { name: "appendPageBreak" },
      { name: "appendLineBreak" },
      { name: "appendSectionBreak" },
      { name: "clearBody" },
      { name: "paragraphs" },
    ],
  },
  {
    title: "Inline & text",
    description:
      "Find and replace, runs and their formatting, paragraph alignment, indents, spacing, borders, and shading.",
    entries: [
      { name: "replaceText" },
      { name: "replaceTextEverywhere" },
      { name: "findText" },
      { name: "findTextEverywhere" },
      { name: "appendTextRun" },
      { name: "setParagraphText" },
      { name: "paragraphText" },
      { name: "setRunFormat" },
      { name: "clearRunFormat" },
      { name: "getRunFormat" },
      { name: "setParagraphAlignment" },
      { name: "getParagraphAlignment" },
      { name: "setParagraphIndent" },
      { name: "setParagraphSpacing" },
      { name: "setParagraphBorders" },
      { name: "setParagraphShading" },
      { name: "getParagraphStyle" },
      { name: "getParagraphNumbering" },
      { name: "mergeAdjacentRuns" },
      { name: "mergeAdjacentRunsInBody" },
    ],
  },
  {
    title: "Styles & numbering",
    description:
      "The style table, bullet and numbered lists, and lifting styles from a designed template.",
    entries: [
      { name: "addStyle" },
      { name: "removeStyle" },
      { name: "listStyles" },
      { name: "ensureHeadingStyles" },
      { name: "findStyleIdByName" },
      { name: "setParagraphStyle" },
      { name: "addBulletList" },
      { name: "addNumberedList" },
      { name: "applyListToParagraph" },
      { name: "mergeStylesFromTemplate" },
    ],
  },
  {
    title: "Tables",
    description: "Build tables and edit rows, cell text, borders, shading, and alignment.",
    entries: [
      { name: "addTable" },
      { name: "tables" },
      { name: "removeTable" },
      { name: "removeAllTables" },
      { name: "unwrapTable" },
      { name: "appendTableRow" },
      { name: "removeTableRow" },
      { name: "setTableRowAsHeader" },
      { name: "setTableRowHeight" },
      { name: "setTableBorders" },
      { name: "setTableCellText" },
      { name: "getTableCellText" },
      { name: "setTableCellShading" },
      { name: "setTableCellVerticalAlign" },
    ],
  },
  {
    title: "Images",
    description:
      "Add inline images, list the ones in a file, and replace them by part name or alt text.",
    entries: [
      { name: "addImage" },
      { name: "addImageRun" },
      { name: "insertImageInto" },
      { name: "images" },
      { name: "imageReferences" },
      { name: "replaceImage" },
      { name: "replaceImageByAltText" },
      { name: "removeAllImages" },
    ],
  },
  {
    title: "Headers, footers, sections",
    description:
      "Headers and footers for default, first, and even pages, plus page size, margins, and orientation.",
    entries: [
      { name: "addHeader" },
      { name: "addFooter" },
      { name: "addPageNumberFooter" },
      { name: "setPageSize" },
      { name: "setPageMargins" },
      { name: "setPageOrientation" },
      { name: "headers" },
      { name: "footers" },
      { name: "removeAllHeaders" },
      { name: "removeAllFooters" },
    ],
  },
  {
    title: "Comments, notes, hyperlinks, bookmarks",
    description: "Review and navigation content: add it, list it, rewrite links, or strip it out.",
    entries: [
      { name: "addComment" },
      { name: "addFootnote" },
      { name: "addEndnote" },
      { name: "removeAllComments" },
      { name: "removeAllFootnotes" },
      { name: "removeAllEndnotes" },
      { name: "addHyperlink" },
      { name: "addInternalHyperlink" },
      { name: "externalHyperlinks" },
      { name: "setHyperlinkUrl" },
      { name: "removeAllHyperlinks" },
      { name: "addBookmark" },
      { name: "removeBookmark" },
      { name: "removeAllBookmarks" },
      { name: "bookmarks" },
    ],
  },
  {
    title: "Fields & tracked changes",
    description:
      "Complex fields such as a table of contents or a merge field, and bulk accept or reject of tracked insertions and deletions.",
    entries: [
      { name: "appendField" },
      { name: "addTableOfContents" },
      { name: "appendMergeField" },
      { name: "acceptAllRevisions" },
      { name: "rejectAllRevisions" },
    ],
  },
  {
    title: "Document properties",
    description: "Core and app properties, with shortcuts for the title and the author.",
    entries: [
      { name: "coreProperties" },
      { name: "setCoreProperties" },
      { name: "appProperties" },
      { name: "setAppProperties" },
      { name: "title" },
      { name: "author" },
      { name: "setTitle" },
      { name: "setAuthor" },
    ],
  },
  {
    title: "Diagnostics",
    description:
      "Validate the package and read a document back: outline, fields, statistics, plain text.",
    entries: [
      { name: "validate" },
      { name: "validatePackage" },
      { name: "statistics" },
      { name: "outline" },
      { name: "fields" },
      { name: "text" },
    ],
  },
  {
    title: "Low-level part access",
    description: "The parsed side parts, for when the functions above do not reach far enough.",
    entries: [
      { name: "stylesPart" },
      { name: "numberingPart" },
      { name: "commentsPart" },
      { name: "footnotesPart" },
      { name: "endnotesPart" },
    ],
  },
  {
    title: "Page-size & margin constants",
    description: "Ready-made values for setPageSize and setPageMargins, plus the library version.",
    entries: [
      { name: "PAGE_SIZE_A4" },
      { name: "PAGE_SIZE_LETTER" },
      { name: "MARGINS_NORMAL" },
      { name: "VERSION" },
    ],
  },
  {
    title: "Browser preview (@office-kit/docx-preview)",
    description: "The companion package's single entry point.",
    entries: [{ name: "previewToDOM", sig: "(source, container, options?) => Promise<Handle>" }],
  },
];

export const apiTotalCount: number = apiGroups.reduce((n, g) => n + g.entries.length, 0);
