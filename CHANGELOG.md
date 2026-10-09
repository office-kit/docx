# Changelog

## 0.2.1

### Patch Changes

- a5fe1b8: `toUint8Array` now produces the same bytes for the same document whenever it is called. The ZIP entries used to carry the time of saving, so two saves of an unchanged document could differ; every entry is now dated 1980-01-01 00:00 (the earliest ZIP date). Word and other readers ignore these dates.

## 0.2.0

### Minor Changes

- 48ed589: `contentControlBlocks(block)` reads the paragraphs and tables inside a block-level content control (`<w:sdt>`), the wrapper Word puts around cover pages, watermarks and tables of contents. The editor canvas uses it to show that content (read-only) instead of a "preserved content" placeholder, so watermarks and cover pages now appear on the page.
- 48ed589: `addTableOfContents`, `appendMergeField` and `addBookmark` are deprecated in favour of `insertTableOfContents`, `insertMergeField` and `insertBookmark`, which do the same and more (any position, computed TOC entries, bookmarks over part of a paragraph). The deprecated functions keep working until the next major release.
- 49e6e8a: `setParagraphIndent` and `setParagraphSpacing` take East Asian units, as Japanese and Chinese Word write them:
  - Indents in hundredths of a character: `leftChars`, `rightChars`, `firstLineChars` and `hangingChars` ("2 字" is `200`).
  - Spacing in hundredths of a line: `beforeLines` and `afterLines` ("0.5 行" is `50`).

- 48ed589: Headers, footers and notes can now be read and edited as stories of their own.
  - `@office-kit/docx`:
    - `storyBody` / `storyView` give typed access to the body of a header, footer, footnote, endnote or comment. Edits to it are saved with the document.
    - `sectionProperties`, `resolveHeaderFooter`, `ensureHeaderFooter`, `isHeaderFooterLinked` and `setHeaderFooterLinked` resolve a section's first, even or default header or footer, including what it inherits from the previous section, and support "Link to Previous".
  - `@office-kit/docx-editor`:
    - A `DocPosition` may carry a `story`, so selection, typing and formatting commands work inside headers, footers and notes.
    - New `stories.*` commands:
      - turn "Different First Page" and "Different Odd & Even Pages" on and off;
      - set the header and footer distances;
      - link a header or footer to the previous section.
    - The new layout module (`paginate`, `documentSections`, `createFlagResolver`, `formatNumber`) breaks a document into pages. It handles:
      - sections and columns;
      - keep with next, keep lines together and widow/orphan control;
      - table rows split across pages, with repeated header rows;
      - footnote areas.

- 48ed589: Add Word's Design and Layout features.
  - `@office-kit/docx`: section layout (`getSectionProperties`, `setSectionProperties` and `insertSectionBreak` for columns, line numbering, page borders, the document grid, vertical alignment, text direction and section start), with a `scope` for `setPageSize`, `setPageMargins` and `setPageOrientation`. Also themes (`setTheme`, `setThemeColors`, `setThemeFonts`, `setThemeEffects`, `getTheme`, plus the built-in Office themes and schemes), style sets (`applyStyleSet`, `currentStyleSet`), default paragraph spacing, page colour including gradient fills (`setPageColor`), and text or picture watermarks (`setWatermark`, `getWatermark`). Document settings are now written in schema order.
  - `@office-kit/docx-editor`: `design.*` and `layout.*` commands for the Design and Layout ribbon tabs and the Page Setup, Columns, Line Numbers, Hyphenation, Text Direction and Page Borders dialogs. Section commands take a `target` of the whole document, the caret section or "this point forward". Also adds the `resolveSectionLayout`, `resolvePageBackground`, `resolveTheme` and `themePalette` helpers.

- 48ed589: Word's Home tab formatting.
  - `@office-kit/docx` adds character and paragraph formatting helpers:
    - Theme colours (`setRunColor`) and every underline style with an underline colour (`setRunUnderline`).
    - Theme fonts (`setRunFont`), character shading and borders (`setRunShading`, `setRunBorder`).
    - Paragraph border sides (`setParagraphBorder`) and tab stops (`setParagraphTabs` / `getParagraphTabs`).
    - Phonetic guides (`buildRubyRun` / `readRuby`) and enclosed characters (`buildEnclosedCharacterRuns`).
    - Custom list definitions with restart (`addListDefinition`, `restartList`).
    - Word's built-in styles created on first use (`builtinStyles`, `ensureBuiltinStyle`) and style editing (`updateStyleFormatting`).
  - Saved documents now always list property children in the order the schema requires, which Word needs to open the file.
  - `@office-kit/docx-editor` adds Home tab commands:
    - Font and Paragraph dialog patches, Change Case, Format Painter, Asian layout and Fit Text.
    - Border presets, paragraph sorting, and the Bullet / Numbering / Multilevel libraries with restart and continue.
    - Applying, creating, modifying and deleting styles.
    - Word-style Find and Replace with wildcards, special characters and format filters.
  - The canvas now draws list numbers, text effects, run shading and borders, paragraph borders, and ruby.

- 48ed589: feat: Word's Insert tab — fields, links, symbols, equations, pages and page numbers

  `@office-kit/docx` can now place what Word's Insert tab places at a character offset of a paragraph:
  - Complex fields (`insertField`, `buildComplexField`) with computed results. DATE/TIME (`\@` pictures with localized month and day names via `lang`), document properties, statistics, `=` formulas, IF/COMPARE, SEQ, REF/PAGEREF/NOTEREF, USERNAME and more. `updateFields` recomputes them, `complexFields` lists them, and `insertFormField` writes legacy form fields.
  - Links with ScreenTips, target frames and in-document anchors (`insertHyperlink`, `editHyperlink`, `removeHyperlink`, `paragraphHyperlinks`).
  - Range bookmarks (`insertBookmark`) and Word's hidden `_Ref` bookmarks for cross-references (`ensureReferenceBookmark`).
  - Symbol-font characters as `w:sym` and the special hyphens (`insertSymbol`).
  - Equations as OMML from a linear format (UnicodeMath), and back (`insertEquation`, `buildEquation`, `equationLinear`, `setEquation`).
  - Drop caps (`setDropCap` / `getDropCap`) and VML signature lines (`insertSignatureLine`).
  - Cover pages, header and footer content, page numbers and their format (`pgNumType`), and merging another document's body with its styles, lists and relationships (`insertDocumentContent`).

  `@office-kit/docx-editor` adds commands for all of these, plus:
  - Queries the Insert dialogs need: headings, bookmarks, notes, numbered items and captions to reference, the link at the caret, and the equation under a rendered element.
  - Canvas rendering of field results (with a field-code view), symbol-font characters, and equations as MathML.

- 48ed589: Pictures, charts and object arrangement.

  `@office-kit/docx` adds a picture API (`src/api/picture.ts`):
  - `readDrawing` reads a `<w:drawing>`: wrap, anchor position, crop, shape, border, effects, color adjustments and image.
  - Arrange: `setDrawingWrap`, `setDrawingPosition`, `setDrawingAnchorOptions`, `arrangeDrawing` (Bring Forward / Send Backward and their variants), `setDrawingTransform`, `setDrawingName`, `setDrawingHidden`, `setDrawingAspectLock`, `setDrawingHyperlink` and `removeDrawing`.
  - Picture formatting: `setPictureCrop`, `setPictureGeometry`, `setPictureOutline`, `setPictureEffects`, `setPictureColorAdjustments`, `changePicture` and `resetPicture`.
  - Charts: `addChartRun`, `readChart` and `setChart` write a real chart part with an embedded workbook.
  - `imagePixelSize` and `imageNaturalSizeEmu` read image dimensions.

  Everything is written as ECMA-376 markup, with no Office extension namespaces.

  Fixed: `setImageSizeEmu` no longer writes into an `a:ext` that sits inside an extension list.

  `@office-kit/docx-editor` adds:
  - Rendering for floating and inline pictures and charts: wrapping, z-order, rotation, crop, effects and blip adjustments.
  - `layoutFloatingObjects` and `floatFrameStart` for placing floating objects in the browser.
  - Picture, arrange and chart commands, including `drawing.layout` for the Layout Options dialog and `chart.insert` / `chart.edit`.

- 48ed589: feat: Word's References and Mailings features, with field results computed so a saved document reads correctly before Word updates it.
  - **References** (`@office-kit/docx`):
    - Tables of contents and tables of figures: `insertTableOfContents`, `removeTableOfContents`, `setTocLevel`, and `updateTables`. Page numbers come from an optional page provider.
    - Footnotes and endnotes: `insertNote`, `noteMarks` (numbering per `w:footnotePr` / `w:endnotePr`), `setNoteProperties`, `noteText` / `setNoteText`, and `convertNotes`.
    - Citations and bibliography: sources are stored in the §22.6 customXml part. Use `setBibliographySources`, `setBibliographyStyle`, `insertCitation` and `insertBibliography`. APA and MLA are fully supported; Chicago, IEEE and the numeric styles are approximated.
    - Captions: `insertCaption` writes SEQ fields with chapter numbers. `addCaptionLabel` adds a caption label.
    - Index: `markIndexEntry`, `markAllIndexEntries` and `insertIndex`.
    - Table of authorities: `markAuthorityCitation` and `insertTableOfAuthorities`.
  - **Mailings** (`@office-kit/docx`):
    - Envelopes and labels: `addEnvelope` and `createLabelDocument`.
    - Mail merge setup: `setMailMergeDocumentType`, plus CSV recipient lists via `parseRecipientCsv` and `attachRecipientList`. Match Fields and recipient inclusion are stored in `w:odso`.
    - Merge fields and rules: `insertMergeField`, `insertAddressBlock`, `insertGreetingLine`, and `insertMergeRule`, which covers ASK, FILLIN, IF, MERGEREC, MERGESEQ, NEXT, NEXTIF, SET and SKIPIF.
    - Preview and merge: `previewMailMerge` and `mergeToNewDocument`, which puts one section per record.
  - **Editor** (`@office-kit/docx-editor`): References and Mailings commands for all of the above, plus `findNoteReference` and `selectedText` queries for the ribbon UI.

- 48ed589: Review and View features.

  `@office-kit/docx`:
  - Tracked changes: `revisions`, `acceptRevisions` / `rejectRevisions` by id, and recording edits as tracked (`insertTrackedText`, `deleteTrackedText`, `trackParagraphMark`, `trackRunFormatChange`, `trackParagraphFormatChange`). Runs inside `<w:ins>` / `<w:del>` now parse as runs that carry `revision`, so they are editable and save back as the same wrappers.
  - `compareDocuments`: diffs two documents into a new document of tracked changes.
  - Comments: `comments`, `setCommentText` and `removeComment`. `addComment` accepts a `range`.
  - Proofing: `getRunLanguage` / `setRunLanguage`, plus a Word-compatible `wordCount`.
  - `checkAccessibility` and `contrastRatio`.
  - Document protection: `protectDocument` / `unprotectDocument` with ECMA-376 SHA-512 password hashes, `writeProtection` (Always Open Read-Only, password to modify), and editable exceptions.
  - `documentView` / `documentZoom`, custom properties (`customProperties`, `setCustomProperty`, `removeCustomProperty`) and the `hyperlinkBase` app property.
  - Settings written through `setDocumentSettingOnOff` / `Val` now follow the schema order of `CT_Settings`.

  `@office-kit/docx-editor`:
  - While Track Changes is on, typing, deleting, splitting and merging paragraphs, and formatting are recorded as revisions.
  - Review commands: comments, accept / reject, language, protection.
  - Outline commands: level, promote / demote, move.
  - Navigation to the previous / next change or comment.
  - Review markup classes on the rendered document.
  - `protectionRefusal` / `isEditingLocked`, so a UI can honour Restrict Editing.

- 48ed589: Shapes, text boxes, WordArt, ink and SmartArt.
  - `@office-kit/docx`: build and edit VML shapes (`addShape` with about 120 Word presets, `addWordArt`, ink strokes, `addDrawingCanvas`, groups), with fill (solid, gradient, pattern, picture), outline, shadow, layout, wrapping, z-order, alt text, linked text boxes and vertical text. Insert and edit SmartArt graphics (`addSmartArt`, `setSmartArtNodes`, layouts, colours, styles, size, `removeSmartArt`) as ECMA-376 DrawingML diagram parts. Existing VML from Word documents is read, including shape types.
  - `@office-kit/docx-editor`: renders VML shapes, text boxes, WordArt, ink and SmartArt on the canvas (`renderPictHtml` is exported for watermarks), with commands for every Shape Format, SmartArt Design and Draw tab action, editable text boxes, and Ink to Shape recognition (`recognizeInkShape`).

- 48ed589: Tables, the way Word edits them.

  `@office-kit/docx` gains Word's table operations: insert and delete rows and columns at any position, delete cells (shifting left or up), merge and split cells, split a table, set table / cell widths, alignment, indent, layout, cell spacing, default and per-cell margins, text direction, floating position and alt text, distribute and AutoFit columns, apply borders to any cell range (including diagonals), sort rows, convert text to a table and back with a chosen separator (`unwrapTable` now takes `{ separator }`), and evaluate table formulas (`=SUM(ABOVE)` …). It also ships Word's 105 built-in table styles (Table Grid, Plain / Grid / List Tables in every accent), colored from the document's theme, with Table Style Options (`w:tblLook`) and per-region style formatting. Table property setters now write their children in schema order.

  `@office-kit/docx-editor` adds commands for every Table Design and Table Layout action, renders tables with their style's conditional formatting (header / total rows, first / last columns, banding with the style's band sizes, corner cells), merged cells, row heights, cell margins, vertical alignment, text direction, diagonal borders, cell spacing and nested tables, and exposes the cell-range selection, Tab navigation, gallery previews and a Table Properties snapshot.

- 48ed589: fix: highlight is written only as a valid `ST_HighlightColor` name

  `<w:highlight w:val>` accepts only 16 named colors and `none` (ECMA-376
  §17.18.40). Before this fix, `setRunFormat`, `appendTextRun`, `setParagraphText`
  and `setTableCellText` wrote any string there (e.g. `FFFF00`), producing a file
  that does not conform to the schema. They now throw a `RangeError`, without
  changing the document, for any other value. The new `HIGHLIGHT_COLORS` export
  (with the `HighlightColor` type) lists the accepted values. `getRunFormat` still
  returns whatever a file holds.

  In the editor, `setHighlightCommand` takes `{ color: HighlightColor }` instead of
  a hex string and rejects anything else before changing the document. The site
  ribbon's highlight control is now a named list instead of a color picker.

- 48ed589: fix: run color and font size are written only as schema-valid values

  `setRunFormat`, `appendTextRun`, `setParagraphText` and `setTableCellText` used
  to write any `color` string and any `fontSizeHalfPoints` number into
  `<w:color w:val>` / `<w:sz w:val>`. Values like `"red"`, `"#FF0000"`, `-1`,
  `23.5` or `NaN` produced files that do not conform to the schema. These writers
  now throw a `RangeError`, before changing anything, unless:
  - `color` is six hex digits without `#` (either case) or `"auto"` (ST_HexColor);
  - `fontSizeHalfPoints` is a non-negative safe integer (ST_HpsMeasure). `0` stays
    valid; there is no application-specific upper bound.

  Untyped JavaScript callers get the same check on the runtime type: a `color`
  that is not a string (e.g. the number `123456`, `null`) or a size that is not a
  number (e.g. `"24"`) is rejected rather than stringified.

  One case that used to pass silently now throws: `appendTextRun(p, text, { color: "" })`
  ignored the empty color, and now rejects it like `setRunFormat` does.
  `getRunFormat` still returns whatever a file holds, and untouched values are
  saved unchanged.

  In the editor, `setColorCommand` (a leading `#` is still dropped) and
  `setFontSizeCommand` (points, rounded to half-points) reject invalid values
  before the document changes, with undo and redo history kept.

  The editor canvas also no longer lets a crafted file break out of a run's
  `style` attribute: the attribute is HTML-escaped, only six-digit hex colors
  become CSS, and quotes, backslashes and line breaks are dropped from font names
  when rendering. The document itself is not changed.

- 48ed589: `wrappedRuns(inline)` reads the runs inside a hyperlink, simple field, inline content control, smart tag, custom XML or bidi override. The editor canvas uses it so link text shows its own formatting, as Word does: the Hyperlink style for an ordinary link, and plain text for a table of contents entry, instead of always blue and underlined.

### Patch Changes

- 48ed589: feat: add `@office-kit/docx-editor` — an MS Office-like WYSIWYG editing core

  New package that turns `@office-kit/docx` into a Word-like editor: an
  `EditorModel` with undo/redo, a command layer (text/paragraph/structure/list/
  table/image/style/section/header-footer/references/review) where every mutation
  routes through the `@office-kit/docx` public API, an AST→HTML canvas renderer,
  and DOM-selection mapping. A SvelteKit ribbon UI ships at the site's `/editor`
  route.

  The editing surface is interactive: Enter splits the paragraph at the caret and
  Backspace/Delete merge across paragraph boundaries (`splitParagraphCommand` /
  `mergeBackCommand`, backed by the library's `splitParagraphAt` /
  `mergeParagraphIntoPrevious`); the caret is preserved across structural
  re-renders; Ctrl+Z/Y drive undo/redo; paste inserts plain text (multi-line →
  paragraphs) keeping the AST in sync; and the UI adds live word/character count,
  find & replace, zoom, caret-synced font/size/style, and heading-styled rendering
  so the document reads with real visual hierarchy.

  Character formatting is now selection-precise: bolding (or any run format on) a
  partial selection splits runs at the selection boundaries and formats only the
  selected characters instead of the whole line — backed by the library's
  `isolateParagraphRunRange` / `runTextLength`. The ribbon UI is localized into
  six languages (English, 日本語, Español, Français, Deutsch, 中文) with a
  language switcher.

  Completeness against the WordprocessingML spec is enforced by a capability
  ledger: every element in the ECMA-376 schema universe is classified as
  `edit` / `render` / `preserve`, and a test fails the build if any element is
  unclassified or an editable element's command goes missing.

  The library gains generic property setters so the editor can reach the long
  tail of WordprocessingML formatting without a bespoke function per element:
  - run / paragraph: `setRunOnOff` / `setRunValProp` / `getRunProp` and the
    `setParagraphOnOff` / `setParagraphValProp` / `getParagraphProp` equivalents;
  - any properties container: `setElementOnOff` / `setElementValProp` /
    `getElementProp` / `makePropsElement` (for `<w:tcPr>` / `<w:trPr>` /
    `<w:tblPr>` / `<w:sectPr>`);
  - document settings: `setDocumentSettingOnOff` / `setDocumentSettingVal` /
    `getDocumentSetting` (creates `word/settings.xml` on demand);
  - style definitions: `setStyleOnOff` / `setStyleValProp` / `getStyleProp`;
  - numbering levels: `setNumberingLevelVal` / `setNumberingLevelOnOff` /
    `getNumberingLevelProp`;
  - images: `imageDrawings` / `setImageSizeEmu` / `setImageAltText` /
    `getImageInfo` (resize and alt-text existing pictures);
  - raw XML nodes: `setElementAttr` / `getElementAttr` / `childElementsOf` /
    `appendChildElement` — the universal escape hatch that lets the editor edit
    any element (DrawingML shape geometry, OMML math, legacy VML) by attribute or
    child, since these live inline in `document.xml` and flush through the
    document AST on save;
  - arbitrary XML parts: `xmlPartNames` / `getRawPartRoot` / `markRawPartDirty` —
    open, edit, and re-serialize any XML part (fontTable, settings, styles,
    numbering, comments, foot-/endnotes, headers, footers, webSettings, docProps),
    so every OOXML element in the whole package is reachable and editable;
  - caret-level structure: `splitParagraphAt` / `mergeParagraphIntoPrevious` —
    split a paragraph at a run/offset and merge a paragraph into the previous one
    (the primitives behind Enter and Backspace-at-start);
  - selection formatting: `isolateParagraphRunRange` / `runTextLength` — isolate a
    character range into its own run(s) so formatting applies to exactly the
    selected text.

  Also newly exported from `@office-kit/docx` (they are part of the public
  surface as parameter/return/AST types): `BuildStyleOptions`, `BuildTableOptions`,
  `DocumentCoreProperties`, `DocumentAppProperties`, and the raw XML AST types
  `XmlElement` / `XmlNode` / `XmlAttr` / `QName`.

  Commands are atomic: `runCommand` returns the command's result, and a command
  that throws (including rejected input such as an out-of-range table size or a
  non-http(s)/mailto hyperlink) leaves the document, selection and redo history
  unchanged (`EditorModel.abortEdit`). `replaceAllCommand`, `insertTextCommand` and
  `deleteSelectionCommand` route Find & Replace, paste, and typing or deleting
  over a selection (also across paragraphs) through the same undoable path.
  Consecutive keystrokes undo as one step, Shift+Enter breaks the line at the
  caret, and hyperlinks and field results are visible on the canvas. Enter,
  Backspace and Delete inside a table cell split and join that cell's
  paragraphs. IME input that starts over a selection replaces it in one undo
  step. `splitParagraphAt` moves a paragraph's
  section break to the second half instead of duplicating it.

- 48ed589: fix: saving a document dropped the attributes of `<w:body>`, `<w:p>`, `<w:r>`,
  `<w:tbl>`, `<w:tr>` and `<w:tc>`, including `w14:paraId`, `w14:textId` and
  `w:rsid*`. They are now kept on save. The parsed nodes carry them as an
  optional `attrs` field.
- 48ed589: fix: find / replace reach table cells, unit font sizes read correctly, and range state sees middle runs
  - `findText` and `replaceText` (and the body part of `findTextEverywhere` /
    `replaceTextEverywhere`) skipped every paragraph inside a table cell, despite
    promising "every paragraph" / "every occurrence". They now include cell
    paragraphs, in document order, so counts and replacements cover tables. A
    table nested inside a cell is still kept as raw XML and not searched.
  - `getRunFormat` read `<w:sz w:val="12pt">` as 12 half-points (6 pt) because it
    ignored the unit. Universal measures (`pt`, `pc`, `pi`, `in`, `cm`, `mm`) are
    now converted exactly, so `"12pt"` reads as 24. A value that is not a whole
    number of half-points (`"3mm"`) or is invalid (`"-4"`, `"12px"`) leaves
    `fontSizeHalfPoints` unset instead of being guessed; the XML keeps the value
    as written.
  - In the editor, `runsInRange` returned every run of the paragraphs at both
    ends of a range, not just the runs the range touches as documented. It now
    returns exactly the covered runs (a collapsed caret still returns its
    paragraph's runs). The site ribbon uses it, so a differently formatted run in
    the middle of a selection shows the field as "Mixed", and choosing a value
    applies it to the whole range.

- 48ed589: Fix two problems that showed up when the file was opened in Word:
  - `addWordArt` now writes the text-path shape type (`_x0000_t136`) that the shape refers to. Without it, Word drew a broken-picture placeholder.
  - In a section with no `w:pgMar`, tables of contents, indexes and other generated content now place their right tab at Word's default margins (1 in each side). Before, the tab sat at the page edge.

## 0.1.0

### Minor Changes

- 3302af3: Initial public release.

  `@office-kit/docx` is an OOXML-compliant (ECMA-376) `.docx` generation and
  editing library for browsers and Node.js. It ships as a single self-contained
  package: the OPC, XML, and WordprocessingML layers are bundled in, not
  published separately. `@office-kit/docx-preview` renders any `Docx` value as a
  read-only DOM tree by wrapping the OSS `docx-preview` renderer.

The format is loosely based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/)
and the project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).
Each user-visible change is also tracked via [Changesets](https://github.com/changesets/changesets);
this file is a hand-curated overview.

## Unreleased

### Changed

- **No-classes API.** Every package now exposes plain data types and
  standalone functions instead of classes. The motivation is
  tree-shaking: classes carry every method along with the prototype,
  so a bundler can't drop unused operations once an instance escapes.
  - `Docx` is an `interface` (just `{ opc, document, partName, … }`).
  - `Docx.create(…)` → `createDocx(…)`. `Docx.open(bytes)` → `openDocx(bytes)`.
    `Docx.fromBlob(blob)` → `fromBlob(blob)`.
  - `doc.appendParagraph(text)` → `appendParagraph(doc, text)`. Same
    pattern for every previous method.
  - Previous getters become functions: `doc.paragraphs` → `paragraphs(doc)`,
    `doc.statistics` → `statistics(doc)`, etc.
  - Previous setters become `setX` functions: `doc.title = x` →
    `setTitle(doc, x)`. `doc.opc`, `doc.document`, `doc.partName` stay
    as direct property access.
  - `OpcPackage`, `ContentTypesIndex`, `RelationshipSet` are likewise
    plain interfaces backed by standalone `addPart`, `getPart`,
    `allRelationships`, `packageRelationships`, … functions.
  - `XmlParseError` is no longer a class. Use `XmlParseError.is(e)` for
    narrowing (in place of `e instanceof XmlParseError`).

  A minimal `createDocx + appendParagraph + toUint8Array` slice bundles
  to ~42 KB minified; the full surface is ~131 KB. CI enforces both the
  byte budget and that no feature-specific string literals from unused
  branches leak into the minimal bundle (see `scripts/check-tree-shake.mjs`).

### Added

- `@office-kit/docx` — the public `Docx` interface and standalone-function API.
  Ships as a single self-contained package; the layers below are bundled in
  (`src/internal/`), not published separately:
  - **OPC layer** (`src/internal/opc`) — Open Packaging Conventions
    reader/writer with byte-stable round-trip for untouched parts, built on
    `fflate` (`OpcPackage` interface + `readOpcPackage`, `writeOpcPackage`,
    `addPart`, `getPart`, `allRelationships`, `packageRelationships`, …).
  - **XML layer** (`src/internal/xml`) — namespace-aware XML parser/serializer
    that preserves attribute order, prefixes, `xml:space="preserve"`, CDATA,
    comments, and PIs.
  - **WordprocessingML layer** (`src/internal/wordprocessingml`) — WML AST plus
    parsers, writers, and builders for paragraphs, runs, tables, styles,
    numbering, headers/footers, sections, comments, footnotes/endnotes,
    hyperlinks, fields, and document properties.
- `@office-kit/docx-preview` — browser-side read-only preview. Single function entry,
  `previewToDOM(source, container, options?) → Promise<Handle>`. v0
  implementation wraps the OSS `docx-preview` (Apache-2.0). The wrap is
  intentional and final; see `docs/PLAN-PREVIEW.md` for the rationale.

#### Authoring (function API on `@office-kit/docx`)

- Lifecycle: `createDocx({ paragraphs? })`, `openDocx(bytes)`,
  `fromBlob(blob)`, `toUint8Array(doc)`, `toBlob(doc)`, `clone(doc)`.
- Paragraphs: `appendParagraph`, `insertParagraphAt`, `removeParagraph`,
  `appendHeading`, `appendPageBreak`, `appendLineBreak`,
  `appendSectionBreak`, `clearBody`.
- Inline / text: `replaceText`, `replaceTextEverywhere`, `findText`,
  `findTextEverywhere`, `appendTextRun`, `setParagraphText`,
  `paragraphText`, `setRunFormat`, `clearRunFormat`, `getRunFormat`,
  `setParagraphAlignment`, `setParagraphIndent`, `setParagraphSpacing`.
- Styles + numbering: `addStyle`, `removeStyle`, `listStyles`,
  `ensureHeadingStyles`, `addBulletList`, `addNumberedList`,
  `applyListToParagraph`.
- Tables: `addTable`, `tables`, `removeTable`, `removeAllTables`,
  `unwrapTable`.
- Images: `addImage`, `addImageRun`, `insertImageInto`, `images`,
  `replaceImage`, `removeAllImages`.
- Headers / footers / sections: `addHeader`, `addFooter`,
  `addPageNumberFooter`, `setPageSize`, `setPageMargins`,
  `setPageOrientation`, `headers`, `footers`,
  `removeAllHeaders`, `removeAllFooters`.
- Comments / footnotes / endnotes: `addComment`, `addFootnote`,
  `addEndnote`, `removeAllComments`, `removeAllFootnotes`,
  `removeAllEndnotes`.
- Hyperlinks + bookmarks: `addHyperlink`, `addInternalHyperlink`,
  `externalHyperlinks`, `setHyperlinkUrl`, `removeAllHyperlinks`,
  `addBookmark`, `removeBookmark`, `removeAllBookmarks`, `bookmarks`.
- Fields: `appendField`, `addTableOfContents`, `appendMergeField`,
  `WORD_FIELDS`, plus the page-number footer helper above.
- Tracked changes: `acceptAllRevisions`, `rejectAllRevisions`.
- Core / app properties: `coreProperties`, `setCoreProperties`,
  `appProperties`, `setAppProperties`, `title`, `author`,
  `setTitle`, `setAuthor`.
- Templates (PowerPoint-style "open a designed base, append content"):
  `mergeStylesFromTemplate`, `findStyleIdByName`,
  `setParagraphStyle`, `imageReferences`, `replaceImageByAltText`.
- Diagnostics: `validate(doc)`, `statistics(doc)`, `outline(doc)`,
  `fields(doc)`.

#### Browser preview (function API on `@office-kit/docx-preview`)

- `previewToDOM(source, container, options?)` renders a `Docx`,
  `Uint8Array`, `Blob`, or `ArrayBuffer` into a DOM container.
  Returns an idempotent `Handle.dispose()` for teardown.
- Options: `classPrefix` (default `"wk-"`), `inWrapper`,
  `breakPages`, `renderFonts`, `experimentalComments`,
  `experimentalChanges`. All overridable.

### Build + tooling

- `tsdown` (rolldown-based) replaces `tsup` for the bundling step.
  Output extensions changed from `.js`/`.d.ts` to `.mjs`/`.d.mts`.
- `pnpm test` now runs `pnpm build` first via the npm-standard
  `pretest` hook so cross-package imports always resolve to fresh
  dist.
- New CI gate: `pnpm check:tree-shake` budgets a minimal
  `createDocx + appendParagraph + toUint8Array` bundle (~42 KB
  minified) against the full surface (~131 KB).
- `pnpm sample` writes 32 demonstration `.docx` files into
  `./samples/` for manual verification in Microsoft Word.
- New perf-smoke test catches accidental quadratic regressions:
  10k-paragraph round-trip in 220 ms locally; budget is 10 s per
  block to leave room for slow CI.

### Implementation notes

- Lossless round-trip: every XML element the library does not yet structure
  is preserved as a `WmlRawBlock`/`WmlRawInline` pass-through with its
  original child position, so re-saving an unmodified template leaves
  Word's "needs repair" prompt out of the picture.
- Verified against the mammoth.js fixture corpus (comments, footnotes,
  endnotes, tables, images, hyperlinks, text boxes, UTF-8 BOM, lists)
  and the python-docx test corpus. The ISO/IEC 29500 Strict variant is
  explicitly out of scope today and remains pass-through only.
- 512 tests, all running in vitest under Node 22 and 24; the public surface
  also runs in modern browsers (no Node-only dependencies in the published
  bundles). The browser-preview tests run under happy-dom (jsdom's
  cross-realm `Uint8Array` confused fflate's type guards).
- CI gate runs typecheck, lint, format check, tests across Node 22 / 24,
  AND the tree-shake budget check. Node 20 was dropped from the matrix
  after it reached end-of-life on 2026-04-30.
