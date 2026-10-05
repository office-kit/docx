---
"@office-kit/docx": minor
"@office-kit/docx-editor": minor
---

feat: Word's References and Mailings features, with field results computed so a saved document reads correctly before Word updates it.

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
