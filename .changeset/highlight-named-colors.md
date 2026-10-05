---
"@office-kit/docx": minor
"@office-kit/docx-editor": minor
---

fix: highlight is written only as a valid `ST_HighlightColor` name

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
