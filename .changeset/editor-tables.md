---
"@office-kit/docx": minor
"@office-kit/docx-editor": minor
---

Tables, the way Word edits them.

`@office-kit/docx` gains Word's table operations: insert and delete rows and columns at any position, delete cells (shifting left or up), merge and split cells, split a table, set table / cell widths, alignment, indent, layout, cell spacing, default and per-cell margins, text direction, floating position and alt text, distribute and AutoFit columns, apply borders to any cell range (including diagonals), sort rows, convert text to a table and back with a chosen separator (`unwrapTable` now takes `{ separator }`), and evaluate table formulas (`=SUM(ABOVE)` …). It also ships Word's 105 built-in table styles (Table Grid, Plain / Grid / List Tables in every accent), colored from the document's theme, with Table Style Options (`w:tblLook`) and per-region style formatting. Table property setters now write their children in schema order.

`@office-kit/docx-editor` adds commands for every Table Design and Table Layout action, renders tables with their style's conditional formatting (header / total rows, first / last columns, banding with the style's band sizes, corner cells), merged cells, row heights, cell margins, vertical alignment, text direction, diagonal borders, cell spacing and nested tables, and exposes the cell-range selection, Tab navigation, gallery previews and a Table Properties snapshot.
