---
"@office-kit/docx": minor
"@office-kit/docx-editor": minor
---

feat: Word's Insert tab — fields, links, symbols, equations, pages and page numbers

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
