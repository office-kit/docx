---
"@office-kit/docx-editor": minor
---

feat: the editor reports and writes formatting the way Microsoft Word does

- New `createStyleResolver(doc)` returns a paragraph's and a run's effective
  formatting: document defaults, the paragraph style's `basedOn` chain, the
  character style and direct formatting, with theme fonts looked up. The canvas
  renders these values, so headings take their look from `styles.xml`.
  Also new: `pageGeometry(doc)`, `highlightCss(value)`, `RunRef`.
- Bold / Italic / Strikethrough / Underline toggles and the alignment buttons
  read the effective value. Turning off formatting that comes from a style
  writes an explicit off value (`<w:b w:val="0"/>`, `<w:u w:val="none"/>`),
  as Word does.
- `clearFormattingCommand` now matches Word's Clear All Formatting: a caret
  resets its paragraph to Normal; a range clears every character property
  except the highlight and resets the paragraphs whose mark it includes.
  Previously it wrote explicit `false` values and left color, size and font.
- New `indentStepCommand` (Increase / Decrease Indent in 0.5 in steps),
  `setLineSpacingCommand` with `lineSpacingOf`, and `setVertAlignCommand`
  (Subscript / Superscript).
- The canvas uses the section's page size and margins, renders paragraph
  indents and spacing, and draws tables with their own width, columns, cell
  margins and borders (none when the document defines none).
