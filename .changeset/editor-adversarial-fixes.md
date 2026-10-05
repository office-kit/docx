---
"@office-kit/docx-editor": patch
---

Editing fixes found by adversarial testing:

- Deleting or typing over a selection that crosses table cells no longer fails silently: within one table the covered cells are emptied, and a selection that leaves the table deletes the rows it touches, as in Word.
- Backspace at the start of a paragraph after a table no longer joins it to the paragraph above the table. An empty paragraph there is removed and the caret moves to the table's last cell.
- Bold, italic, font and other character formatting with a caret (no selection) now behaves like Word: inside a word it formats the whole word (Japanese words included); elsewhere it applies to the text typed next instead of the whole run.
- `runPoint` places a caret after the placeholder of an empty run, so text typed there lands in that run.
- New `isEmptyParagraph` export.
