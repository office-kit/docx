---
"@office-kit/docx": patch
"@office-kit/docx-editor": patch
---

fix: find / replace reach table cells, unit font sizes read correctly, and range state sees middle runs

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
