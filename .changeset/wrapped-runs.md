---
"@office-kit/docx": minor
"@office-kit/docx-editor": patch
---

`wrappedRuns(inline)` reads the runs inside a hyperlink, simple field, inline content control, smart tag, custom XML or bidi override. The editor canvas uses it so link text shows its own formatting, as Word does: the Hyperlink style for an ordinary link, and plain text for a table of contents entry, instead of always blue and underlined.
