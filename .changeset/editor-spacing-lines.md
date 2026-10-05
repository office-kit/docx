---
"@office-kit/docx-editor": patch
---

`layoutSpacingCommand` takes spacing in lines (`beforeLines` / `afterLines`, hundredths of a line), the unit Japanese Word's 段落前 / 段落後 boxes use. It also writes the equivalent twips at the section's grid pitch. `spacingLineTwips` returns that pitch.

`layoutIndentCommand` and `layoutSpacingCommand` keep a paragraph's existing character and line units (`leftChars`, `beforeLines` …). Setting a value in twips clears the matching unit, so the new value shows. Before, a unit that was set kept overriding the new twips value.
