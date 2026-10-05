---
"@office-kit/docx-editor": patch
---

`setFontCommand` now treats East Asian fonts the way Word's font box does:

- An East Asian font, such as 游明朝, ＭＳ ゴシック or SimSun, is applied to East Asian text as well as Latin text.
- A Latin font, such as Arial, changes only the Latin text, so Japanese text keeps its font.

Two new exports, `isEastAsianFont` and `hasEastAsianText`, expose the checks this uses.
