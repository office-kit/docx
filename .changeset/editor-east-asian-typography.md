---
"@office-kit/docx-editor": patch
---

The canvas now lays out Japanese and Chinese documents the way Word does. Line heights and spacing were checked against Japanese Word for Mac 16:

- **Document grid (§17.6.5).** Lines snap to whole grid lines; for example, 10.5 pt 游明朝 takes one 18 pt line and 28 pt takes three. A character grid spaces the characters out to the grid pitch.
- **East Asian units.** Indents in characters and spacing in lines are converted to points.
- **Single line height.** 游明朝 and 游ゴシック use the font's own single-line height (1.447 × the size).
- **Character formatting.**
  - Emphasis marks (傍点)
  - Fit Text (均等割り付け), including a width shared by runs with the same id
  - Character scale (文字の拡大/縮小)
  - Combine Characters / Two Lines in One (組み文字・割注)
  - Horizontal-in-vertical (縦中横)
- **Paragraph settings.** Distributed alignment; kinsoku line breaking; word wrap; hanging punctuation; and the automatic space between Japanese and Latin text or numbers (`autoSpaceDE` / `autoSpaceDN`).
- **Fonts.** Japanese and Chinese font names (游明朝, ＭＳ 明朝, メイリオ, 宋体 …) fall back to the system's Japanese or Chinese fonts.
