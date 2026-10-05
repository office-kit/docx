---
"@office-kit/docx-editor": patch
---

The canvas now lays out tabs against the paragraph's tab stops, as Word does:

- Custom stops are inherited through styles, and `clear` stops remove inherited ones.
- Past the custom stops, tabs go to the document's default tab interval.
- Right, center and decimal stops align the text that follows the tab.
- Dot, hyphen, underscore, heavy and middle-dot leaders are drawn.

Typing a tab now works too. Tab inserts a tab character, and at the start of a list item Tab and Shift+Tab change the item's level. A caret steps over a tab as a single character.

Drop caps (`w:framePr w:dropCap`) float into the lines of the paragraph that follows them, instead of sitting on a line of their own.
