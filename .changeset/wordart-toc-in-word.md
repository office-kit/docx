---
"@office-kit/docx": patch
---

Fix two problems that showed up when the file was opened in Word:

- `addWordArt` now writes the text-path shape type (`_x0000_t136`) that the shape refers to. Without it, Word drew a broken-picture placeholder.
- In a section with no `w:pgMar`, tables of contents, indexes and other generated content now place their right tab at Word's default margins (1 in each side). Before, the tab sat at the page edge.
