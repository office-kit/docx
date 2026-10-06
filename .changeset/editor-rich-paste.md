---
"@office-kit/docx-editor": patch
---

Formatted paste and pending formatting:

- New `insertFragmentCommand` and `parseClipboardHtml`: pasting HTML (from Word, a web page, Google Docs or the editor itself) keeps bold, italic, underline, strikethrough, superscript / subscript, text color, headings, bullet and numbered lists with their levels, line breaks, tabs and tables. As in Word, the first pasted paragraph continues the caret's paragraph and a pasted table goes between its halves; inside a table cell a pasted table becomes paragraphs.
- New `releasePendingFormat`: formatting set at a caret (Bold with no selection, say) and left without typing is dropped once the caret moves on, as in Word, instead of leaving an empty run behind.
