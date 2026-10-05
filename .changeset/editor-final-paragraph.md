---
"@office-kit/docx-editor": patch
---

The editor keeps a paragraph after a table that ends the document, as Word does. Before, there was no way to type below such a table. This affected documents that end with a table, and tables inserted at the end of a document.
