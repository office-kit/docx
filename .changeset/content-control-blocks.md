---
"@office-kit/docx": minor
"@office-kit/docx-editor": patch
---

`contentControlBlocks(block)` reads the paragraphs and tables inside a block-level content control (`<w:sdt>`), the wrapper Word puts around cover pages, watermarks and tables of contents. The editor canvas uses it to show that content (read-only) instead of a "preserved content" placeholder, so watermarks and cover pages now appear on the page.
