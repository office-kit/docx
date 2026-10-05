---
"@office-kit/docx": minor
"@office-kit/docx-editor": minor
---

Headers, footers and notes can now be read and edited as stories of their own.

- `@office-kit/docx`:
  - `storyBody` / `storyView` give typed access to the body of a header, footer, footnote, endnote or comment. Edits to it are saved with the document.
  - `sectionProperties`, `resolveHeaderFooter`, `ensureHeaderFooter`, `isHeaderFooterLinked` and `setHeaderFooterLinked` resolve a section's first, even or default header or footer, including what it inherits from the previous section, and support "Link to Previous".
- `@office-kit/docx-editor`:
  - A `DocPosition` may carry a `story`, so selection, typing and formatting commands work inside headers, footers and notes.
  - New `stories.*` commands:
    - turn "Different First Page" and "Different Odd & Even Pages" on and off;
    - set the header and footer distances;
    - link a header or footer to the previous section.
  - The new layout module (`paginate`, `documentSections`, `createFlagResolver`, `formatNumber`) breaks a document into pages. It handles:
    - sections and columns;
    - keep with next, keep lines together and widow/orphan control;
    - table rows split across pages, with repeated header rows;
    - footnote areas.
