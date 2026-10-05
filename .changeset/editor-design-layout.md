---
"@office-kit/docx": minor
"@office-kit/docx-editor": minor
---

Add Word's Design and Layout features.

- `@office-kit/docx`: section layout (`getSectionProperties`, `setSectionProperties` and `insertSectionBreak` for columns, line numbering, page borders, the document grid, vertical alignment, text direction and section start), with a `scope` for `setPageSize`, `setPageMargins` and `setPageOrientation`. Also themes (`setTheme`, `setThemeColors`, `setThemeFonts`, `setThemeEffects`, `getTheme`, plus the built-in Office themes and schemes), style sets (`applyStyleSet`, `currentStyleSet`), default paragraph spacing, page colour including gradient fills (`setPageColor`), and text or picture watermarks (`setWatermark`, `getWatermark`). Document settings are now written in schema order.
- `@office-kit/docx-editor`: `design.*` and `layout.*` commands for the Design and Layout ribbon tabs and the Page Setup, Columns, Line Numbers, Hyphenation, Text Direction and Page Borders dialogs. Section commands take a `target` of the whole document, the caret section or "this point forward". Also adds the `resolveSectionLayout`, `resolvePageBackground`, `resolveTheme` and `themePalette` helpers.
