---
"@office-kit/docx": minor
"@office-kit/docx-editor": minor
---

Word's Home tab formatting.

- `@office-kit/docx` adds character and paragraph formatting helpers:
  - Theme colours (`setRunColor`) and every underline style with an underline colour (`setRunUnderline`).
  - Theme fonts (`setRunFont`), character shading and borders (`setRunShading`, `setRunBorder`).
  - Paragraph border sides (`setParagraphBorder`) and tab stops (`setParagraphTabs` / `getParagraphTabs`).
  - Phonetic guides (`buildRubyRun` / `readRuby`) and enclosed characters (`buildEnclosedCharacterRuns`).
  - Custom list definitions with restart (`addListDefinition`, `restartList`).
  - Word's built-in styles created on first use (`builtinStyles`, `ensureBuiltinStyle`) and style editing (`updateStyleFormatting`).
- Saved documents now always list property children in the order the schema requires, which Word needs to open the file.
- `@office-kit/docx-editor` adds Home tab commands:
  - Font and Paragraph dialog patches, Change Case, Format Painter, Asian layout and Fit Text.
  - Border presets, paragraph sorting, and the Bullet / Numbering / Multilevel libraries with restart and continue.
  - Applying, creating, modifying and deleting styles.
  - Word-style Find and Replace with wildcards, special characters and format filters.
- The canvas now draws list numbers, text effects, run shading and borders, paragraph borders, and ruby.
