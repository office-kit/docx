---
"@office-kit/docx-editor": patch
---

Style resolution follows more of ECMA-376: toggle properties (bold, italic, caps …) toggle across the table, paragraph and character style levels (§17.7.3), list numbers take their level's run properties (§17.9.24), and complex-script text (Arabic, Hebrew, Thai …) uses its own bold, italic, size and font (`bCs`, `iCs`, `szCs`, `w:cs`). The canvas lists the East Asian and complex-script fonts after the Latin one, so each script draws in its own font.
