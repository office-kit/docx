---
"@office-kit/docx": minor
---

`setParagraphIndent` and `setParagraphSpacing` take East Asian units, as Japanese and Chinese Word write them:

- Indents in hundredths of a character: `leftChars`, `rightChars`, `firstLineChars` and `hangingChars` ("2 字" is `200`).
- Spacing in hundredths of a line: `beforeLines` and `afterLines` ("0.5 行" is `50`).
