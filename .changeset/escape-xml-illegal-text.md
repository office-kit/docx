---
"@office-kit/docx": patch
---

Text containing `]]>` no longer produces an ill-formed `word/document.xml`; `>` is now always written as `&gt;`. Text containing a character that XML 1.0 cannot represent (U+0000 to U+0008, U+000B, U+000C, U+000E to U+001F, a lone surrogate, U+FFFE or U+FFFF) now throws a `RangeError` instead of producing a file that Word and other XML readers fail to open. Remove such characters before saving. This also applies to header, footer, chart, SmartArt, bibliography, theme and mail merge text (#37).
