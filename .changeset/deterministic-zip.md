---
"@office-kit/docx": patch
---

`toUint8Array` now produces the same bytes for the same document whenever it is called. The ZIP entries used to carry the time of saving, so two saves of an unchanged document could differ; every entry is now dated 1980-01-01 00:00 (the earliest ZIP date). Word and other readers ignore these dates.
