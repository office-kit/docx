---
"@office-kit/docx": patch
---

fix: saving a document dropped the attributes of `<w:body>`, `<w:p>`, `<w:r>`,
`<w:tbl>`, `<w:tr>` and `<w:tc>`, including `w14:paraId`, `w14:textId` and
`w:rsid*`. They are now kept on save. The parsed nodes carry them as an
optional `attrs` field.
