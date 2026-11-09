---
"@office-kit/docx": patch
---

Opening and saving a document no longer drops a no-break space (U+00A0), narrow no-break space (U+202F) or ideographic space (U+3000) that sits alone in its run. French text such as `Nom : Dupont` or `« Oui »` came back as `Nom: Dupont` and `«Oui»`, even in paragraphs that were never edited. Only XML white space (space, tab, carriage return, line feed) between elements is now treated as formatting (#35).
