---
"@office-kit/docx": minor
"@office-kit/docx-editor": minor
---

Review and View features.

`@office-kit/docx`:

- Tracked changes: `revisions`, `acceptRevisions` / `rejectRevisions` by id, and recording edits as tracked (`insertTrackedText`, `deleteTrackedText`, `trackParagraphMark`, `trackRunFormatChange`, `trackParagraphFormatChange`). Runs inside `<w:ins>` / `<w:del>` now parse as runs that carry `revision`, so they are editable and save back as the same wrappers.
- `compareDocuments`: diffs two documents into a new document of tracked changes.
- Comments: `comments`, `setCommentText` and `removeComment`. `addComment` accepts a `range`.
- Proofing: `getRunLanguage` / `setRunLanguage`, plus a Word-compatible `wordCount`.
- `checkAccessibility` and `contrastRatio`.
- Document protection: `protectDocument` / `unprotectDocument` with ECMA-376 SHA-512 password hashes, `writeProtection` (Always Open Read-Only, password to modify), and editable exceptions.
- `documentView` / `documentZoom`, custom properties (`customProperties`, `setCustomProperty`, `removeCustomProperty`) and the `hyperlinkBase` app property.
- Settings written through `setDocumentSettingOnOff` / `Val` now follow the schema order of `CT_Settings`.

`@office-kit/docx-editor`:

- While Track Changes is on, typing, deleting, splitting and merging paragraphs, and formatting are recorded as revisions.
- Review commands: comments, accept / reject, language, protection.
- Outline commands: level, promote / demote, move.
- Navigation to the previous / next change or comment.
- Review markup classes on the rendered document.
- `protectionRefusal` / `isEditingLocked`, so a UI can honour Restrict Editing.
