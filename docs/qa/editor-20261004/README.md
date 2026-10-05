# Editor UI screenshots (2026-10-04)

Screenshots of the `/editor` route for PR #21, taken from a local dev server
(`pnpm --filter word-kit-site dev`) in an isolated headless Chromium at a
1280 × 800 viewport. The documents are small generated fixtures with no real
content. They are browser screenshots of this editor only; none of them is a
comparison with Microsoft Word.

The current images were taken on 2026-10-04 between 02:06 and 02:10 UTC, with
`site/src/routes/editor/+page.svelte` at blob `fa1f742` (the commit that
updates these images; the rest of the tree is as in `9ff3929`).

The images first committed in `9ff3929` (taken ~01:50 UTC, same fixtures and
viewport) serve as the "before" for the contrast fix: the site root is
`color-scheme: dark`, so the ribbon selects and the find inputs drew their
values in light text on a white background and were unreadable. The editor
now declares `color-scheme: light` on its own root. There is no "before" for
the earlier fixes (mixed font, named highlight, table replace): capturing
`15aec19` would have needed a second build, which was not done.

| File                            | What it shows                                                                                                                                                                                                  |
| ------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `01-ribbon-mixed-font.png`      | Runs `Aaaa` (Arial), `Bbbb` (Times New Roman), `Cccc` (Arial), selected from inside the first to inside the last. The Font field reads "Mixed" because the middle run differs (before the range fix: "Arial"). |
| `02-named-highlight.png`        | `darkYellow` applied to `Hello` from the named highlight list, then the caret placed inside `Hello` (collapsed, so no selection color): the highlight and the ribbon value `darkYellow` are both visible.      |
| `03-replace-all-table-cell.png` | Replace All `DRAFT` → `FINAL` on a document whose body starts with a table: "Replaced 3 occurrences.", including the table cell (before the fix the cell was skipped: 2 occurrences).                          |

Visible but not fixed here: the font-size box is too narrow for its
"Default" placeholder ("Defa").
