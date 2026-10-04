# Editor UI screenshots (2026-10-04)

Screenshots of the `/editor` route for PR #21. They show the source tree of
the commit that adds this folder (the folder itself is the only change in that
commit). They were taken from a local dev server
(`pnpm --filter word-kit-site dev`) with the packages built from that same
source, in an isolated Chromium at a 1280 × 800 viewport, on
2026-10-04 around 01:50 UTC. The documents are small generated fixtures with
no real content.

No "before" screenshots: capturing the previous commit (`15aec19`) with the
same fixtures would need a second build, which was not done.

| File                            | What it shows                                                                                                                                                                                                                                         |
| ------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `01-ribbon-mixed-font.png`      | Runs `Aaaa` (Arial), `Bbbb` (Times New Roman), `Cccc` (Arial), selected from inside the first to inside the last. The Font field's value is "Mixed" (read from the DOM) because the middle run differs; before this PR it showed "Arial" (ends only). |
| `02-named-highlight.png`        | `Hello` highlighted with the named color `darkYellow`, chosen from the highlight list that replaces the color picker (`w:highlight` accepts only named colors). The caret is in the second line so the highlight is not covered by the selection.     |
| `03-replace-all-table-cell.png` | Replace All `DRAFT` → `FINAL` on a document whose body starts with a table: "Replaced 3 occurrences.", including the table cell. Before this PR the cell was skipped (2 occurrences).                                                                 |

Known issue visible here, not fixed in this PR: the site sets
`color-scheme: dark` on the root, and the editor ribbon gives its selects and
the size box a white background without a text color, so the value of an
_enabled_ select (Font, highlight, Styles) is drawn in the dark scheme's light
text color and is not readable (01, 02). Disabled controls use the grey
disabled color and are readable (03, "No selection"). The same CSS is in
`15aec19`.
