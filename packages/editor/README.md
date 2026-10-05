# @office-kit/docx-editor

An MS Office-like WYSIWYG editing core for [`@office-kit/docx`](https://github.com/office-kit/docx)
`.docx` documents. Framework-agnostic; a Svelte/SvelteKit UI lives in the
monorepo's `site/` (`/editor` route).

## Design

The editor never serializes OOXML itself. Every mutation runs through a
**command** that calls the `@office-kit/docx` public API, so that library stays
the single source of truth for the document. Open a `.docx`, drive commands from
a ribbon or keyboard, and save a real, valid `.docx` back.

```ts
import { openEditor, runCommand, commands, renderDocumentHtml } from "@office-kit/docx-editor";
import { toUint8Array } from "@office-kit/docx";

const model = openEditor(bytes); // EditorModel over a live Docx
model.setSelection({ anchor: { block: 0 }, focus: { block: 0 } });
runCommand(model, commands.toggleBoldCommand, undefined);
container.innerHTML = renderDocumentHtml(model.doc);
const out = toUint8Array(model.doc); // a valid .docx
```

## The completeness guarantee

"Cover every docx representation" is a _checkable_ claim here, not an aspiration.
Every element in the ECMA-376 schema universe (extracted into
`src/capability/element-universe.json`) is classified in the capability ledger
(`src/capability/ledger.ts`) into exactly one disposition:

- **`edit`** — a real command creates/modifies it (the command id is recorded).
- **`render`** — the canvas shows it faithfully, read-only.
- **`preserve`** — round-tripped losslessly by `@office-kit/docx` (raw XML
  pass-through), not yet surfaced in the UI.

`src/capability/coverage.test.ts` fails the build if any element is unclassified,
or if an `edit` element points at a command that no longer exists. The live
scoreboard is regenerated into [`COVERAGE.md`](./COVERAGE.md). Nothing can
silently fall through: an element is either editable, rendered, or explicitly
preserved with the round-trip test that proves it.

Regenerate the universe from the schemas with:

```sh
node scripts/extract-ooxml-elements.mjs
```

## Command groups

`text`, `paragraph`, `structure`, `list`, `table`, `image`, `style`, `section`,
`headerFooter`, `references`, `review`. See `ALL_COMMANDS` for the full list.

## Editing contract

- **One path.** Every document change goes through `runCommand`, including
  Find & Replace (`replaceAllCommand`), paste and typing over a selection
  (`insertTextCommand`), deleting a range (`deleteSelectionCommand`) and
  Shift+Enter (`insertLineBreakCommand`, at the caret). Collapsed-caret typing
  is synced from the canvas DOM, with consecutive keystrokes grouped into one
  undo step. An IME composition that _starts over a range_ is applied, when
  it ends, as one `insertTextCommand` on that range. Until then the canvas
  stops syncing the DOM into the model. Enter, Backspace and Delete inside a
  table cell split and join paragraphs of that cell only.
- **Highlight is named.** `setHighlightCommand` takes `{ color: HighlightColor }`:
  one of `HIGHLIGHT_COLORS` from `@office-kit/docx` (ECMA-376 §17.18.40
  `ST_HighlightColor`, 16 colors and `"none"`). `"none"` is saved as
  `<w:highlight w:val="none"/>`, an explicit no-highlight. Hex / RGB is
  rejected with a `RangeError` before anything changes: `<w:highlight>` has no
  RGB form, and a hex is never mapped to a "nearest" name. A value another tool
  wrote outside the list is kept in the file, shown as-is in the ribbon (not
  selectable) and rendered without a background.
- **Atomic.** `runCommand` returns the command's result, or `undefined` when
  the command is disabled. If a command throws (including on rejected input:
  table sizes outside 1–1000 × 1–63, hyperlinks that are not
  `http:`/`https:`/`mailto:`, an empty Replace All query, a highlight that is
  not one of `HIGHLIGHT_COLORS`, a color that is not six hex digits or `auto`,
  a font size that is negative or not finite), the document,
  selection and redo history stay as they were. The error is rethrown. Edits
  do not nest: `beginEdit` while another edit is open throws.
- **`model.doc` is not a stable reference.** Undo, redo and a rolled-back
  command _replace_ the `Docx` object. Always read `model.doc` again and do not
  keep an earlier reference: a `Docx` passed to `editorFor(doc)` stops being the
  editor's document after the first undo.
- **Cost.** Each undo step snapshots by serializing and reparsing the whole
  package (`clone`), so one command, or one typing burst, costs O(package
  size). Collapsed-caret typing also rescans every run span of the canvas.
  Expect this to be slow on very large documents.
- **Preservation.** Unmodelled XML is kept verbatim, as are the attributes of
  `<w:body>`, `<w:p>`, `<w:r>`, `<w:tbl>`, `<w:tr>` and `<w:tc>` (`w14:paraId`,
  `w:rsid*`, …). Paragraphs created by a split carry no `<w:p>` attributes, so
  `w14:paraId` is never duplicated. Run halves keep their `w:rsid*` (Word
  allows repeats).

See [`examples/edit-and-save.mjs`](./examples/edit-and-save.mjs) for a
load → edit → save script that uses only the published entry points and can
safely be re-run on its own output.

## Matching Microsoft Word

The `/editor` ribbon follows Word for Mac 16 (light appearance): the Quick
Access Toolbar and a search field in the title bar; the Home, Insert, Layout,
References, Mailings, Review and View tabs with Comments and Download at the
right; a ribbon without group captions; Paste / Cut / Copy; font and size
combo boxes; Increase / Decrease Font Size; Clear All Formatting; Bold,
Italic, Underline (with its styles menu), Strikethrough, Subscript,
Superscript; Text Highlight Color and Font Color split buttons with Word's
palettes; Bullets, Numbering, Decrease / Increase Indent, Show/Hide ¶; the four
alignments; Line and Paragraph Spacing; a style gallery; a status bar with the
word count and a zoom slider. Labels and tooltips use the wording of Word's own
localized UI in all six languages. Colors and sizes were measured from Word's
window.

What the ribbon reports and what the commands write were checked against Word
itself (Word for Mac 16, driven through AppleScript and the saved XML read
back):

- **Effective formatting.** The ribbon shows the style-resolved value, as Word
  does: a caret in a Heading 1 reads Calibri / 20 with Bold and Align Left
  pressed. `createStyleResolver` merges the document defaults, the paragraph
  style's `basedOn` chain, the character style chain and direct formatting;
  theme fonts show as "Calibri (Body)". The canvas renders the same values, so
  headings take their color, size and spacing from `styles.xml`.
- **Toggles over a style.** Bold on a bold Heading 1 writes `<w:b w:val="0"/>`;
  turning it back on removes the override. Underline writes `w:u="none"` the
  same way.
- **Clear All Formatting.** A caret resets its paragraph to Normal and removes
  direct paragraph formatting, leaving the runs. A range removes every
  character property except the highlight, and resets the paragraphs whose
  paragraph mark it includes. Section properties stay.
- **Increase / Decrease Indent** step the effective left indent to the next
  0.5 in; **Line Spacing** writes an auto multiple and keeps the paragraph's own
  space before / after.
- **Page and tables.** The page uses the section's paper size and margins
  (Letter and 1 in when the document has none, as Word saves it), with Word's
  crop marks. Tables use their own width, grid and cell margins and draw only
  the borders the document defines; in Compatibility Mode the cell text, not
  the border, sits on the margin.

Side-by-side screenshots are in `docs/qa/editor-20261005/`.

## Building and consuming the package

The package exports only `dist/index.mjs` / `dist/index.d.mts`. It has no
source alias and no `prepublishOnly`, and `@office-kit/docx` is kept as an
external runtime dependency (not bundled). In a clean checkout, nothing that
imports `@office-kit/docx-editor` — the site, `examples/edit-and-save.mjs` and
editor type checks — resolves until the packages are built in dependency order:

```sh
pnpm install --frozen-lockfile
pnpm build:packages          # core → preview → editor, one after another
pnpm -r run typecheck        # includes packages/editor
pnpm --filter word-kit-site check
node packages/editor/examples/edit-and-save.mjs in.docx out.docx
node packages/editor/examples/edit-and-save.mjs out.docx out2.docx   # re-run: 0 replacements
```

`pnpm release` (CI only) runs `build:packages` before `changeset publish`, so
the editor's `dist/` exists when it is published. After editing editor source,
rebuild it (`pnpm --filter @office-kit/docx-editor run build`) before running
the site or the example.

Tests need no build. The root `vitest.config.ts` aliases `@office-kit/docx` to
`src/index.ts`, and the editor's tests import their own package by relative
path (`./index.js`), never as `@office-kit/docx-editor`. So the CI `test` job,
which does not share a checkout with the `static` job, runs editor tests
without any `dist/`. Keep it that way: a test that imports
`@office-kit/docx-editor` would need a build step in that job.

`packages/editor/LICENSE` (like `packages/preview/LICENSE`) is a copy of the
root `LICENSE`, so it ships in the tarball. Keep the copies in sync.

## Verification status

**PASS (2026-10-03T19:00Z)** rows were run once locally on Node v26.6.0 / pnpm
10.25.0 (not a CI Node version; not a clean checkout), on uncommitted changes
over HEAD `15aec198`. See the run log below. **NOT RUN** means the check exists
but has not been run.

| Requirement                                                                                                                                                                                                                                                                         | Evidence                                                                                                | Status                                                                            |
| ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------- |
| Failed command restores doc, selection and redo                                                                                                                                                                                                                                     | `src/history.test.ts`                                                                                   | PASS (2026-10-03T19:00Z)                                                          |
| `runCommand` result type, all call sites compile                                                                                                                                                                                                                                    | `tsc --noEmit` (editor), `svelte-check` (site)                                                          | PASS (2026-10-03T19:00Z)                                                          |
| Replace All / paste / range delete undo as one step                                                                                                                                                                                                                                 | `history.test.ts`, `range-edit.test.ts`                                                                 | PASS (2026-10-03T19:00Z)                                                          |
| Invalid table size / unsafe link rejected atomically                                                                                                                                                                                                                                | `history.test.ts`                                                                                       | PASS (2026-10-03T19:00Z)                                                          |
| Shift+Enter at caret; Enter / typing replace a range                                                                                                                                                                                                                                | `range-edit.test.ts`                                                                                    | PASS (2026-10-03T19:00Z)                                                          |
| Section break kept once across split / range delete                                                                                                                                                                                                                                 | `history.test.ts`, `range-edit.test.ts`                                                                 | PASS (2026-10-03T19:00Z)                                                          |
| `<w:body>/<w:p>/<w:r>/<w:tbl>/<w:tr>/<w:tc>` attrs round-trip                                                                                                                                                                                                                       | `src/internal/wordprocessingml/attribute-preservation.test.ts` (repo root)                              | PASS (2026-10-03T19:00Z)                                                          |
| Attrs survive edit → undo/redo → save → reopen; split never copies them                                                                                                                                                                                                             | `range-edit.test.ts`                                                                                    | PASS (2026-10-03T19:00Z)                                                          |
| Same edit script run twice is stable                                                                                                                                                                                                                                                | `history.test.ts`, `examples/edit-and-save.mjs`                                                         | PASS (2026-10-03T19:00Z)                                                          |
| Links / field results render read-only; cell paragraphs anchored                                                                                                                                                                                                                    | `range-edit.test.ts`                                                                                    | PASS (2026-10-03T19:00Z)                                                          |
| Enter / Backspace / range delete within a cell; cross-cell delete refused atomically                                                                                                                                                                                                | `range-edit.test.ts`                                                                                    | PASS (2026-10-03T19:00Z)                                                          |
| Range-replacing IME input is one undo step and restores the selection (model only)                                                                                                                                                                                                  | `range-edit.test.ts` (model contract)                                                                   | PASS (2026-10-03T19:00Z)                                                          |
| DOM caret on a link / the paragraph element / cell paragraph maps correctly                                                                                                                                                                                                         | `src/dom-selection.test.ts` (happy-dom)                                                                 | PASS (2026-10-03T19:00Z)                                                          |
| Nested edit refused without corrupting history                                                                                                                                                                                                                                      | `history.test.ts`                                                                                       | PASS (2026-10-03T19:00Z)                                                          |
| Header / footer `<w:p>/<w:r>` attrs survive text replace → save → reopen                                                                                                                                                                                                            | `src/api/header-attributes.test.ts` (repo root)                                                         | PASS (2026-10-03T19:00Z)                                                          |
| Existing parser/writer and paragraph-split contracts unaffected: `src/internal/wordprocessingml/round-trip.test.ts` (9), `src/api/paragraph-edit.test.ts` (7), both unmodified                                                                                                      | those two files only (AST-equivalence, not byte-golden)                                                 | PASS (2026-10-03T19:38Z)                                                          |
| IME: Enter/Backspace during composition do not split or merge                                                                                                                                                                                                                       | browser (Japanese IME, Chrome + Safari), steps below                                                    | NOT RUN                                                                           |
| IME over a range: one undo step, DOM re-rendered from the model (CDP-simulated composition, Chromium only)                                                                                                                                                                          | browser, steps below                                                                                    | PASS (2026-10-03T19:21Z, Chromium smoke)                                          |
| Caret restore after undo / cell Enter / range delete (paste not exercised)                                                                                                                                                                                                          | browser                                                                                                 | PASS (2026-10-03T19:21Z, Chromium smoke)                                          |
| Browser: load fixture → typing over a selection → Undo/Redo → Download → reopen keeps text, unknown XML and body/header attrs                                                                                                                                                       | Chromium smoke (log below)                                                                              | PASS (2026-10-03T19:21Z, Chromium smoke)                                          |
| Paste in the browser; Enter/Backspace during composition; physical Japanese IME; Safari                                                                                                                                                                                             | browser                                                                                                 | NOT RUN                                                                           |
| Typing burst ends on a caret key that really moves the caret: real ArrowLeft (offset 8 → 7), the next typing is its own undo step. Cmd+Arrow not exercised                                                                                                                          | Chromium UI check (below)                                                                               | PASS (2026-10-03T19:36Z, Chromium)                                                |
| Ribbon UI: SVG align icons; none / value / inherited / mixed states for Font, size, colors, Styles; table cell shows its own paragraph style; ribbon refreshes after a command, Undo and Redo; inspecting writes nothing to the document                                            | Chromium UI check (below), site `svelte-check`                                                          | PASS (2026-10-03T19:36Z, Chromium); hydration console warnings remain (see below) |
| Highlight writes only ST_HighlightColor (16 + `none`): editor command and core `appendTextRun` / `setRunFormat` / `setParagraphText` / `setTableCellText` reject hex before any change; undo/redo kept; save → reopen; script run twice gives identical XML; ribbon is a named list | `packages/editor/src/highlight.test.ts`, `src/api/run-highlight.test.ts`, dist script, Chromium (below) | PASS (2026-10-03T19:49Z)                                                          |
| Color (ST_HexColor: 6 hex / `auto`) and size (ST_HpsMeasure: non-negative integer, 0 valid) rejected before any change by core typed writers and editor commands; valid values and size bounds survive save → reopen; raw foreign values kept                                       | `src/api/run-format-validation.test.ts`, `packages/editor/src/format-validation.test.ts`                | PASS (2026-10-03T20:29Z)                                                          |
| Canvas does not let a crafted color / font break out of the run `style` attribute                                                                                                                                                                                                   | `packages/editor/src/format-validation.test.ts`                                                         | PASS (2026-10-03T20:29Z)                                                          |
| Examples select the first editable body paragraph and fail loudly (no vacuous success); main part via `Docx.partName`                                                                                                                                                               | `examples/edit-and-save.mjs`, `examples/highlight-and-save.mjs`                                         | PASS (2026-10-03T20:30Z, table-first input, each run twice)                       |
| Find / Replace All (and the body part of the `Everywhere` variants) include table-cell paragraphs; one undo step, redo, save → reopen                                                                                                                                               | `src/api/table-text-search.test.ts`, `history.test.ts`, examples                                        | PASS (2026-10-03T20:48Z)                                                          |
| `getRunFormat` converts unit font sizes (`12pt` → 24) exactly; non-whole / invalid sizes stay unset; the XML keeps them                                                                                                                                                             | `src/api/run-format-size-units.test.ts`                                                                 | PASS (2026-10-03T20:48Z)                                                          |
| `runsInRange` returns exactly the covered runs; ribbon shows "Mixed" for a different middle run, and choosing a value unifies the range (Undo / Redo)                                                                                                                               | `packages/editor/src/range-runs.test.ts`, Chromium (below)                                              | PASS (2026-10-03T20:50Z)                                                          |
| Saved file opens cleanly in Word, and Word shows the same ribbon state and layout as `/editor` for it                                                                                                                                                                               | Word for Mac 16, `docs/qa/editor-20261005/`                                                             | PASS (2026-10-05T09:03Z, Word for Mac 16, light appearance)                       |
| Local serial `build:packages` (core → preview → editor), then root/editor/preview `tsc` and site `svelte-check`                                                                                                                                                                     | commands in the run log                                                                                 | PASS (2026-10-03T19:00Z)                                                          |
| Clean checkout / CI `static` + deploy-site / publish produce and use editor `dist/`                                                                                                                                                                                                 | CI                                                                                                      | NOT RUN                                                                           |
| Packaged example runs twice on its own output (0 replacements on re-run)                                                                                                                                                                                                            | commands above                                                                                          | PASS (2026-10-03T19:00Z)                                                          |
| Full `vitest` run (all existing suites, incl. any golden XML comparison)                                                                                                                                                                                                            | `pnpm test`                                                                                             | PASS (2026-10-05, 94 files / 791 tests, 8 skipped)                                |
| Repository gates `pnpm format:check`, `pnpm lint`                                                                                                                                                                                                                                   | run log (UI fix)                                                                                        | PASS (2026-10-03T19:35Z)                                                          |
| Repository gate `pnpm check:api-page` (34 pre-existing missing entries + `HIGHLIGHT_COLORS` added to `site/src/lib/api-groups.ts`)                                                                                                                                                  | run log (Highlight fix check)                                                                           | PASS (2026-10-03T19:52Z)                                                          |
| Repository gate `pnpm check:tree-shake`                                                                                                                                                                                                                                             | `pnpm check:tree-shake`                                                                                 | PASS (2026-10-05)                                                                 |
| Tarball contents (`dist/`, `README.md`, `LICENSE`)                                                                                                                                                                                                                                  | `pnpm --filter @office-kit/docx-editor pack --dry-run`                                                  | NOT RUN                                                                           |

### Run log (2026-10-03, local)

All exit codes are 0. Start times are UTC. Tools: Vitest 2.1.9, tsdown, tsc,
svelte-check.

```text
19:00:05Z pnpm exec vitest run packages/editor/src/history.test.ts \
            packages/editor/src/range-edit.test.ts packages/editor/src/dom-selection.test.ts \
            src/internal/wordprocessingml/attribute-preservation.test.ts \
            src/api/header-attributes.test.ts                                   3s
          Test Files 5 passed (5) / Tests 36 passed (36)
          (history 10, range-edit 16, dom-selection 5, attribute-preservation 4, header-attributes 1)
          exit 0 inferred from this output: zsh does not set PIPESTATUS, so it was not captured
19:00:16Z pnpm build:packages                                                    4s
          core, preview, editor: "Build complete"; editor dist imports @office-kit/docx (not inlined)
19:00:31Z pnpm typecheck                                    (root tsc --noEmit)  1s
19:00:32Z pnpm --filter @office-kit/docx-editor run typecheck                    2s
19:00:34Z pnpm --filter @office-kit/docx-preview run typecheck                   1s
19:00:45Z pnpm --filter word-kit-site check                                      4s
          COMPLETED 328 FILES 0 ERRORS 0 WARNINGS 0 FILES_WITH_PROBLEMS
19:01:24Z fixture: body "DRAFT"×2, unknown <x:keep xmlns:x="urn:wk-test">, w:rsidR / w:rsidRPr
          on body and header <w:p>/<w:r>
          node packages/editor/examples/edit-and-save.mjs in.docx out1.docx      <1s
            rejected as expected: Table rows must be an integer from 1 to 1000, got 0.
            replaced 2 occurrence(s); undo depth >0
          node packages/editor/examples/edit-and-save.mjs out1.docx out2.docx    <1s
            rejected as expected: Table rows must be an integer from 1 to 1000, got 0.
            replaced 0 occurrence(s); undo depth >0
          verifier (public dist API), out1 + out2: validate clean, no DRAFT, FINAL×2, body and header
            <w:p>/<w:r> attrs kept, unknown element kept once, Heading1 defined once;
            undo/redo of Replace All (2), invalid input rejected without redo,
            attrs + unknown element kept after undo/redo and save — ALL PASS
```

The example replaces text in the body only, so the header attributes above
survived a save in which the header was not rewritten. Header attributes
surviving a rewrite are covered by `header-attributes.test.ts`.

### Browser smoke (2026-10-03, local Chromium)

From 19:19:51Z to 19:23:01Z UTC, plus a recheck from 19:24:08Z to 19:24:58Z,
Node v26.6.0. A dev server was started for the
smoke only (`pnpm --filter word-kit-site dev --port 5317 --strictPort --host
127.0.0.1`, Vite 8.0.13) and stopped afterwards (again for the recheck). The browser was the existing
Playwright MCP (isolated Chromium, UA Chrome/154); nothing was installed. The
fixture had DRAFT text, an unknown-namespace element, `w:rsidR` / `w:rsidRPr`
on body and header `<w:p>`/`<w:r>`, and a table cell `abc`. Results (all pass after the step 2 fix):

1. Load through "Open…". Blocks, canvas text and word counts are correct, and
   Undo is disabled.
2. Select "DRAFT" and type `FINAL` with real key events. Model and DOM agree,
   and the caret is at 5. **First run (19:21Z): this took two undos**: the
   first character's range replacement and the rest of the word were separate
   steps. That contradicted the one-step typing burst documented above, so it
   was a bug. Fixed in `EditorCanvas.svelte`: the burst stays open after a
   range-replacing keystroke. **Recheck (19:24Z, after the fix):** one Ctrl+Z
   returns to `DRAFT` with the range selection 0..5 restored, and Ctrl+Y
   re-applies `FINAL`. After pressing `End`, typing ` x` is its own undo step.
   On macOS Chrome, `End` does not move the caret, so ` x` landed at offset 5.
   What this verified is that a caret key ends the burst, not an actual caret
   relocation. Site `svelte-check` was re-run afterwards: 328 files, 0 errors.
3. In the cell, Enter at `a|bc` gives cell paragraphs `["a","bc"]`. Nothing is
   added after the table, and the caret is in paragraph 1. Backspace gives
   `["abc"]`, caret at 1.
4. Simulated IME over "report": CDP `Input.imeSetComposition` → `Input.insertText`
   (`日本`). The model stays unchanged during the composition. The commit is one
   step, with model = DOM = `FINAL 日本`. Undo restores `report` and the range
   selection.
5. A range from paragraph 0 offset 5 to paragraph 1 offset 7, then Backspace,
   gives `["FINAL DRAFT","tail"]`, caret at 0:5. Undo restores both paragraphs
   and the selection.
6. Download, then reopen with the public API: `validate` is clean, and the
   expected body and cell text, unknown element, and body and header attributes
   are all present. This download was taken before the burst fix, which only
   changes how undo steps are grouped, not the saved content.

Console: 0 errors, 2 warnings (a color input received `""`), no page errors.
Cosmetic issues found here were fixed in source afterwards and checked in
"UI fix check" below. The align icons are now inline SVG instead of ⯇/⯈
glyphs. Color inputs always get a valid `#rrggbb`, with the selection's
`auto` / `none` / inherited / mixed color shown in the label. The empty
Font / size / Styles controls now show explicit "No selection" / "Mixed" /
"Default (inherited)" states and are disabled without a selection. Composition was synthetic CDP events in
Chromium: no physical Japanese IME and no Safari were used. Paste was not
exercised.

### UI fix check (2026-10-03, local)

Serial, Node v26.6.0, start times UTC. Exit codes were captured directly.

```text
19:34:17Z pnpm --filter word-kit-site check                       exit 0  4s  328 files, 0 errors, 0 warnings
19:34:26Z pnpm format:check                                       exit 1  1s  README.md, char-offset.ts,
          commands/structure.ts (all three are new / changed in this work)
19:34:51Z pnpm exec oxfmt <those three files only>                exit 0      line wraps and *x* → _x_ only
19:34:52Z pnpm format:check                                       exit 0  1s  239 files
19:34:59Z pnpm lint                                               exit 1  1s  unicorn(consistent-function-scoping)
          in range-edit.test.ts (new file); the `at` helper moved to module scope
19:35:14Z pnpm lint                                               exit 0  1s  0 warnings, 0 errors
19:35:15Z pnpm format:check                                       exit 0  1s
19:35:23Z pnpm exec vitest run (history, range-edit, dom-selection)  exit 0  2s  31 passed
19:35:35Z pnpm --filter @office-kit/docx-editor run build         exit 0  2s
19:35:37Z pnpm --filter word-kit-site run build                   exit 0  4s  vite build + pagefind
          (site/build and .svelte-kit are git-ignored)
19:38:40Z pnpm exec vitest run src/internal/wordprocessingml/round-trip.test.ts  exit 0  1s  9 passed
19:38:41Z pnpm exec vitest run src/api/paragraph-edit.test.ts                    exit 0  2s  7 passed
```

The two existing files compare re-parsed ASTs, not bytes against golden XML.
That elements with no attributes serialize exactly as before rests on the writer
code (`attrs ?? []`), not on a byte-level test. The full suite was not run.

Browser: one isolated Playwright MCP Chromium; dev server on 5317 only, from
19:35:53Z to 19:38:32Z (stopped, port closed, other listeners untouched). The fixture was built with the
public API: paragraph 0 = Arial 14pt red run + run without direct formatting;
paragraph 1 = Comic Sans MS, color `auto`, highlight yellow, Heading1; a table
whose cell (0,0) is Heading2 and cell (0,1) has no style.

1. Align left / Center / Align right / Justify render as SVG (`svg path`), with no glyph text.
2. Nothing selected: Font / size / Styles disabled and showing "No selection". The color labels show "No selection".
   Run with direct formatting: Arial / 14 / `#FF0000`. Run without direct formatting:
   "Default (inherited)". A range over both runs: "Mixed". Font outside the list:
   shows "Comic Sans MS"; color `auto` → "Automatic"; highlight "yellow";
   Styles "Heading 1".
3. Cell (0,0) shows Heading 2 and cell (0,1) shows Normal, so each cell shows the
   style of its own paragraph. Afterwards the document body was byte-identical
   (JSON) and Undo stayed disabled: inspecting wrote no fallback into the document.
4. Font → Georgia shows Georgia; toolbar Undo → "Default (inherited)"; Redo → Georgia.
5. Type `ab` at the end of " plain", press real ArrowLeft (DOM offset 8 → 7),
   type `X` → `Direct plainaXb`. Ctrl+Z → `Direct plainab`; Ctrl+Z → `Direct plain`.
6. Select "Mixed" and type `FINAL` → `FINAL para`. One Ctrl+Z → `Mixed para`, selection
   restored to 0..5, nothing left to undo; Ctrl+Y → `FINAL para`.

Screenshots: `ui-01-none.png` … `ui-05-final.png` in the session's temp
evidence folder (not committed). Console: 0 errors, no page errors, **2
warnings** "The specified value "" does not conform … valid CSS color". They
appear at page load, before any interaction. A stack trace shows them coming
from Svelte 5.55.7 hydration (`remove_input_defaults` → `removeAttribute('value')`),
one per color input, while the app's values are always valid hex (`#111111` /
`#ffff00`, also in the SSR HTML). The Svelte compiler emits that call for any
dynamic `value` attribute (`RegularElement.js`), not only `bind:value`, so a
`value={…}` + handler rewrite would not remove them. They are left as a known
limitation and not hidden. The save path was not changed by the UI fix, so the
earlier download / reopen evidence (unknown XML and body / header attributes)
still applies and was not re-run.

### Highlight fix check (2026-10-03, local)

`setHighlightCommand` used to write any hex into `<w:highlight w:val>`, and the
core `RunFormatting.highlight` doc comment called it "Hex RGB". Both are wrong
per ECMA-376: the local schema (`references/python-docx/ref/xsd/wml.xsd`,
`ST_HighlightColor`) and the Open XML SDK `HighlightColorValues` docs list the
same 16 colors plus `none`. The list now lives once, as `HIGHLIGHT_COLORS` in
`@office-kit/docx`. Its typed writers validate it before mutating, and the
editor checks it again before splitting runs. The read type stays `string`, so
`getRunFormat` still returns foreign values unchanged.

```text
19:44:03Z vitest packages/editor/src/highlight.test.ts (before the fix)  exit 1  8 failed / 19 passed
19:47:48Z pnpm build:packages                                  exit 0  4s  3 × "Build complete"
19:47:52Z pnpm typecheck / editor typecheck / preview typecheck   exit 0  1s each
19:48:02Z vitest: run-highlight, run-format, run-format-read, table-helpers,
          paragraph-helpers, editor highlight, commands, text-selection,
          integration, history, range-edit                      exit 0  2s  11 files, 146 tests
19:48:11Z pnpm --filter word-kit-site check                     exit 1  narrowing lost in a template closure;
          fixed with an `isListedHighlight` helper
19:48:16Z pnpm lint                                             exit 1  4 × no-array-sort in the new tests → toSorted()
19:48:18Z pnpm check:api-page                                   exit 1  34 exports missing from api-groups.ts;
          all 34 already exist at HEAD (pre-existing drift, not fixed here);
          HIGHLIGHT_COLORS was added to the page
19:48:40Z pnpm --filter word-kit-site check                     exit 0  328 files, 0 errors
19:48:43Z pnpm format:check                                     exit 0
19:48:44Z pnpm lint                                             exit 0  0 / 0
19:48:45Z vitest run-highlight + editor highlight               exit 0  74 passed
19:49:24Z node two-runs.mjs (built dist) × 2                    exit 0  "#FFFF00" and core "FFFF00" → RangeError;
          <w:highlight w:val="darkYellow"/> once; validate 0; document.xml identical across runs
          (scratch script importing dist/index.mjs and packages/editor/dist/index.mjs by
          absolute path, not the package entry points; superseded by the example below)
19:52:46Z pnpm check:api-page                                   exit 0  155 core exports, 156 entries
19:53:20Z node packages/editor/examples/highlight-and-save.mjs in.docx out1.docx    exit 0
          "#FFFF00" → RangeError; w:highlight values: darkYellow; all named; validate 0
19:53:21Z node packages/editor/examples/highlight-and-save.mjs out1.docx out2.docx  exit 0
          same output; word/document.xml byte-identical between out1 and out2
19:53:28Z pnpm format:check → exit 1 (line wrap in the new example) → oxfmt that file → 19:53:39Z exit 0
19:53:30Z pnpm lint exit 0; 19:53:31Z pnpm --filter word-kit-site check exit 0 (328 files, 0 errors)
```

`examples/highlight-and-save.mjs` is separate from the initial
`edit-and-save.mjs`. It imports the bare specifiers `@office-kit/docx` and
`@office-kit/docx-editor`, which resolve through the package `exports` (via
the workspace link and the editor's self-reference) to `dist/index.mjs` and
`packages/editor/dist/index.mjs`, both built at 19:47:48Z after the last
source change. Nothing was installed. It reads the saved XML back through
the public `getRawPartRoot`.

The 34 `check:api-page` misses were exports already present at HEAD
`15aec198` (`src/api/index.ts` / `src/api/docx.ts`) but never listed in
`api-groups.ts`. The fix only adds list entries to existing groups:

- Inline & text: `setRunOnOff`, `setRunValProp`, `getRunProp`,
  `setParagraphOnOff`, `setParagraphValProp`, `getParagraphProp`,
  `splitParagraphAt`, `mergeParagraphIntoPrevious`, `runTextLength`,
  `isolateParagraphRunRange`.
- Styles & numbering: `setStyleOnOff`, `setStyleValProp`, `getStyleProp`,
  `setNumberingLevelOnOff`, `setNumberingLevelVal`, `getNumberingLevelProp`.
- Images: `imageDrawings`, `getImageInfo`, `setImageAltText`, `setImageSizeEmu`.
- Document properties: `getDocumentSetting`, `setDocumentSettingOnOff`,
  `setDocumentSettingVal`.
- Low-level part access: `xmlPartNames`, `getRawPartRoot`, `markRawPartDirty`,
  `childElementsOf`, `appendChildElement`, `makePropsElement`,
  `getElementAttr`, `setElementAttr`, `getElementProp`, `setElementOnOff`,
  `setElementValProp`.

`HIGHLIGHT_COLORS` got a one-line description on the API page.

Browser (one isolated Chromium, dev server on 5317 from 19:49:37Z to 19:50:08Z,
stopped, other listeners untouched):

- Without a selection the highlight list is disabled and shows "No selection".
- A run another tool wrote with `FFFF00` shows `FFFF00` as a disabled entry,
  with no background, and its value stays unchanged.
- Choosing `darkYellow` writes `darkYellow` and renders `rgb(128, 128, 0)`.
  Undo restores no highlight ("Default (inherited)") and Redo re-applies it.
- No color picker is left for highlight (one `input[type=color]`, font color).
- Console: 0 errors, no page errors, 1 warning: the Svelte hydration warning
  above, now only for the font color input.
- Screenshot: `hl-01-applied.png`.

### Color / size validation check (2026-10-03, local)

Run serially on Node v26.6.0 in the worktree, after the verification slot was
handed back. Start times are UTC. The earlier 146-test run (19:48Z) predates
this change and is superseded here. Vitest used its default workers: the
single-worker request arrived after these runs had finished.

```text
20:29:11Z vitest run-format-validation + editor format-validation  exit 0  2s  2 files, 122 tests
20:29:19Z vitest run-highlight, run-format, run-format-read, table-helpers,
          paragraph-helpers, editor highlight, commands, text-selection,
          integration, history, range-edit, dom-selection        exit 0  2s  12 files, 151 tests
20:29:27Z pnpm build:packages                                     exit 0  4s  3 × "Build complete"
20:29:31Z pnpm typecheck / editor typecheck / preview typecheck    exit 0  1s each
20:29:40Z pnpm --filter word-kit-site check                       exit 0  3s  328 files, 0 errors
20:29:43Z pnpm format:check                                       exit 1  1s  whitespace only in README.md and
          run-format-validation.test.ts (both changed here) → oxfmt those two → 20:29:57Z exit 0
20:29:44Z pnpm lint                                               exit 0  0 / 0
20:29:45Z pnpm check:api-page                                     exit 0  155 core exports
20:30:04Z vitest run-format-validation (after the reformat)        exit 0  1s  100 tests
20:30:20Z node examples/edit-and-save.mjs in(table-first) → es1   exit 0  replaced 2; table size rejected
20:30:20Z node examples/edit-and-save.mjs es1 → es2               exit 0  replaced 0; document.xml es1 = es2
20:30:20Z node examples/highlight-and-save.mjs es2 → hs1          exit 1  "setHighlightCommand was disabled":
          a false alarm in the example. runCommand returns undefined for a void command
          even when it ran, so the example now checks isEnabled before running.
20:30:59Z node examples/highlight-and-save.mjs es2 → hs1          exit 0  block 1 (after the table): darkYellow; validate 0
20:30:59Z node examples/highlight-and-save.mjs hs1 → hs2          exit 0  same; document.xml hs1 = hs2
```

The examples import the bare package names, which resolve through `exports` to
`dist/index.mjs` and `packages/editor/dist/index.mjs` (built at 20:29:27Z). The
table-first input was generated with the built core API.

Found, not fixed (outside this change): the input's `cell DRAFT` stays
`DRAFT`. Core `replaceText` (and so `replaceAllCommand`) walks top-level body
paragraphs only and skips table cells (`text-search.ts`, same at HEAD
`15aec198`). Fixed later; see "Table search, unit sizes, range runs".

### Table search, unit sizes, range runs (2026-10-03, local)

Three bugs against existing contracts, each with a test that fails first:

1. `findText` / `replaceText` promise "every paragraph" / "every occurrence"
   but skipped table cells. They now walk cell paragraphs too
   (`text-search.ts`), which the body part of the `Everywhere` variants and
   `replaceAllCommand` use. A table nested in a cell stays raw XML and is not
   searched.
2. `runsInRange` is documented as "every run touched by an ordered selection"
   but returned every run of the end paragraphs, and the ribbon looked only at
   the two end runs, so `Arial [Times] Arial` showed "Arial". Choosing Arial
   then fired no change, so the middle run could not be fixed. `runsInRange`
   now uses the same covered-run computation as the formatting commands
   (`selection-runs.ts`), and the ribbon reads a range through it.
3. `getRunFormat` used `parseInt` on `w:sz`, so `"12pt"` (ST_HpsMeasure
   universal measure) read as 12 half-points. Units are now converted exactly
   with integer fractions (1pt = 2, 1pc = 1pi = 24, 1in = 144,
   1cm = 7200/127, 1mm = 720/127 half-points). A value that is not a whole
   number of half-points, or is invalid, leaves the typed size unset; the XML
   is not touched.

Clear formatting is not changed (its scope is undecided).

Run one command at a time on Node v26.6.0. Vitest used
`--maxWorkers=1 --minWorkers=1`. Start times are UTC.

```text
20:46:58Z vitest (original code) table-text-search, run-format-size-units, range-runs
          exit 1  2s  17 failed / 6 passed (23):
            table-text-search 3/3 failed (cell matches missing: 2 found / replaced, expected 5)
            run-format-size-units 12/16 failed (7 unit sizes misread, 4 invalid values given a
              size, the 3mm keep-in-XML case read a size); 4 passed ("24", "0", "abc", "")
            range-runs 2/4 failed (end-paragraph runs outside the range returned);
              2 passed (whole-paragraph range, collapsed caret: same before and after)
20:48:22Z same 3 files after the fixes                           exit 0  2s  23 passed
20:48:38Z related files: an invocation error (zsh did not split the file list,
          "No test files found", exit 1, no test ran); re-run at 20:48:53Z
20:48:53Z 34 related files (core find / replace / run-format / table / paragraph /
          round-trip / samples, all editor suites)                exit 0  9s  33 passed + 1 skipped file;
          420 passed, 2 skipped (fixture.test.ts and styles.test.ts skip by condition)
20:49:13Z pnpm build:packages                                     exit 0  4s  3 × "Build complete"
20:49:17Z pnpm typecheck / editor / preview typecheck              exit 0  1s each
20:49:26Z pnpm --filter word-kit-site check                       exit 0  3s  328 files, 0 errors
20:49:29Z pnpm format:check                                       exit 0  250 files
20:49:30Z pnpm lint                                               exit 0  0 / 0
20:49:31Z pnpm check:api-page                                     exit 0  155 core exports
20:49:38Z node examples/edit-and-save.mjs table-first → es1       exit 0  replaced 3 (the table cell is now included)
20:49:39Z node examples/edit-and-save.mjs es1 → es2               exit 0  replaced 0; document.xml es1 = es2
20:49:39Z node examples/highlight-and-save.mjs es2 → hs1, hs1 → hs2  exit 0 ×2  darkYellow; validate 0; hs1 = hs2
20:50:01Z vitest history.test.ts (+ cell Replace All undo / redo / reopen case)  exit 0  11 passed
20:50:13Z pnpm lint  0 warnings / 0 errors (exit not captured: piped)
20:50:19Z pnpm format:check                                       exit 0
```

Browser (one isolated Chromium, own dev server on 5317 from 20:50:36Z to
20:51:03Z, stopped, port closed, other listeners untouched): runs
`Aaaa`(Arial) `Bbbb`(Times New Roman) `Cccc`(Arial), range from inside the
first run to inside the last.

- Font shows "Mixed".
- Choosing Arial makes every covered run Arial, middle run included.
- Undo restores the three fonts and shows "Mixed"; Redo shows Arial.
- Console: 0 errors, no page errors, 1 warning (the known Svelte hydration
  warning).
- Screenshot: `b3-01-mixed.png`.

Commands ran one after another. Only two light steps were issued at the same
time: a source `sed` read alongside a source edit, and a test-file write
alongside the 20:46:58Z Vitest run.

### Browser check steps (NOT RUN with a physical IME / Safari)

Run in Chrome and in Safari, with a Japanese IME:

1. Type `にほん`, then press Enter to confirm the conversion. The paragraph must
   not split. Press Enter again: now it splits.
2. Select text across two paragraphs, type `にほんご` and confirm. The selected
   text is replaced. Ctrl/Cmd+Z restores the original text and the selection
   in one step. Ctrl/Cmd+Y re-applies it.
3. Repeat step 2 inside one table cell, then across two cells. Across cells the
   status bar shows the refusal, and the document and caret are unchanged once
   the IME closes.
4. In a table cell, Enter → Backspace → Delete. Paragraphs split and join only
   within that cell, and the caret stays at the join.
5. Click inside a hyperlink's text and press Enter. The split happens next to
   the link, not at the start of the paragraph.

### Known gaps

- Links and fields are read-only on the canvas: their text cannot be edited in
  place. Read-only inlines sitting exactly at a range-deletion boundary are
  kept; within a single paragraph, a range delete removes runs only.
- Deleting a range that spans several cells or leaves the table is refused
  (atomic error).
- Editor test files are excluded from `tsc` (`tsconfig.json`), and Vitest does
  not type-check, so type errors in tests are not caught.
- `createStyleResolver` treats toggle properties (`b`, `i`, `strike`) as plain
  overrides; ECMA-376 §17.7.3 XORs them when two style levels both set one.
  Table-style conditional formatting (first row, banding) and numbering-level
  run properties are not applied. Theme fonts resolve for Latin text only.
- The `/editor` ribbon leaves out what the library cannot drive: the Draw and
  Design tabs, Change Case, text effects, shading, borders, sort, theme colors
  in the Font Color menu, and the Styles Pane. The style gallery offers Normal
  and Heading 1–3. The canvas is one continuous page (no pagination), so the
  status bar has no "Page 1 of 1".
- Word for Mac draws 100 % zoom at 72 dpi; the canvas uses CSS points (96 dpi,
  like Word for Windows), so a page looks larger than in Word for Mac at the
  same zoom.
- The generic raw setters of `@office-kit/docx` (`setRunValProp`,
  `setElementValProp`, `setStyleValProp`) write any value for any element by
  design, so they can still author an invalid `w:highlight`. Only the typed
  `RunFormatting` writers validate it.
- An IME composition over a range that fails, for example one that crosses
  cells, discards the composed text. The canvas re-renders from the unchanged
  model.
