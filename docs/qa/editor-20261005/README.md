# Editor vs. Microsoft Word (2026-10-05)

Each image shows Microsoft Word for Mac 16 (light appearance) on top and the
`/editor` route below, for the same document. Word's window was captured on
2026-10-05 around 09:03 UTC. The editor was captured from a local dev server
(`pnpm --filter word-kit-site dev`) in headless Chromium at 1512 × 949, the same
size as Word's window, with the site header cropped off.

| File                 | What it shows                                                                                                                                                                                                                                       |
| -------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `01-initial.png`     | The editor's sample document right after loading, and the same document saved by the editor and opened in Word. Both carets sit at the start of the heading: Calibri, 20, Bold and Align Left pressed, Heading 1 selected in the gallery, 50 words. |
| `02-after-edits.png` | After these edits in the editor: Bold on the heading (turns the style's bold off), Bright Green highlight and line spacing 1.5 and Increase Indent on the body paragraph. The saved file opened in Word shows the same result.                      |

The page in Word for Mac looks smaller: Word for Mac draws 100 % at 72 dpi, the
browser at 96 dpi (as Word for Windows does).
