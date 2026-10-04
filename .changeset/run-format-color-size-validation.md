---
"@office-kit/docx": minor
"@office-kit/docx-editor": minor
---

fix: run color and font size are written only as schema-valid values

`setRunFormat`, `appendTextRun`, `setParagraphText` and `setTableCellText` used
to write any `color` string and any `fontSizeHalfPoints` number into
`<w:color w:val>` / `<w:sz w:val>`. Values like `"red"`, `"#FF0000"`, `-1`,
`23.5` or `NaN` produced files that do not conform to the schema. These writers
now throw a `RangeError`, before changing anything, unless:

- `color` is six hex digits without `#` (either case) or `"auto"` (ST_HexColor);
- `fontSizeHalfPoints` is a non-negative safe integer (ST_HpsMeasure). `0` stays
  valid; there is no application-specific upper bound.

Untyped JavaScript callers get the same check on the runtime type: a `color`
that is not a string (e.g. the number `123456`, `null`) or a size that is not a
number (e.g. `"24"`) is rejected rather than stringified.

One case that used to pass silently now throws: `appendTextRun(p, text, { color: "" })`
ignored the empty color, and now rejects it like `setRunFormat` does.
`getRunFormat` still returns whatever a file holds, and untouched values are
saved unchanged.

In the editor, `setColorCommand` (a leading `#` is still dropped) and
`setFontSizeCommand` (points, rounded to half-points) reject invalid values
before the document changes, with undo and redo history kept.

The editor canvas also no longer lets a crafted file break out of a run's
`style` attribute: the attribute is HTML-escaped, only six-digit hex colors
become CSS, and quotes, backslashes and line breaks are dropped from font names
when rendering. The document itself is not changed.
