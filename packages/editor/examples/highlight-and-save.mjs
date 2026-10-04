// Highlight the first word of a .docx through the editor and save it, using
// only the published entry points. `<w:highlight>` takes a named color
// (`HIGHLIGHT_COLORS`), never RGB: a hex is rejected without changing anything.
// Safe to run repeatedly on its own output: the saved XML does not change.
//
//   node packages/editor/examples/highlight-and-save.mjs in.docx out.docx
import { readFile, writeFile } from "node:fs/promises";
import {
  getRawPartRoot,
  HIGHLIGHT_COLORS,
  openDocx,
  runTextLength,
  toUint8Array,
  validate,
} from "@office-kit/docx";
import { commands, openEditor, runCommand } from "@office-kit/docx-editor";

const [input, output] = process.argv.slice(2);
if (!input || !output) {
  console.error("usage: highlight-and-save.mjs <in.docx> <out.docx>");
  process.exit(2);
}

const model = openEditor(new Uint8Array(await readFile(input)));

// The first body paragraph whose first run has text, wherever it is (the body
// may start with a table). Positions count runs only, so that run is inline 0.
const blocks = model.doc.document.body.blocks;
const block = blocks.findIndex((b) => {
  const run = b.kind === "paragraph" ? b.children.find((c) => c.kind === "run") : undefined;
  return run !== undefined && runTextLength(run) > 0;
});
if (block === -1) throw new Error(`${input}: no body paragraph starts with a text run`);
const firstRun = blocks[block].children.find((c) => c.kind === "run");
model.setSelection({
  anchor: { block, inline: 0, offset: 0 },
  focus: { block, inline: 0, offset: runTextLength(firstRun) },
});

// A void command's runCommand result is undefined whether or not it ran, so
// check that it is enabled for this selection first.
if (!commands.setHighlightCommand.isEnabled?.(model)) {
  throw new Error("setHighlightCommand is disabled for the selected run");
}

let rejected = false;
try {
  runCommand(model, commands.setHighlightCommand, { color: "#FFFF00" });
} catch (err) {
  if (!(err instanceof RangeError)) throw err;
  rejected = true;
}
if (!rejected) throw new Error("a hex highlight was not rejected");
runCommand(model, commands.setHighlightCommand, { color: "darkYellow" });

const bytes = toUint8Array(model.doc);
await writeFile(output, bytes);

// Read the saved main document part back (`partName` comes from the package's
// officeDocument relationship) and list every w:highlight value in it.
const saved = openDocx(bytes);
const body = getRawPartRoot(saved, saved.partName);
if (!body) throw new Error(`main document part ${saved.partName} is missing`);
const values = [];
const walk = (el) => {
  if (el.name.local === "highlight") {
    values.push(el.attrs.find((a) => a.name.local === "val")?.value);
  }
  for (const c of el.children) if (c.kind === "element") walk(c);
};
walk(body);
if (values.length === 0) throw new Error("no w:highlight was saved");
const invalid = values.filter((v) => !HIGHLIGHT_COLORS.includes(v));
if (invalid.length > 0) throw new Error(`invalid w:highlight values saved: ${invalid.join(", ")}`);
console.log(
  `block ${block}: w:highlight values: ${values.join(", ")}; validate issues: ${validate(saved).length}`,
);
