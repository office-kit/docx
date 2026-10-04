// Open a .docx, edit it through editor commands, and save it back — using only
// the published entry points. Safe to run repeatedly on its own output: every
// step is idempotent, so a second run reports 0 replacements and adds nothing.
//
//   node packages/editor/examples/edit-and-save.mjs in.docx out.docx
import { readFile, writeFile } from "node:fs/promises";
import { toUint8Array } from "@office-kit/docx";
import { commands, openEditor, runCommand } from "@office-kit/docx-editor";

const [input, output] = process.argv.slice(2);
if (!input || !output) {
  console.error("usage: edit-and-save.mjs <in.docx> <out.docx>");
  process.exit(2);
}

const model = openEditor(new Uint8Array(await readFile(input)));
// Put the caret in the first body paragraph (the body may start with a table).
const block = model.doc.document.body.blocks.findIndex((b) => b.kind === "paragraph");
if (block === -1) throw new Error(`${input}: the body has no paragraph`);
model.setSelection({ anchor: { block }, focus: { block } });

runCommand(model, commands.ensureHeadingStylesCommand, undefined);
const replaced = runCommand(model, commands.replaceAllCommand, {
  query: "DRAFT",
  replacement: "FINAL",
});

if (replaced === undefined) throw new Error("replaceAllCommand was disabled");

// A rejected command throws and leaves the document untouched (atomic).
let rejected = false;
try {
  runCommand(model, commands.insertTableCommand, { rows: 0, cols: 3 });
} catch (err) {
  if (!(err instanceof RangeError)) throw err;
  rejected = true;
  console.log(`rejected as expected: ${err.message}`);
}
if (!rejected) throw new Error("an invalid table size was not rejected");

await writeFile(output, toUint8Array(model.doc));
console.log(`replaced ${replaced} occurrence(s); undo depth ${model.canUndo() ? ">0" : "0"}`);
