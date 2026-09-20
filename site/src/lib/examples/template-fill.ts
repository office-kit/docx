// Open an existing .docx as a template, run cross-part placeholder
// substitution, write it back. The interesting bit: replaceTextEverywhere
// also walks headers, footers, footnotes, endnotes and comments — not just
// the body — so {{name}} in a header swaps out too.

import { readFile, writeFile } from "node:fs/promises";
import { openDocx, replaceTextEverywhere, toUint8Array } from "@office-kit/docx";

const values: Record<string, string> = { name: "Ada Lovelace" };

const doc = openDocx(await readFile("letter-template.docx"));

const replaced = replaceTextEverywhere(
  doc,
  /\{\{(\w+)\}\}/g,
  (m) => values[m.captures[0] ?? ""] ?? "",
);
console.log(`Filled ${replaced} placeholders.`);

await writeFile("letter.docx", toUint8Array(doc));
