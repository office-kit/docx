import {
  createDocx,
  mailMergeSettings,
  noteMarks,
  noteProperties,
  openDocx,
  paragraphs,
  parseRecipientCsv,
  toUint8Array,
  validate,
} from "@office-kit/docx";
import { describe, expect, it } from "vitest";
import { editorFor } from "../index.js";
import type { EditorModel } from "../model.js";
import { caretAt } from "../selection.js";
import {
  insertAddressBlockCommand,
  insertMergeFieldCommand,
  previewResultsCommand,
  selectRecipientsCommand,
  startMailMergeCommand,
} from "./mailings.js";
import {
  addFootnoteCommand,
  findNoteReference,
  insertCitationCommand,
  insertTocCommand,
  noteOptionsCommand,
  setSourcesCommand,
} from "./references.js";
import { runCommand } from "./types.js";

function editorWith(texts: string[], at = { block: 0, inline: 0, offset: 0 }): EditorModel {
  const model = editorFor(createDocx({ paragraphs: texts }));
  model.setSelection(caretAt(at));
  return model;
}

function reopen(model: EditorModel) {
  const doc = openDocx(toUint8Array(model.doc));
  expect(validate(doc)).toEqual([]);
  return doc;
}

describe("references commands", () => {
  it("inserts a footnote at the caret, applies options, and finds it", () => {
    const model = editorWith(["Hello world", "Second"], { block: 0, inline: 0, offset: 5 });
    expect(runCommand(model, addFootnoteCommand, { text: "A note." })).toBeGreaterThan(0);
    runCommand(model, noteOptionsCommand, {
      kind: "footnote",
      properties: { numberFormat: "lowerRoman" },
      scope: "document",
    });
    const doc = reopen(model);
    expect(noteMarks(doc, "footnote").map((m) => m.mark)).toEqual(["i"]);
    expect(noteProperties(doc, "footnote").numberFormat).toBe("lowerRoman");

    model.setSelection(caretAt({ block: 1 }));
    expect(findNoteReference(model, "footnote", "previous")?.block).toBe(0);
    expect(findNoteReference(model, "footnote", "next")).toBeUndefined();
  });

  it("inserts a table of contents before the caret block when the caret is at its start", () => {
    const model = editorWith(["Body"]);
    runCommand(model, insertTocCommand, { levels: { from: 1, to: 3 } });
    const blocks = model.doc.document.body.blocks;
    expect(blocks.length).toBeGreaterThan(1);
    expect(paragraphs(model.doc).at(-1)?.children.length).toBeGreaterThan(0);
    reopen(model);
  });

  it("stores a source and cites it with a computed result", () => {
    const model = editorWith(["Claim."], { block: 0, inline: 0, offset: 6 });
    runCommand(model, setSourcesCommand, {
      sources: [
        { tag: "Lov43", type: "Book", authors: [{ last: "Lovelace", first: "Ada" }], fields: { Year: "1843" } },
      ],
    });
    runCommand(model, insertCitationCommand, { tag: "Lov43" });
    const doc = reopen(model);
    const text = JSON.stringify(paragraphs(doc)[0]);
    expect(text).toContain("CITATION Lov43");
    expect(text).toContain("(Lovelace, 1843)");
  });
});

describe("mailings commands", () => {
  it("attaches a CSV list, inserts fields, and previews a record", () => {
    const model = editorWith(["Dear "], { block: 0, inline: 0, offset: 5 });
    const list = parseRecipientCsv("First Name,Last Name\nAda,Lovelace\n");
    runCommand(model, startMailMergeCommand, { type: "formLetters" });
    runCommand(model, selectRecipientsCommand, { path: "recipients.csv", list });
    runCommand(model, insertMergeFieldCommand, { name: "First Name" });
    runCommand(model, insertAddressBlockCommand, {});
    runCommand(model, previewResultsCommand, { list, index: 0 });
    const doc = reopen(model);
    expect(mailMergeSettings(doc)?.type).toBe("formLetters");
    expect(JSON.stringify(paragraphs(doc)[0])).toContain("Ada");
  });
});
