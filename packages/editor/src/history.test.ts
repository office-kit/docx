/**
 * Edit-transaction guarantees: a command either applies completely as one undo
 * step or leaves the document, selection and redo history exactly as they
 * were. Also covers the load → edit → save → reopen loop when the same edit
 * script runs twice. Uses only the package's public surface (`./index.js`).
 */

import {
  addTable,
  appendSectionBreak,
  createDocx,
  getTableCellText,
  listStyles,
  openDocx,
  paragraphs,
  paragraphText,
  tables,
  toUint8Array,
  validate,
} from "@office-kit/docx";
import { describe, expect, it } from "vitest";
import {
  type Command,
  commands,
  editorFor,
  type EditorModel,
  openEditor,
  runCommand,
} from "./index.js";
import { caretAt } from "./selection.js";

function editorWith(texts: string[]): EditorModel {
  const model = editorFor(createDocx({ paragraphs: texts }));
  model.setSelection({ anchor: { block: 0 }, focus: { block: 0 } });
  return model;
}

function texts(model: EditorModel): string[] {
  return paragraphs(model.doc).map(paragraphText);
}

const failingCommand: Command<void> = {
  id: "test.failing",
  group: "text",
  label: "Fails halfway",
  run(model) {
    model.doc.document.body.blocks.length = 0;
    throw new Error("boom");
  },
};

describe("atomic commands", () => {
  it("rolls a failed command back without polluting redo history", () => {
    const model = editorWith(["one"]);
    runCommand(model, commands.replaceAllCommand, { query: "one", replacement: "two" });
    model.undo();
    expect(texts(model)).toEqual(["one"]);
    expect(model.canRedo()).toBe(true);

    expect(() => runCommand(model, failingCommand, undefined)).toThrow("boom");

    // Document and selection restored; the pre-existing redo entry survives and
    // the failed half-edit is not reachable through redo.
    expect(texts(model)).toEqual(["one"]);
    expect(model.selection).toEqual({ anchor: { block: 0 }, focus: { block: 0 } });
    expect(model.canUndo()).toBe(false);
    model.redo();
    expect(texts(model)).toEqual(["two"]);
    expect(model.canRedo()).toBe(false);
  });

  it("rolls back even when the snapshot was evicted by the history limit", () => {
    const model = editorFor(createDocx({ paragraphs: ["keep"] }), { historyLimit: 0 });
    expect(() => runCommand(model, failingCommand, undefined)).toThrow("boom");
    expect(texts(model)).toEqual(["keep"]);
  });

  it("refuses a nested edit instead of corrupting the pending snapshot", () => {
    const model = editorWith(["x"]);
    const nesting: Command<void> = {
      id: "test.nesting",
      group: "text",
      label: "Runs a command inside a command",
      run(m) {
        runCommand(m, commands.replaceAllCommand, { query: "x", replacement: "y" });
      },
    };
    expect(() => runCommand(model, nesting, undefined)).toThrow(/another edit is open/);
    expect(texts(model)).toEqual(["x"]);
    expect(model.canUndo()).toBe(false);
  });

  it("rejects invalid table sizes without changing the document", () => {
    const model = editorWith(["a"]);
    for (const [rows, cols] of [
      [0, 2],
      [-1, 2],
      [2.5, 2],
      [Number.NaN, 2],
      [2, 64],
      [100_000, 1],
    ] as const) {
      expect(() => runCommand(model, commands.insertTableCommand, { rows, cols })).toThrow(
        RangeError,
      );
    }
    expect(model.doc.document.body.blocks).toHaveLength(1);
    expect(model.canUndo()).toBe(false);
  });

  it("rejects unsafe or malformed hyperlink targets", () => {
    const model = editorWith(["a"]);
    for (const url of ["javascript:alert(1)", "data:text/html,x", "https://", "not a url"]) {
      expect(() =>
        runCommand(model, commands.insertLinkCommand, { text: "x", target: { url } }),
      ).toThrow();
    }
    expect(model.canUndo()).toBe(false);
    runCommand(model, commands.insertLinkCommand, {
      text: "x",
      target: { url: "https://example.com" },
    });
    expect(model.canUndo()).toBe(true);
  });
});

describe("replaceAllCommand", () => {
  it("is one undoable step and reports the count", () => {
    const model = editorWith(["cat cat", "cat"]);
    const n = runCommand(model, commands.replaceAllCommand, { query: "cat", replacement: "dog" });
    expect(n).toBe(3);
    expect(texts(model)).toEqual(["dog dog", "dog"]);
    model.undo();
    expect(texts(model)).toEqual(["cat cat", "cat"]);
  });

  it("rejects an empty query atomically", () => {
    const model = editorWith(["x"]);
    expect(() =>
      runCommand(model, commands.replaceAllCommand, { query: "", replacement: "y" }),
    ).toThrow();
    expect(model.canUndo()).toBe(false);
  });

  it("keeps a paragraph after a table that ends the document", () => {
    const doc = createDocx({ paragraphs: ["intro"] });
    addTable(doc, [["cell"]]);
    const model = editorFor(doc);
    expect(model.doc.document.body.blocks.map((b) => b.kind)).toEqual([
      "paragraph",
      "table",
      "paragraph",
    ]);
    // Inserting a table at the end leaves the caret's paragraph after it, too.
    model.setSelection(caretAt({ block: 2, inline: 0, offset: 0 }));
    runCommand(model, commands.insertTableCommand, { rows: 1, cols: 1 });
    expect(model.doc.document.body.blocks.at(-1)?.kind).toBe("paragraph");
  });

  it("replaces inside table cells in the same undo step, and redo re-applies it", () => {
    const doc = createDocx({ paragraphs: ["cat"] });
    addTable(doc, [["cat cell"]]);
    // The editor adds the paragraph Word keeps after a table that ends the body.
    const model = editorFor(doc);
    expect(
      runCommand(model, commands.replaceAllCommand, { query: "cat", replacement: "dog" }),
    ).toBe(2);
    const cell = (): string => getTableCellText(tables(model.doc)[0]!, 0, 0);
    expect(cell()).toBe("dog cell");
    model.undo();
    expect([texts(model), cell()]).toEqual([["cat", ""], "cat cell"]);
    expect(model.canUndo()).toBe(false);
    model.redo();
    expect([texts(model), cell()]).toEqual([["dog", ""], "dog cell"]);
    const reopened = openDocx(toUint8Array(model.doc));
    expect(getTableCellText(tables(reopened)[0]!, 0, 0)).toBe("dog cell");
  });
});

describe("insertTextCommand", () => {
  it("inserts multi-line text at the caret as a single undo step", () => {
    const model = editorWith(["Hello world"]);
    model.setSelection({
      anchor: { block: 0, inline: 0, offset: 5 },
      focus: { block: 0, inline: 0, offset: 5 },
    });
    runCommand(model, commands.insertTextCommand, { text: " big\r\nnew\nlines" });
    expect(texts(model)).toEqual(["Hello big", "new", "lines world"]);
    expect(model.selection?.focus).toEqual({ block: 2, inline: 0, offset: 5 });
    model.undo();
    expect(texts(model)).toEqual(["Hello world"]);
  });
});

describe("splitParagraphAt with a section break", () => {
  it("moves the paragraph-level sectPr to the second half instead of duplicating it", () => {
    const doc = createDocx({ paragraphs: [] });
    appendSectionBreak(doc, "nextPage");
    const model = editorFor(doc);
    model.setSelection({
      anchor: { block: 0, inline: 0, offset: 0 },
      focus: { block: 0, inline: 0, offset: 0 },
    });
    runCommand(model, commands.splitParagraphCommand, undefined);
    const reopened = openDocx(toUint8Array(model.doc));
    const withSectPr = paragraphs(reopened).filter((p) =>
      p.pPr?.children.some((c) => c.kind === "element" && c.name.local === "sectPr"),
    );
    expect(withSectPr).toHaveLength(1);
    expect(paragraphs(reopened).indexOf(withSectPr[0]!)).toBe(1);
  });
});

describe("load → edit → save → reopen", () => {
  /** A user-style edit script: idempotent, so running it twice is a no-op the second time. */
  function editScript(bytes: Uint8Array): { bytes: Uint8Array; replaced: number } {
    const model = openEditor(bytes);
    model.setSelection({ anchor: { block: 0 }, focus: { block: 0 } });
    runCommand(model, commands.ensureHeadingStylesCommand, undefined);
    const replaced =
      runCommand(model, commands.replaceAllCommand, { query: "DRAFT", replacement: "FINAL" }) ?? 0;
    return { bytes: toUint8Array(model.doc), replaced };
  }

  it("is stable when the same script runs twice", () => {
    const source = toUint8Array(createDocx({ paragraphs: ["DRAFT report", "status: DRAFT"] }));
    const first = editScript(source);
    expect(first.replaced).toBe(2);
    const second = editScript(first.bytes);
    expect(second.replaced).toBe(0);

    const a = openDocx(first.bytes);
    const b = openDocx(second.bytes);
    expect(validate(b)).toHaveLength(0);
    expect(paragraphs(b).map(paragraphText)).toEqual(paragraphs(a).map(paragraphText));
    expect(paragraphs(b).map(paragraphText)).toEqual(["FINAL report", "status: FINAL"]);
    // ensureHeadingStyles must not duplicate style definitions on re-run.
    expect(listStyles(b).filter((st) => st.styleId === "Heading1")).toHaveLength(1);
  });
});
