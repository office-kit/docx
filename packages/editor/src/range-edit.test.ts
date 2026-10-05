/**
 * Range editing (delete / replace a selection, Shift+Enter at the caret),
 * read-only rendering of unmodelled inlines, and preservation of element
 * attributes through edit → save → reopen. Public surface only.
 */

import {
  addHyperlink,
  appendField,
  appendSectionBreak,
  createDocx,
  openDocx,
  paragraphs,
  paragraphText,
  toUint8Array,
  validate,
  type WmlParagraph,
  type XmlAttr,
} from "@office-kit/docx";
import { describe, expect, it } from "vitest";
import { commands, editorFor, type EditorModel, renderDocumentHtml, runCommand } from "./index.js";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

function editorWith(texts: string[]): EditorModel {
  return editorFor(createDocx({ paragraphs: texts }));
}

function texts(model: EditorModel): string[] {
  return paragraphs(model.doc).map(paragraphText);
}

function select(model: EditorModel, a: [number, number], b: [number, number]): void {
  model.setSelection({
    anchor: { block: a[0], inline: 0, offset: a[1] },
    focus: { block: b[0], inline: 0, offset: b[1] },
  });
}

function hasSectPr(p: WmlParagraph): boolean {
  return !!p.pPr?.children.some((c) => c.kind === "element" && c.name.local === "sectPr");
}

describe("deleteSelectionCommand", () => {
  it("deletes inside one paragraph and leaves the caret at the start", () => {
    const model = editorWith(["Hello cruel world"]);
    select(model, [0, 6], [0, 12]);
    runCommand(model, commands.deleteSelectionCommand, undefined);
    expect(texts(model)).toEqual(["Hello world"]);
    expect(model.selection?.focus).toMatchObject({ block: 0, offset: 6 });
  });

  it("deletes across paragraphs (backwards selection) as one undo step", () => {
    const model = editorWith(["alpha", "beta", "gamma", "delta"]);
    select(model, [2, 2], [0, 2]);
    runCommand(model, commands.deleteSelectionCommand, undefined);
    expect(texts(model)).toEqual(["almma", "delta"]);
    expect(model.selection?.focus).toMatchObject({ block: 0, offset: 2 });
    expect(validate(openDocx(toUint8Array(model.doc)))).toHaveLength(0);
    model.undo();
    expect(texts(model)).toEqual(["alpha", "beta", "gamma", "delta"]);
  });

  it("keeps the section break that closes the last selected paragraph", () => {
    const doc = createDocx({ paragraphs: ["one", "two"] });
    appendSectionBreak(doc, "nextPage");
    const model = editorFor(doc);
    select(model, [1, 1], [2, 0]);
    runCommand(model, commands.deleteSelectionCommand, undefined);
    const ps = paragraphs(openDocx(toUint8Array(model.doc)));
    expect(ps.map(paragraphText)).toEqual(["one", "t"]);
    expect(ps.filter(hasSectPr)).toHaveLength(1);
  });

  it("is disabled for a collapsed caret", () => {
    const model = editorWith(["x"]);
    select(model, [0, 1], [0, 1]);
    expect(runCommand(model, commands.deleteSelectionCommand, undefined)).toBeUndefined();
    expect(model.canUndo()).toBe(false);
  });
});

describe("typing / Enter over a range replaces it", () => {
  it("insertTextCommand replaces the selection", () => {
    const model = editorWith(["one", "two"]);
    select(model, [0, 1], [1, 2]);
    runCommand(model, commands.insertTextCommand, { text: "X" });
    expect(texts(model)).toEqual(["oXo"]);
  });

  it("splitParagraphCommand replaces the selection with a paragraph break", () => {
    const model = editorWith(["abcdef"]);
    select(model, [0, 2], [0, 4]);
    runCommand(model, commands.splitParagraphCommand, undefined);
    expect(texts(model)).toEqual(["ab", "ef"]);
  });
});

describe("insertLineBreakCommand", () => {
  it("inserts the break at the caret, not at the paragraph end", () => {
    const model = editorWith(["HelloWorld"]);
    select(model, [0, 5], [0, 5]);
    runCommand(model, commands.insertLineBreakCommand, { kind: "line" });
    const para = paragraphs(model.doc)[0]!;
    const pieces = para.children.flatMap((c) => (c.kind === "run" ? c.pieces : []));
    expect(pieces.map((p) => p.kind)).toEqual(["text", "break", "text"]);
    // Caret sits on the run after the break, ready to keep typing.
    expect(model.selection?.focus).toMatchObject({ block: 0, inline: 2, offset: 0 });
    runCommand(model, commands.insertTextCommand, { text: "!" });
    expect(paragraphText(paragraphs(model.doc)[0]!)).toContain("!World");
  });
});

describe("rendering unmodelled inlines", () => {
  it("shows hyperlink and field result text read-only, without run anchors", () => {
    const doc = createDocx({ paragraphs: ["Body"] });
    addHyperlink(doc, "https://example.com", "Example <site>");
    appendField(doc, paragraphs(doc)[0]!, "PAGE", "7");
    const html = renderDocumentHtml(doc);
    // Wrapped runs keep their own formatting (the Hyperlink style here).
    expect(html).toMatch(
      /class="wk-link" contenteditable="false"><span style="[^"]*">Example &lt;site&gt;<\/span><\/span>/,
    );
    expect(html).toContain('class="wk-inline-raw" contenteditable="false">7</span>');
    expect(html).not.toContain("PAGE"); // field code (instrText) is not visible text
    // Raw inlines never get a run anchor, so typing cannot write into them.
    expect(html.match(/class="wk-run"/g)).toHaveLength(1);
  });

  it("anchors every paragraph of a multi-paragraph table cell", () => {
    const model = editorWith(["x"]);
    model.setSelection({ anchor: { block: 0 }, focus: { block: 0 } });
    runCommand(model, commands.insertTableCommand, { rows: 1, cols: 1 });
    const table = model.doc.document.body.blocks[1];
    if (table?.kind !== "table") throw new Error("expected a table");
    const cell = table.rows[0]!.cells[0]!;
    cell.paragraphs.push(structuredClone(cell.paragraphs[0]!));
    const html = renderDocumentHtml(model.doc);
    expect(html).toContain('data-wk-cell="0,0" data-wk-para="0"');
    expect(html).toContain('data-wk-cell="0,0" data-wk-para="1"');
  });
});

describe("element attributes through editing", () => {
  const rsid: XmlAttr = {
    name: { uri: W, local: "rsidR", prefix: "w" },
    value: "00AB12CD",
    isNamespaceDecl: false,
  };

  it("survive edit → save → reopen, and a split never copies them", () => {
    const doc = createDocx({ paragraphs: ["Hello world"] });
    paragraphs(doc)[0]!.attrs = [rsid];
    const model = editorFor(doc);
    select(model, [0, 5], [0, 5]);
    runCommand(model, commands.splitParagraphCommand, undefined);
    // The undo snapshot (a serialize/reparse clone) must keep them as well.
    model.undo();
    model.redo();
    const [first, second] = paragraphs(openDocx(toUint8Array(model.doc)));
    expect(first?.attrs).toEqual([rsid]);
    expect(second?.attrs).toBeUndefined();
  });
});

/** A position in row 0 of the table at block 1 (the `table cells` fixture). */
const at = (para: number, offset: number, col = 0) => ({
  block: 1,
  cell: { row: 0, col },
  para,
  inline: 0,
  offset,
});

describe("table cells", () => {
  function cellModel(): EditorModel {
    const model = editorWith(["before"]);
    model.setSelection({ anchor: { block: 0 }, focus: { block: 0 } });
    runCommand(model, commands.insertTableCommand, { rows: 1, cols: 2 });
    runCommand(model, commands.insertTextCommand, { text: "abcdef" });
    return model;
  }

  function cellTexts(model: EditorModel, col = 0): string[] {
    const table = model.doc.document.body.blocks[1];
    if (table?.kind !== "table") throw new Error("expected a table");
    return table.rows[0]!.cells[col]!.paragraphs.map(paragraphText);
  }

  it("Enter splits the cell paragraph instead of adding one after the table", () => {
    const model = cellModel();
    const blocks = model.doc.document.body.blocks.length;
    model.setSelection({ anchor: at(0, 3), focus: at(0, 3) });
    runCommand(model, commands.splitParagraphCommand, undefined);
    expect(cellTexts(model)).toEqual(["abc", "def"]);
    expect(model.doc.document.body.blocks).toHaveLength(blocks);
    expect(model.selection?.focus).toMatchObject({ block: 1, cell: { row: 0, col: 0 }, para: 1 });
    expect(validate(openDocx(toUint8Array(model.doc)))).toHaveLength(0);
  });

  it("Backspace at a later cell paragraph start merges within the cell", () => {
    const model = cellModel();
    model.setSelection({ anchor: at(0, 3), focus: at(0, 3) });
    runCommand(model, commands.splitParagraphCommand, undefined);
    runCommand(model, commands.mergeBackCommand, undefined);
    expect(cellTexts(model)).toEqual(["abcdef"]);
    expect(model.selection?.focus).toMatchObject({ para: 0, offset: 3 });
  });

  it("Backspace at the first cell paragraph is disabled (no cross-table merge)", () => {
    const model = cellModel();
    model.setSelection({ anchor: at(0, 0), focus: at(0, 0) });
    expect(runCommand(model, commands.mergeBackCommand, undefined)).toBeUndefined();
    expect(paragraphText(paragraphs(model.doc)[0]!)).toBe("before");
  });

  it("deletes a range across paragraphs of the same cell", () => {
    const model = cellModel();
    runCommand(model, commands.insertTextCommand, { text: "\nghi\njkl" });
    expect(cellTexts(model)).toEqual(["abcdef", "ghi", "jkl"]);
    model.setSelection({ anchor: at(0, 2), focus: at(2, 1) });
    runCommand(model, commands.deleteSelectionCommand, undefined);
    expect(cellTexts(model)).toEqual(["abkl"]);
  });

  it("refuses a range across cells atomically", () => {
    const model = cellModel();
    model.setSelection({ anchor: at(0, 1), focus: at(0, 0, 1) });
    expect(() => runCommand(model, commands.deleteSelectionCommand, undefined)).toThrow(
      /crosses table cells/,
    );
    expect(cellTexts(model)).toEqual(["abcdef"]);
    expect(model.selection).toEqual({ anchor: at(0, 1), focus: at(0, 0, 1) });
    expect(model.canUndo()).toBe(true); // only the setup commands remain
  });

  it("an IME-style replacement of a range is one undo step", () => {
    // The canvas applies a composition that started over a range as a single
    // insertTextCommand on the remembered range; this is that model contract.
    const model = cellModel();
    model.setSelection({ anchor: at(0, 1), focus: at(0, 4) });
    runCommand(model, commands.insertTextCommand, { text: "日本語" });
    expect(cellTexts(model)).toEqual(["a日本語ef"]);
    model.undo();
    expect(cellTexts(model)).toEqual(["abcdef"]);
    expect(model.selection).toEqual({ anchor: at(0, 1), focus: at(0, 4) });
  });
});
