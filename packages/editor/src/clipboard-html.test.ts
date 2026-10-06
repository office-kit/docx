// @vitest-environment happy-dom
/**
 * Clipboard HTML → pasted blocks, on the HTML that Word, Google Docs, web
 * pages and this editor's canvas put on the clipboard; and pasting those
 * blocks into a document.
 */

import {
  createDocx,
  getParagraphNumbering,
  getParagraphStyle,
  getRunFormat,
  openDocx,
  paragraphs,
  paragraphText,
  toUint8Array,
  validate,
  type WmlParagraph,
  type WmlRun,
} from "@office-kit/docx";
import { describe, expect, it } from "vitest";
import { commands, editorFor, type EditorModel, parseClipboardHtml, runCommand } from "./index.js";

function runsOf(p: WmlParagraph): Array<[string, string]> {
  return p.children
    .filter((c): c is WmlRun => c.kind === "run")
    .map((r) => {
      const f = getRunFormat(r);
      const flags = `${f.bold ? "B" : ""}${f.italic ? "I" : ""}${f.underline ? "U" : ""}`;
      return [
        r.pieces
          .map((p) => (p.kind === "text" ? p.value : p.kind === "break" ? "\n" : ""))
          .join(""),
        flags,
      ];
    });
}

function caretModel(texts: string[], block: number, offset: number): EditorModel {
  const model = editorFor(createDocx({ paragraphs: texts }));
  const pos = { block, inline: 0, offset };
  model.setSelection({ anchor: pos, focus: pos });
  return model;
}

function paste(model: EditorModel, html: string): void {
  runCommand(model, commands.insertFragmentCommand, { blocks: parseClipboardHtml(html) });
}

describe("parseClipboardHtml", () => {
  it("reads inline formatting from tags and styles", () => {
    expect(
      parseClipboardHtml(
        '<p>a <b>bold</b> <span style="font-style:italic;color:#c00">red</span> <u>u</u><sup>2</sup></p>',
      ),
    ).toEqual([
      {
        kind: "paragraph",
        runs: [
          { text: "a ", format: {} },
          { text: "bold", format: { bold: true } },
          { text: " ", format: {} },
          { text: "red", format: { italic: true, color: "CC0000" } },
          { text: " ", format: {} },
          { text: "u", format: { underline: true } },
          { text: "2", format: { verticalAlign: "superscript" } },
        ],
      },
    ]);
  });

  it("lets Google Docs' font-weight:normal wrapper cancel <b>", () => {
    const [p] = parseClipboardHtml(
      '<b style="font-weight:normal" id="docs-internal-guid-1"><p><span>plain</span> <span style="font-weight:700">bold</span></p></b>',
    );
    expect(p).toEqual({
      kind: "paragraph",
      runs: [
        { text: "plain ", format: { bold: false } },
        { text: "bold", format: { bold: true } },
      ],
    });
  });

  it("reads headings, nested lists, line breaks and collapses whitespace", () => {
    const blocks = parseClipboardHtml(`
      <h2>Title</h2>
      <ul><li>one<br>more</li><li>two<ol><li>deep</li></ol></li></ul>
      <p>  spaced
         out  </p>`);
    expect(blocks).toEqual([
      { kind: "paragraph", heading: 2, runs: [{ text: "Title", format: {} }] },
      {
        kind: "paragraph",
        list: { kind: "bullet", level: 0 },
        runs: [{ text: "one\nmore", format: {} }],
      },
      {
        kind: "paragraph",
        list: { kind: "bullet", level: 0 },
        runs: [{ text: "two", format: {} }],
      },
      {
        kind: "paragraph",
        list: { kind: "numbered", level: 1 },
        runs: [{ text: "deep", format: {} }],
      },
      { kind: "paragraph", runs: [{ text: "spaced out", format: {} }] },
    ]);
  });

  it("reads Word's list paragraphs and drops their typed markers", () => {
    const blocks = parseClipboardHtml(
      `<p class=MsoListParagraph style='mso-list:l0 level1 lfo1'><![if !supportLists]><span style='mso-list:Ignore'>1.<span>&nbsp;&nbsp;</span></span><![endif]>First</p>
       <p class=MsoListParagraph style='mso-list:l1 level2 lfo2'><span style='mso-list:Ignore'>o<span>&nbsp;</span></span>Sub</p>`,
    );
    expect(blocks).toEqual([
      {
        kind: "paragraph",
        list: { kind: "numbered", level: 0 },
        runs: [{ text: "First", format: {} }],
      },
      {
        kind: "paragraph",
        list: { kind: "bullet", level: 1 },
        runs: [{ text: "Sub", format: {} }],
      },
    ]);
  });

  it("reads tables cell by cell", () => {
    const [table] = parseClipboardHtml(
      "<table><tr><th>h</th><td><p>a</p><p>b</p></td></tr><tr><td>c</td></tr></table>",
    );
    expect(table).toEqual({
      kind: "table",
      rows: [
        [
          [{ kind: "paragraph", runs: [{ text: "h", format: { bold: true } }] }],
          [
            { kind: "paragraph", runs: [{ text: "a", format: {} }] },
            { kind: "paragraph", runs: [{ text: "b", format: {} }] },
          ],
        ],
        [[{ kind: "paragraph", runs: [{ text: "c", format: {} }] }]],
      ],
    });
  });

  it("reads this editor's own copy: no list label, tabs kept", () => {
    const blocks = parseClipboardHtml(
      '<p class="wk-p"><span class="wk-list-label" contenteditable="false">•\t</span><span class="wk-run">a<span class="wk-tab" contenteditable="false">\t</span>b​</span></p>',
    );
    expect(blocks).toEqual([
      {
        kind: "paragraph",
        list: { kind: "bullet", level: 0 },
        runs: [{ text: "a\tb", format: {} }],
      },
    ]);
  });
});

describe("insertFragmentCommand", () => {
  it("pastes formatted text inside a paragraph as one undo step", () => {
    const model = caretModel(["Hello world"], 0, 6);
    paste(model, "<b>big</b>&nbsp;");
    expect(runsOf(paragraphs(model.doc)[0]!)).toEqual([
      ["Hello ", ""],
      ["big", "B"],
      [" ", ""],
      ["world", ""],
    ]);
    expect(model.selection?.focus).toMatchObject({ inline: 2, offset: 1 });
    model.undo();
    expect(paragraphText(paragraphs(model.doc)[0]!)).toBe("Hello world");
  });

  it("splits the caret paragraph around several pasted paragraphs (Word)", () => {
    const model = caretModel(["xyz"], 0, 2);
    paste(model, "<p>A</p><h1>B</h1><ul><li>C</li><li>D</li></ul>");
    const ps = paragraphs(model.doc);
    expect(ps.map(paragraphText)).toEqual(["xyA", "B", "C", "Dz"]);
    expect(getParagraphStyle(ps[1]!)).toBe("Heading1");
    const lists = ps.slice(2).map((p) => getParagraphNumbering(p)?.numId);
    expect(lists[0]).toBeDefined();
    expect(lists[1]).toBe(lists[0]);
    expect(validate(openDocx(toUint8Array(model.doc)))).toHaveLength(0);
  });

  it("gives the first of several pasted paragraphs their list too, leaving no empty runs", () => {
    const model = caretModel(["x", ""], 1, 0);
    paste(model, "<ol><li>one</li><li>two</li></ol>");
    const [, first, second] = paragraphs(model.doc);
    const numIds = [first, second].map((p) => getParagraphNumbering(p!)?.numId);
    expect(numIds[0]).toBeDefined();
    expect(numIds[1]).toBe(numIds[0]);
    expect([runsOf(first!), runsOf(second!)]).toEqual([[["one", ""]], [["two", ""]]]);
  });

  it("keeps the destination's paragraph formatting for text pasted within a paragraph", () => {
    const model = caretModel(["x"], 0, 1);
    paste(model, "<h1>big</h1>");
    expect(getParagraphStyle(paragraphs(model.doc)[0]!)).toBeUndefined();
    expect(paragraphText(paragraphs(model.doc)[0]!)).toBe("xbig");
  });

  it("puts a pasted table between the halves of the paragraph", () => {
    const model = caretModel(["before", "after"], 0, 3);
    paste(model, "<table><tr><td><i>a</i></td><td>b</td></tr></table>");
    const kinds = model.doc.document.body.blocks.map((b) => b.kind);
    expect(kinds).toEqual(["paragraph", "table", "paragraph", "paragraph"]);
    expect(paragraphs(model.doc).map(paragraphText)).toEqual(["bef", "ore", "after"]);
    const table = model.doc.document.body.blocks[1];
    if (table?.kind !== "table") throw new Error("expected a table");
    expect(runsOf(table.rows[0]!.cells[0]!.paragraphs[0]!)).toEqual([["a", "I"]]);
    expect(model.selection?.focus).toMatchObject({ block: 2, offset: 0 });
    const reopened = openDocx(toUint8Array(model.doc));
    expect(validate(reopened)).toHaveLength(0);
    expect(reopened.document.body.blocks.map((b) => b.kind)).toEqual(kinds);
  });

  it("puts a table pasted at a paragraph start before it, with no empty paragraph", () => {
    const model = caretModel(["text"], 0, 0);
    paste(model, "<table><tr><td>a</td></tr></table>");
    expect(model.doc.document.body.blocks.map((b) => b.kind)).toEqual(["table", "paragraph"]);
    expect(paragraphs(model.doc).map(paragraphText)).toEqual(["text"]);
  });

  it("pastes a table's cells as paragraphs inside a table cell", () => {
    const model = caretModel(["x"], 0, 1);
    runCommand(model, commands.insertTableCommand, { rows: 1, cols: 1 });
    paste(model, "<table><tr><td>a</td><td>b</td></tr></table>");
    const table = model.doc.document.body.blocks[1];
    if (table?.kind !== "table") throw new Error("expected a table");
    expect(table.rows[0]!.cells[0]!.paragraphs.map(paragraphText)).toEqual(["a", "b"]);
  });
});
