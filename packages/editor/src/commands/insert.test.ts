/**
 * Insert-tab commands, exercised through the command layer and a save/reopen
 * cycle, plus the canvas rendering of what they create (fields, symbols,
 * equations).
 */

import {
  addFootnote,
  appendHeading,
  appendParagraph,
  createDocx,
  ensureHeadingStyles,
  footers,
  headers,
  openDocx,
  paragraphs,
  toUint8Array,
  validatePackage,
  type Docx,
} from "@office-kit/docx";
import { describe, expect, it } from "vitest";
import { editorFor } from "../index.js";
import {
  documentBookmarks,
  documentCaptions,
  documentHeadings,
  documentNotes,
  documentNumberedItems,
  hyperlinkAtCaret,
  selectionText,
} from "../insert-queries.js";
import type { EditorModel } from "../model.js";
import { renderDocumentHtml } from "../render.js";
import { caretAt } from "../selection.js";
import { runAtPath } from "../text-edit.js";
import {
  formatPageNumbersCommand,
  insertFooterPresetCommand,
  insertHeaderPresetCommand,
  insertPageNumberCommand,
  removeHeaderCommand,
  removePageNumbersCommand,
  HEADER_FOOTER_PRESETS,
  COVER_PAGE_PRESETS,
} from "./header-footer.js";
import {
  addBookmarkAtSelectionCommand,
  dropCapCommand,
  editEquationCommand,
  insertBlankPageCommand,
  insertCoverPageCommand,
  insertCrossReferenceCommand,
  insertEquationCommand,
  insertFieldAtCaretCommand,
  insertLinkCommand,
  insertSymbolCommand,
  insertTextFromFileCommand,
  removeCoverPageCommand,
  removeLinkCommand,
  updateFieldsCommand,
} from "./insert.js";
import { insertTextCommand } from "./structure.js";
import { runCommand } from "./types.js";

function editorWith(texts: string[], caret = { block: 0, inline: 0, offset: 0 }): EditorModel {
  const model = editorFor(createDocx({ paragraphs: texts }));
  model.setSelection(caretAt(caret));
  return model;
}

function reopen(doc: Docx): { doc: Docx; xml: string } {
  const back = openDocx(toUint8Array(doc));
  expect(validatePackage(back.opc)).toEqual([]);
  const xml = new TextDecoder().decode(back.opc.parts.get(back.partName)?.data ?? new Uint8Array());
  return { doc: back, xml };
}

const NOW = new Date(2026, 0, 15, 9, 30);

describe("Insert ▸ fields", () => {
  it("inserts a field at the caret and keeps typing outside it", () => {
    const model = editorWith(["Date: "], { block: 0, inline: 0, offset: 6 });
    runCommand(model, insertFieldAtCaretCommand, {
      instruction: "DATE \\@ yyyy-MM-dd",
      context: { now: NOW },
    });
    runCommand(model, insertTextCommand, { text: "!" });
    const html = renderDocumentHtml(model.doc);
    expect(html).toContain("wk-fresult");
    expect(html).toContain(">2026-01-15<");
    expect(html).toContain(
      '<span class="wk-fcode" contenteditable="false"> DATE \\@ yyyy-MM-dd </span>',
    );
    const { xml } = reopen(model.doc);
    expect(xml).toMatch(
      /<w:t>2026-01-15<\/w:t><\/w:r><w:r><w:fldChar w:fldCharType="end"\/><\/w:r><w:r><w:t>!<\/w:t>/,
    );
  });

  it("updates fields through the command", () => {
    const model = editorWith([""]);
    runCommand(model, insertFieldAtCaretCommand, { instruction: "PAGE", result: "7" });
    expect(runCommand(model, updateFieldsCommand, { page: 2 })).toBe(1);
    expect(renderDocumentHtml(model.doc)).toContain(">2<");
  });

  it("replaces a selection with a link and removes it again", () => {
    const model = editorWith(["click here now"]);
    model.setSelection({
      anchor: { block: 0, inline: 0, offset: 6 },
      focus: { block: 0, inline: 0, offset: 10 },
    });
    runCommand(model, insertLinkCommand, { text: "here", target: { url: "https://example.com/" } });
    let { xml } = reopen(model.doc);
    expect(xml).toMatch(
      /<w:t xml:space="preserve">click <\/w:t><\/w:r><w:hyperlink r:id="rId\d+" w:history="1">/,
    );
    model.setSelection(caretAt({ block: 0, inline: 0, offset: 0 }));
    runCommand(model, removeLinkCommand, { index: 0 });
    ({ xml } = reopen(model.doc));
    expect(xml).not.toContain("hyperlink");
  });

  it("bookmarks a selection and cross-references it with above/below", () => {
    const model = editorWith(["Target text", "See "]);
    model.setSelection({
      anchor: { block: 0, inline: 0, offset: 0 },
      focus: { block: 0, inline: 0, offset: 6 },
    });
    runCommand(model, addBookmarkAtSelectionCommand, { name: "Goal" });
    model.setSelection(caretAt({ block: 1, inline: 0, offset: 4 }));
    runCommand(model, insertCrossReferenceCommand, {
      target: { kind: "bookmark", name: "Goal" },
      field: "REF",
      hyperlink: true,
      aboveBelow: true,
    });
    const html = renderDocumentHtml(model.doc);
    expect(html).toContain(">Target<");
    expect(html).toContain(">above<");
    const { xml } = reopen(model.doc);
    expect(xml).toContain(" REF Goal \\h ");
    expect(xml).toContain(" REF Goal \\p \\h ");
  });

  it("cross-references a heading through a hidden _Ref bookmark", () => {
    const model = editorWith(["Chapter one", "As seen in "]);
    const heading = paragraphs(model.doc)[0];
    if (!heading) throw new Error("no heading");
    model.setSelection(caretAt({ block: 1, inline: 0, offset: 11 }));
    runCommand(model, insertCrossReferenceCommand, {
      target: { kind: "paragraph", paragraph: heading },
      field: "REF",
    });
    const { xml } = reopen(model.doc);
    expect(xml).toMatch(/<w:bookmarkStart w:id="\d+" w:name="_Ref\d{9}"\/>/);
    expect(renderDocumentHtml(model.doc)).toContain(">Chapter one<");
  });
});

describe("Insert ▸ symbols and equations", () => {
  it("renders a symbol-font character and types after it", () => {
    const model = editorWith(["ab"], { block: 0, inline: 0, offset: 1 });
    runCommand(model, insertSymbolCommand, { symbol: { font: "Symbol", char: "F061" } });
    const html = renderDocumentHtml(model.doc);
    expect(html).toContain('<span class="wk-sym" contenteditable="false"');
    expect(html).toContain(">α<");
    // The caret moved past the symbol: typing lands after it.
    runCommand(model, insertTextCommand, { text: "c" });
    const pos = model.selection?.focus;
    expect(pos && runAtPath(model.doc, pos)?.pieces).toEqual([
      { kind: "text", value: "cb", preserveSpace: false },
    ]);
  });

  it("inserts a display equation in an empty paragraph and edits it", () => {
    const model = editorWith([""]);
    runCommand(model, insertEquationCommand, { linear: "x=(-b±√(b^2-4ac))/2a" });
    let html = renderDocumentHtml(model.doc);
    expect(html).toContain(
      '<span class="wk-math wk-math-display" contenteditable="false" data-wk-math-block="0" data-wk-math="0"><math display="block">',
    );
    expect(html).toContain("<mfrac>");
    expect(html).toContain("<msqrt>");
    runCommand(model, editEquationCommand, { at: { block: 0 }, index: 0, linear: "∑_(k=1)^n▒k" });
    html = renderDocumentHtml(model.doc);
    expect(html).toContain('<munderover><mo largeop="true" movablelimits="false">∑</mo>');
    const { xml } = reopen(model.doc);
    expect(xml).toContain("<m:oMathPara");
  });

  it("inserts an inline equation into text", () => {
    const model = editorWith(["Area  m²"], { block: 0, inline: 0, offset: 5 });
    runCommand(model, insertEquationCommand, { linear: "πr^2" });
    const html = renderDocumentHtml(model.doc);
    expect(html).toContain(
      '<span class="wk-math" contenteditable="false" data-wk-math-block="0" data-wk-math="0"><math><mrow><mi>π</mi><msup>',
    );
  });
});

describe("Insert ▸ pages, drop cap, text from file", () => {
  it("inserts and removes each cover page design", () => {
    for (const preset of COVER_PAGE_PRESETS) {
      const model = editorWith(["Body"]);
      runCommand(model, insertCoverPageCommand, { preset });
      expect(reopen(model.doc).xml).toContain("_CoverPage");
      runCommand(model, removeCoverPageCommand, undefined);
      expect(paragraphs(model.doc)).toHaveLength(1);
    }
  });

  it("adds a blank page after the caret paragraph", () => {
    const model = editorWith(["one", "two"]);
    runCommand(model, insertBlankPageCommand, undefined);
    expect(paragraphs(model.doc)).toHaveLength(5);
    expect(model.selection?.focus.block).toBe(2);
    expect(reopen(model.doc).xml.match(/w:type="page"/g)).toHaveLength(2);
  });

  it("drops and restores a capital", () => {
    const model = editorWith(["Hello"]);
    runCommand(model, dropCapCommand, { position: "drop", lines: 2 });
    expect(paragraphs(model.doc)).toHaveLength(2);
    expect(model.selection?.focus.block).toBe(1);
    runCommand(model, dropCapCommand, { position: "none" });
    expect(paragraphs(model.doc)).toHaveLength(1);
  });

  it("inserts another document's body at the caret", () => {
    const source = toUint8Array(createDocx({ paragraphs: ["Imported A", "Imported B"] }));
    const model = editorWith(["HeadTail"], { block: 0, inline: 0, offset: 4 });
    expect(runCommand(model, insertTextFromFileCommand, { bytes: source })).toBe(2);
    const texts = paragraphs(model.doc).map((p) =>
      p.children
        .map((c) =>
          c.kind === "run" ? c.pieces.map((x) => (x.kind === "text" ? x.value : "")).join("") : "",
        )
        .join(""),
    );
    expect(texts).toEqual(["Head", "Imported A", "Imported B", "Tail"]);
  });
});

describe("Insert ▸ Header & Footer", () => {
  it("applies every header and footer preset", () => {
    for (const preset of HEADER_FOOTER_PRESETS) {
      const model = editorWith(["Body"]);
      runCommand(model, insertHeaderPresetCommand, { preset });
      runCommand(model, insertFooterPresetCommand, { preset });
      const { doc } = reopen(model.doc);
      expect(headers(doc)).toHaveLength(1);
      expect(footers(doc)).toHaveLength(1);
    }
  });

  it("puts page numbers in, formats and removes them", () => {
    const model = editorWith(["Body"]);
    runCommand(model, insertHeaderPresetCommand, { preset: "blank" });
    runCommand(model, insertPageNumberCommand, { position: "top", align: "right" });
    runCommand(model, formatPageNumbersCommand, { format: "upperRoman", start: 1 });
    expect(reopen(model.doc).xml).toContain('<w:pgNumType w:fmt="upperRoman" w:start="1"/>');
    expect(runCommand(model, removePageNumbersCommand, undefined)).toBe(1);
    runCommand(model, removeHeaderCommand, {});
    expect(headers(reopen(model.doc).doc)).toHaveLength(0);
  });
});

describe("Insert dialogs ▸ document targets", () => {
  it("lists headings, bookmarks, notes and captions and links to a heading", () => {
    const doc = createDocx({ paragraphs: ["Intro"] });
    ensureHeadingStyles(doc, 2);
    appendHeading(doc, "Methods", 1);
    appendHeading(doc, "Setup", 2);
    const noted = appendParagraph(doc, "A claim");
    addFootnote(doc, noted, "Source");
    const caption = appendParagraph(doc, "Figure ");
    const model = editorFor(doc);
    model.setSelection(caretAt({ block: 4, inline: 0, offset: 7 }));
    runCommand(model, insertFieldAtCaretCommand, { instruction: "SEQ Figure \\* ARABIC" });
    expect(documentCaptions(model.doc, "Figure").map((c) => c.paragraph)).toEqual([caption]);
    expect(documentCaptions(model.doc, "Table")).toEqual([]);
    expect(documentHeadings(model.doc).map((h) => [h.label, h.level, h.block])).toEqual([
      ["Methods", 1, 1],
      ["Setup", 2, 2],
    ]);
    expect(documentNotes(model.doc, "footnote").map((n) => n.number)).toEqual([1]);
    expect(documentNumberedItems(model.doc)).toEqual([]);

    model.setSelection(caretAt({ block: 0, inline: 0, offset: 5 }));
    const heading = documentHeadings(model.doc)[1]?.paragraph;
    if (!heading) throw new Error("no heading");
    runCommand(model, insertLinkCommand, { text: " (see Setup)", target: {}, heading });
    const hidden = documentBookmarks(model.doc).filter((b) => b.hidden);
    expect(hidden).toHaveLength(1);
    expect(hidden[0]?.paragraph).toBe(heading);
    const { xml } = reopen(model.doc);
    expect(xml).toContain(`<w:hyperlink w:anchor="${hidden[0]?.name}"`);
    model.setSelection(caretAt({ block: 0, inline: 0, offset: 0 }));
    expect(hyperlinkAtCaret(model)).toBeUndefined();
  });

  it("reads the selected text and the link at the caret", () => {
    const model = editorWith(["one two", "three"]);
    model.setSelection({
      anchor: { block: 0, inline: 0, offset: 4 },
      focus: { block: 1, inline: 0, offset: 3 },
    });
    expect(selectionText(model)).toBe("two\nthr");
    model.setSelection({
      anchor: { block: 0, inline: 0, offset: 4 },
      focus: { block: 0, inline: 0, offset: 7 },
    });
    runCommand(model, insertLinkCommand, { text: "two", target: { url: "https://example.com/" } });
    model.setSelection(caretAt({ block: 0, inline: 0, offset: 4 }));
    expect(hyperlinkAtCaret(model)?.link.url).toBe("https://example.com/");
  });
});
