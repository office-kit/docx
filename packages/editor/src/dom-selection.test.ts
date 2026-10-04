// @vitest-environment happy-dom
/**
 * DOM → document-position mapping on HTML produced by the real renderer:
 * carets on run text, on the paragraph element, inside read-only inline text
 * (hyperlinks), and in multi-paragraph table cells.
 */

import { addHyperlink, createDocx, paragraphs } from "@office-kit/docx";
import { describe, expect, it } from "vitest";
import { commands, editorFor, positionFromDom, renderDocumentHtml, runCommand } from "./index.js";

function mount(html: string): HTMLElement {
  const root = document.createElement("div");
  root.innerHTML = html;
  return root;
}

describe("positionFromDom", () => {
  it("maps run text offsets", () => {
    const root = mount(renderDocumentHtml(createDocx({ paragraphs: ["Hello"] })));
    const text = root.querySelector(".wk-run")!.firstChild!;
    expect(positionFromDom(text, 3)).toEqual({ block: 0, inline: 0, offset: 3 });
  });

  it("maps a point on the run element itself to a character offset", () => {
    const root = mount(renderDocumentHtml(createDocx({ paragraphs: ["Hello"] })));
    const span = root.querySelector(".wk-run")!;
    expect(positionFromDom(span, 1)).toEqual({ block: 0, inline: 0, offset: 5 });
  });

  it("maps a caret inside read-only link text to the end of the preceding run", () => {
    const doc = createDocx({ paragraphs: ["Before "] });
    addHyperlink(doc, "https://example.com", "link");
    // Move the link's inline into the first paragraph, after its run.
    const [first, second] = paragraphs(doc);
    first!.children.push(...second!.children);
    doc.document.body.blocks.pop();
    const root = mount(renderDocumentHtml(doc));
    const linkText = root.querySelector(".wk-link")!.firstChild!;
    expect(positionFromDom(linkText, 2)).toEqual({ block: 0, inline: 0, offset: 7 });
  });

  it("maps a child-index point on the paragraph element", () => {
    const root = mount(renderDocumentHtml(createDocx({ paragraphs: ["ab"] })));
    const p = root.querySelector(".wk-p")!;
    expect(positionFromDom(p, 1)).toEqual({ block: 0, inline: 0, offset: 2 });
    expect(positionFromDom(p, 0)).toEqual({ block: 0, inline: 0, offset: 0 });
  });

  it("keeps the paragraph index of a multi-paragraph table cell", () => {
    const model = editorFor(createDocx({ paragraphs: ["x"] }));
    model.setSelection({ anchor: { block: 0 }, focus: { block: 0 } });
    runCommand(model, commands.insertTableCommand, { rows: 1, cols: 1 });
    runCommand(model, commands.insertTextCommand, { text: "a\nb" });
    const root = mount(renderDocumentHtml(model.doc));
    const second = root.querySelector('.wk-run[data-wk-para="1"]')!.firstChild!;
    expect(positionFromDom(second, 1)).toEqual({
      block: 1,
      cell: { row: 0, col: 0 },
      para: 1,
      inline: 0,
      offset: 1,
    });
  });
});
