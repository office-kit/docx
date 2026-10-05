/**
 * Editing inside headers, footers and notes: the selection carries a story,
 * and the ordinary commands (typing, Enter, Backspace, formatting, undo) act
 * on that story instead of the body.
 */

import {
  addFootnote,
  appendParagraph,
  appendSectionBreak,
  createDocx,
  getRunFormat,
  openDocx,
  paragraphText,
  resolveHeaderFooter,
  storyBody,
  type StoryRef,
  toUint8Array,
  validatePackage,
  type WmlParagraph,
} from "@office-kit/docx";
import { describe, expect, it } from "vitest";
import {
  differentFirstPageCommand,
  differentOddEvenCommand,
  editHeaderFooterCommand,
  headerDistanceCommand,
  headerFooterLinked,
  linkToPreviousCommand,
} from "./commands/stories.js";
import {
  insertTextCommand,
  mergeBackCommand,
  splitParagraphCommand,
} from "./commands/structure.js";
import { toggleBoldCommand } from "./commands/text.js";
import { runCommand } from "./commands/types.js";
import { paragraphAt, paragraphsInRange } from "./doc-access.js";
import { runAtPath } from "./text-edit.js";
import { editorFor } from "./index.js";
import { documentSections } from "./layout/sections.js";
import { orderSelection } from "./selection.js";
import { runsInRange } from "./selection-runs.js";

function storyTexts(model: ReturnType<typeof editorFor>, story: StoryRef): string[] {
  return (storyBody(model.document, story)?.blocks ?? [])
    .filter((b): b is WmlParagraph => b.kind === "paragraph")
    .map(paragraphText);
}

/** The story's paragraph texts after saving and reopening the document. */
function savedTexts(doc: ReturnType<typeof createDocx>, story: StoryRef): string[] {
  return (storyBody(openDocx(toUint8Array(doc)), story)?.blocks ?? [])
    .filter((b): b is WmlParagraph => b.kind === "paragraph")
    .map(paragraphText);
}

function headerModel() {
  const model = editorFor(createDocx({ paragraphs: ["Body text"] }));
  runCommand(model, editHeaderFooterCommand, { section: 0, kind: "header", type: "default" });
  const story = model.selection?.focus.story;
  if (story?.kind !== "header") throw new Error("caret is not in the header");
  return { model, story };
}

describe("header editing", () => {
  it("creates the header and puts the caret in it", () => {
    const { model, story } = headerModel();
    expect(model.story).toEqual(story);
    expect(resolveHeaderFooter(model.document, 0, "header", "default")?.partName).toBe(
      story.partName,
    );
    expect(storyTexts(model, story)).toEqual([""]);
  });

  it("types, splits and merges paragraphs in the header, leaving the body alone", () => {
    const { model, story } = headerModel();
    runCommand(model, insertTextCommand, { text: "Annual report" });
    expect(model.selection?.focus.story).toEqual(story);
    expect(storyTexts(model, story)).toEqual(["Annual report"]);

    model.setSelection({
      anchor: { story, block: 0, inline: 0, offset: 6 },
      focus: { story, block: 0, inline: 0, offset: 6 },
    });
    runCommand(model, splitParagraphCommand, undefined);
    expect(storyTexts(model, story)).toEqual(["Annual", " report"]);
    expect(model.selection?.focus).toMatchObject({ story, block: 1 });

    runCommand(model, mergeBackCommand, undefined);
    expect(storyTexts(model, story)).toEqual(["Annual report"]);

    expect(model.document.document.body.blocks).toHaveLength(1);
    expect(savedTexts(model.document, story)).toEqual(["Annual report"]);
    expect(validatePackage(openDocx(toUint8Array(model.document)).opc)).toEqual([]);
  });

  it("formats a selection in the header and reads it back through doc-access", () => {
    const { model, story } = headerModel();
    runCommand(model, insertTextCommand, { text: "Hello world" });
    model.setSelection({
      anchor: { story, block: 0, inline: 0, offset: 0 },
      focus: { story, block: 0, inline: 0, offset: 5 },
    });
    runCommand(model, toggleBoldCommand, undefined);
    const sel = model.selection;
    if (!sel) throw new Error("no selection");
    const runs = runsInRange(model.doc, orderSelection(sel));
    expect(runs.map((r) => getRunFormat(r).bold)).toEqual([true]);
    expect(paragraphsInRange(model.doc, orderSelection(sel))).toHaveLength(1);
    expect(paragraphAt(model.doc, sel.focus)).toBe(storyBody(model.document, story)?.blocks[0]);
    expect(runAtPath(model.doc, { story, block: 0, inline: 1 })).toBeDefined();
    // The body paragraph is untouched.
    const body = model.document.document.body.blocks[0];
    expect(body?.kind === "paragraph" && paragraphText(body)).toBe("Body text");
  });

  it("undoes header edits", () => {
    const { model, story } = headerModel();
    runCommand(model, insertTextCommand, { text: "Draft" });
    expect(storyTexts(model, story)).toEqual(["Draft"]);
    model.undo();
    expect(storyTexts(model, story)).toEqual([""]);
    model.redo();
    expect(storyTexts(model, story)).toEqual(["Draft"]);
  });

  it("toggles first-page / odd-even headers and sets distances", () => {
    const { model } = headerModel();
    runCommand(model, differentFirstPageCommand, { section: 0, on: true });
    runCommand(model, differentOddEvenCommand, { on: true });
    runCommand(model, headerDistanceCommand, { section: 0, twips: 1080 });
    expect(differentOddEvenCommand.isActive?.(model)).toBe(true);
    const [section] = documentSections(model.document);
    expect(section?.titlePage).toBe(true);
    expect(section?.margins.header).toBe(1080);
    expect(() => runCommand(model, headerDistanceCommand, { section: 0, twips: -1 })).toThrow();
  });

  it("links and unlinks a section's header", () => {
    const doc = createDocx({ paragraphs: ["One"] });
    appendSectionBreak(doc, "nextPage");
    appendParagraph(doc, "Two");
    const model = editorFor(doc);
    runCommand(model, editHeaderFooterCommand, { section: 1, kind: "header", type: "default" });
    const target = { section: 1, kind: "header", type: "default" } as const;
    expect(headerFooterLinked(model, target)).toBe(false);
    // Created on the section itself? No: the first section had none either,
    // so the new header belongs to section 1 and section 0 shows nothing.
    expect(resolveHeaderFooter(model.document, 0, "header", "default")).toBeUndefined();
    runCommand(model, linkToPreviousCommand, { ...target, linked: true });
    expect(headerFooterLinked(model, target)).toBe(true);
    runCommand(model, linkToPreviousCommand, { ...target, linked: false });
    expect(headerFooterLinked(model, target)).toBe(false);
    expect(model.selection?.focus.story?.kind).toBe("header");
  });
});

describe("footnote editing", () => {
  it("types into a footnote story", () => {
    const doc = createDocx({ paragraphs: ["Body"] });
    const para = doc.document.body.blocks[0];
    if (para?.kind !== "paragraph") throw new Error("expected a paragraph");
    const id = addFootnote(doc, para, "Note");
    const model = editorFor(doc);
    const story = { kind: "footnote", id } as const;
    model.setSelection({
      anchor: { story, block: 0, inline: 0, offset: 5 },
      focus: { story, block: 0, inline: 0, offset: 5 },
    });
    // Enter after the note text (" Note": the library leads with a space),
    // then type a second paragraph.
    runCommand(model, splitParagraphCommand, undefined);
    runCommand(model, insertTextCommand, { text: "More" });
    expect(storyTexts(model, story)).toEqual([" Note", "More"]);
    expect(savedTexts(model.document, story)).toEqual([" Note", "More"]);
    expect(model.document.document.body.blocks).toHaveLength(1);
  });
});
