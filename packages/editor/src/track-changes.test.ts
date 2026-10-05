import {
  comments,
  createDocx,
  documentProtection,
  openDocx,
  paragraphs,
  revisions,
  text,
  toUint8Array,
  validate,
} from "@office-kit/docx";
import { describe, expect, it } from "vitest";
import { alignCenterCommand } from "./commands/paragraph.js";
import {
  acceptRevisionsCommand,
  addCommentCommand,
  protectDocumentCommand,
  rejectRevisionsCommand,
  toggleTrackChangesCommand,
  trackedDeleteCommand,
  unprotectDocumentCommand,
} from "./commands/review.js";
import {
  deleteSelectionCommand,
  insertTextCommand,
  mergeBackCommand,
  splitParagraphCommand,
} from "./commands/structure.js";
import { toggleBoldCommand } from "./commands/text.js";
import { runCommand } from "./commands/types.js";
import { demoteCommand, moveBlocksCommand, promoteCommand } from "./commands/view.js";
import { editorFor } from "./index.js";
import type { EditorModel } from "./model.js";
import { outlineLevelOf } from "./outline.js";
import { renderDocumentHtml } from "./render.js";
import { adjacentReviewMark, revisionIdsAtSelection } from "./review-nav.js";
import { caretAt } from "./selection.js";
import { isTrackingRevisions, setReviewer } from "./track-changes.js";

function tracked(texts: string[]): EditorModel {
  const model = editorFor(createDocx({ paragraphs: texts }));
  setReviewer(model, { author: "Ann", initials: "A" });
  model.setSelection(caretAt({ block: 0, inline: 0, offset: 0 }));
  runCommand(model, toggleTrackChangesCommand, undefined);
  return model;
}

/** The text a reader sees with All Markup off: deleted runs left out. */
function visible(model: EditorModel): string {
  return paragraphs(model.doc)
    .map((p) =>
      p.children
        .map((c) =>
          c.kind === "run" && c.revision?.kind !== "del"
            ? c.pieces.map((x) => (x.kind === "text" ? x.value : "")).join("")
            : "",
        )
        .join(""),
    )
    .join("\n");
}

function valid(model: EditorModel): void {
  expect(validate(openDocx(toUint8Array(model.doc)))).toEqual([]);
}

describe("Track Changes recording", () => {
  it("toggles w:trackRevisions", () => {
    const model = tracked(["x"]);
    expect(isTrackingRevisions(model)).toBe(true);
    runCommand(model, toggleTrackChangesCommand, undefined);
    expect(isTrackingRevisions(model)).toBe(false);
  });

  it("records typing at the caret as one insertion by the reviewer", () => {
    const model = tracked(["Hello world"]);
    model.setSelection(caretAt({ block: 0, inline: 0, offset: 5 }));
    runCommand(model, insertTextCommand, { text: "," });
    runCommand(model, insertTextCommand, { text: " there" });
    expect(text(model.doc)).toBe("Hello, there world");
    expect(revisions(model.doc).map((r) => [r.kind, r.author, r.text])).toEqual([
      ["insert", "Ann", ", there"],
    ]);
    valid(model);
  });

  it("typing over a selection marks it deleted and inserts after it", () => {
    const model = tracked(["one two three"]);
    model.setSelection({
      anchor: { block: 0, inline: 0, offset: 4 },
      focus: { block: 0, inline: 0, offset: 7 },
    });
    runCommand(model, insertTextCommand, { text: "2" });
    expect(revisions(model.doc).map((r) => [r.kind, r.text])).toEqual([
      ["delete", "two"],
      ["insert", "2"],
    ]);
    expect(visible(model)).toBe("one 2 three");
  });

  it("Backspace marks characters deleted and skips already deleted text", () => {
    const model = tracked(["abc"]);
    model.setSelection(caretAt({ block: 0, inline: 0, offset: 3 }));
    runCommand(model, trackedDeleteCommand, { direction: -1 });
    runCommand(model, trackedDeleteCommand, { direction: -1 });
    const list = revisions(model.doc);
    expect(list.map((r) => r.text).join("")).toBe("bc");
    expect(visible(model)).toBe("a");
    runCommand(model, rejectRevisionsCommand, { ids: list.map((r) => r.id) });
    expect(text(model.doc)).toBe("abc");
  });

  it("deleting a selection across paragraphs marks the paragraph mark", () => {
    const model = tracked(["first", "second"]);
    model.setSelection({
      anchor: { block: 0, inline: 0, offset: 3 },
      focus: { block: 1, inline: 0, offset: 2 },
    });
    runCommand(model, deleteSelectionCommand, undefined);
    expect(revisions(model.doc).map((r) => r.kind)).toEqual([
      "delete",
      "paragraphDelete",
      "delete",
    ]);
    runCommand(model, acceptRevisionsCommand, { ids: revisions(model.doc).map((r) => r.id) });
    expect(text(model.doc)).toBe("fircond");
  });

  it("Enter records an inserted paragraph mark; Backspace over it joins again", () => {
    const model = tracked(["abcd"]);
    model.setSelection(caretAt({ block: 0, inline: 0, offset: 2 }));
    runCommand(model, splitParagraphCommand, undefined);
    expect(paragraphs(model.doc)).toHaveLength(2);
    expect(revisions(model.doc).map((r) => [r.kind, r.block])).toEqual([["paragraphInsert", 0]]);
    // Backspace at the new paragraph's start removes the reviewer's own mark.
    runCommand(model, mergeBackCommand, undefined);
    expect(paragraphs(model.doc)).toHaveLength(1);
    expect(revisions(model.doc)).toEqual([]);
  });

  it("Backspace at an original paragraph start marks the mark deleted", () => {
    const model = tracked(["a", "b"]);
    model.setSelection(caretAt({ block: 1, inline: 0, offset: 0 }));
    runCommand(model, mergeBackCommand, undefined);
    expect(paragraphs(model.doc)).toHaveLength(2);
    expect(revisions(model.doc).map((r) => r.kind)).toEqual(["paragraphDelete"]);
    expect(model.selection?.focus).toMatchObject({ block: 0, offset: 1 });
  });

  it("records formatting changes as rPrChange / pPrChange", () => {
    const model = tracked(["styled"]);
    model.setSelection({
      anchor: { block: 0, inline: 0, offset: 0 },
      focus: { block: 0, inline: 0, offset: 3 },
    });
    runCommand(model, toggleBoldCommand, undefined);
    runCommand(model, alignCenterCommand, undefined);
    expect(revisions(model.doc).map((r) => r.kind)).toEqual(["format", "format"]);
    runCommand(model, rejectRevisionsCommand, { ids: revisions(model.doc).map((r) => r.id) });
    expect(revisions(model.doc)).toEqual([]);
    valid(model);
  });
});

describe("review navigation", () => {
  it("finds the change at the caret and the next one", () => {
    const model = tracked(["aaa bbb"]);
    model.setSelection(caretAt({ block: 0, inline: 0, offset: 7 }));
    runCommand(model, insertTextCommand, { text: "!" });
    model.setSelection(caretAt({ block: 0, inline: 0, offset: 0 }));
    const next = adjacentReviewMark(model, "revision", 1);
    expect(next?.position).toMatchObject({ block: 0, inline: 1 });
    model.setSelection(caretAt(next!.position));
    expect(revisionIdsAtSelection(model)).toEqual([next!.id]);
    runCommand(model, acceptRevisionsCommand, {});
    expect(revisions(model.doc)).toEqual([]);
  });
});

describe("comments", () => {
  it("comments on the word at the caret", () => {
    const model = tracked(["one two three"]);
    model.setSelection(caretAt({ block: 0, inline: 0, offset: 5 }));
    runCommand(model, addCommentCommand, { text: "Hm" });
    expect(comments(model.doc)).toMatchObject([{ author: "Ann", initials: "A", text: "Hm" }]);
    expect(renderDocumentHtml(model.doc)).toMatch(/class="wk-run wk-commented"[^>]*>two</);
    valid(model);
  });
});

describe("rendering", () => {
  it("marks insertions, deletions and changed paragraphs", () => {
    const model = tracked(["abc"]);
    model.setSelection({
      anchor: { block: 0, inline: 0, offset: 0 },
      focus: { block: 0, inline: 0, offset: 1 },
    });
    runCommand(model, insertTextCommand, { text: "X" });
    const html = renderDocumentHtml(model.doc);
    expect(html).toContain('class="wk-p wk-changed"');
    expect(html).toMatch(
      /<span class="wk-del-run wk-del wk-author-0" contenteditable="false"[^>]*>a<\/span>/,
    );
    expect(html).toMatch(/<span class="wk-run wk-ins wk-author-0"[^>]*>X<\/span>/);
  });
});

describe("protection", () => {
  it("forces tracking while protected for tracked changes and checks the password", () => {
    const model = editorFor(createDocx({ paragraphs: ["x"] }));
    model.setSelection(caretAt({ block: 0 }));
    runCommand(model, protectDocumentCommand, {
      edit: "trackedChanges",
      password: "pw",
      spinCount: 5,
    });
    expect(isTrackingRevisions(model)).toBe(true);
    expect(toggleTrackChangesCommand.isEnabled?.(model)).toBe(false);
    expect(() => runCommand(model, unprotectDocumentCommand, { password: "no" })).toThrow();
    expect(documentProtection(model.doc)?.enforced).toBe(true);
    runCommand(model, unprotectDocumentCommand, { password: "pw" });
    expect(documentProtection(model.doc)?.enforced).toBe(false);
  });
});

describe("outline", () => {
  it("promotes, demotes and moves paragraphs", () => {
    const model = editorFor(createDocx({ paragraphs: ["A", "B", "C"] }));
    model.setSelection(caretAt({ block: 1 }));
    runCommand(model, promoteCommand, undefined);
    expect(outlineLevelOf(paragraphs(model.doc)[1]!)).toBe(1);
    runCommand(model, demoteCommand, undefined);
    expect(outlineLevelOf(paragraphs(model.doc)[1]!)).toBe(2);
    runCommand(model, moveBlocksCommand, { direction: -1 });
    expect(text(model.doc)).toBe("B\nA\nC");
    expect(model.selection?.focus.block).toBe(0);
    valid(model);
  });
});
