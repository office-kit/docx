import { createDocx, protectDocument } from "@office-kit/docx";
import { describe, expect, it } from "vitest";
import { toggleBoldCommand } from "./commands/text.js";
import {
  acceptAllRevisionsCommand,
  addCommentCommand,
  unprotectDocumentCommand,
} from "./commands/review.js";
import { insertTextCommand } from "./commands/structure.js";
import { EditorModel } from "./model.js";
import { isEditingLocked, PROTECTED_MESSAGE, protectionRefusal } from "./protection-guard.js";

function protectedModel(edit: Parameters<typeof protectDocument>[1]): EditorModel {
  const doc = createDocx({ paragraphs: ["Hello"] });
  protectDocument(doc, edit);
  return new EditorModel(doc);
}

describe("protectionRefusal", () => {
  it("allows everything without enforced protection", () => {
    const model = new EditorModel(createDocx({ paragraphs: ["Hello"] }));
    expect(protectionRefusal(model, insertTextCommand)).toBeUndefined();
    expect(isEditingLocked(model)).toBe(false);
  });

  it("read only refuses edits but still lets protection be stopped", () => {
    const model = protectedModel({ edit: "readOnly" });
    expect(protectionRefusal(model, insertTextCommand)).toBe(PROTECTED_MESSAGE);
    expect(protectionRefusal(model, addCommentCommand)).toBe(PROTECTED_MESSAGE);
    expect(protectionRefusal(model, unprotectDocumentCommand)).toBeUndefined();
    expect(isEditingLocked(model)).toBe(true);
  });

  it("comments only allows commenting", () => {
    const model = protectedModel({ edit: "comments" });
    expect(protectionRefusal(model, addCommentCommand)).toBeUndefined();
    expect(protectionRefusal(model, insertTextCommand)).toBe(PROTECTED_MESSAGE);
  });

  it("tracked changes allows editing but not accepting changes", () => {
    const model = protectedModel({ edit: "trackedChanges" });
    expect(protectionRefusal(model, insertTextCommand)).toBeUndefined();
    expect(protectionRefusal(model, acceptAllRevisionsCommand)).toBe(PROTECTED_MESSAGE);
    expect(isEditingLocked(model)).toBe(false);
  });

  it("limited formatting refuses direct formatting", () => {
    const model = protectedModel({ edit: "none", formatting: true });
    expect(protectionRefusal(model, toggleBoldCommand)).toBe(PROTECTED_MESSAGE);
    expect(protectionRefusal(model, insertTextCommand)).toBeUndefined();
  });
});
