/**
 * `<w:highlight w:val>` is ST_HighlightColor (ECMA-376 Part 1 §17.18.40): a
 * closed list of names, never RGB. These tests check the saved XML itself, not
 * only the reparsed model. Public surface only.
 */

import {
  createDocx,
  getRunFormat,
  HIGHLIGHT_COLORS,
  type HighlightColor,
  openDocx,
  paragraphs,
  setRunValProp,
  toUint8Array,
  validate,
  type Docx,
  type WmlRun,
} from "@office-kit/docx";
import { strFromU8, unzipSync } from "fflate";
import { describe, expect, it } from "vitest";
import {
  commands,
  editorFor,
  type EditorModel,
  openEditor,
  renderDocumentHtml,
  runCommand,
} from "./index.js";

// Copied from ST_HighlightColor in the transitional WML schema (wml.xsd), in
// schema order; the same 17 values as the Open XML SDK's HighlightColorValues.
const SCHEMA_VALUES = [
  "black",
  "blue",
  "cyan",
  "green",
  "magenta",
  "red",
  "yellow",
  "white",
  "darkBlue",
  "darkCyan",
  "darkGreen",
  "darkMagenta",
  "darkRed",
  "darkYellow",
  "darkGray",
  "lightGray",
  "none",
];

function selectHello(model: EditorModel): void {
  model.setSelection({
    anchor: { block: 0, inline: 0, offset: 0 },
    focus: { block: 0, inline: 0, offset: 5 },
  });
}

function editorOnHello(): EditorModel {
  const model = editorFor(createDocx({ paragraphs: ["Hello"] }));
  selectHello(model);
  return model;
}

function documentXml(bytes: Uint8Array): string {
  const part = unzipSync(bytes)["word/document.xml"];
  if (!part) throw new Error("word/document.xml missing");
  return strFromU8(part);
}

function firstRun(doc: Docx): WmlRun {
  const run = paragraphs(doc)[0]?.children.find((c) => c.kind === "run");
  if (run?.kind !== "run") throw new Error("expected a run");
  return run;
}

function highlightTags(xml: string): string[] {
  return xml.match(/<w:highlight\b[^>]*>/g) ?? [];
}

describe("setHighlightCommand", () => {
  it("offers exactly the ST_HighlightColor values", () => {
    expect([...HIGHLIGHT_COLORS].toSorted()).toEqual([...SCHEMA_VALUES].toSorted());
  });

  it.each(SCHEMA_VALUES as HighlightColor[])(
    "writes w:val=%s, valid after save → reopen → save",
    (color) => {
      const model = editorOnHello();
      runCommand(model, commands.setHighlightCommand, { color });
      const saved = toUint8Array(model.doc);
      expect(highlightTags(documentXml(saved))).toEqual([`<w:highlight w:val="${color}"/>`]);

      const reopened = openDocx(saved);
      expect(getRunFormat(firstRun(reopened)).highlight).toBe(color);
      expect(validate(reopened)).toHaveLength(0);
      expect(highlightTags(documentXml(toUint8Array(reopened)))).toEqual([
        `<w:highlight w:val="${color}"/>`,
      ]);
    },
  );

  it.each(["#FFFF00", "FFFF00", "ffff00", "Yellow", "transparent", "", " none"])(
    "rejects %j before changing anything, keeping undo and redo",
    (bad) => {
      const model = editorOnHello();
      runCommand(model, commands.setHighlightCommand, { color: "yellow" });
      model.undo();
      const before = documentXml(toUint8Array(model.doc));
      const selection = model.selection;

      expect(() =>
        runCommand(model, commands.setHighlightCommand, { color: bad as HighlightColor }),
      ).toThrow(RangeError);

      expect(documentXml(toUint8Array(model.doc))).toBe(before);
      expect(highlightTags(before)).toEqual([]);
      expect(model.selection).toEqual(selection);
      expect(model.canUndo()).toBe(false);
      expect(model.canRedo()).toBe(true);
      model.redo();
      expect(getRunFormat(firstRun(model.doc)).highlight).toBe("yellow");
    },
  );

  it("renders named colors only; none and foreign hex get no background", () => {
    const model = editorOnHello();
    runCommand(model, commands.setHighlightCommand, { color: "darkYellow" });
    expect(renderDocumentHtml(model.doc)).toContain("background-color:#808000");
    runCommand(model, commands.setHighlightCommand, { color: "none" });
    expect(renderDocumentHtml(model.doc)).not.toContain("background-color");
    // A file written by another tool may carry an invalid hex value; it is
    // kept as-is in the document but not guessed into a color.
    const foreign = createDocx({ paragraphs: ["Hello"] });
    setRunValProp(firstRun(foreign), "highlight", "FFFF00");
    expect(renderDocumentHtml(foreign)).not.toContain("background-color");
  });

  it("is one undo step; redo re-applies it", () => {
    const model = editorOnHello();
    runCommand(model, commands.setHighlightCommand, { color: "darkBlue" });
    runCommand(model, commands.setHighlightCommand, { color: "none" });
    model.undo();
    expect(getRunFormat(firstRun(model.doc)).highlight).toBe("darkBlue");
    model.undo();
    expect(getRunFormat(firstRun(model.doc)).highlight).toBeUndefined();
    model.redo();
    model.redo();
    expect(getRunFormat(firstRun(model.doc)).highlight).toBe("none");
  });

  it("the same open → highlight → save script run twice gives the same XML", () => {
    const script = (bytes: Uint8Array): Uint8Array => {
      const model = openEditor(bytes);
      selectHello(model);
      runCommand(model, commands.setHighlightCommand, { color: "lightGray" });
      return toUint8Array(model.doc);
    };
    const once = script(toUint8Array(createDocx({ paragraphs: ["Hello"] })));
    const twice = script(once);
    expect(documentXml(twice)).toBe(documentXml(once));
    expect(highlightTags(documentXml(twice))).toEqual(['<w:highlight w:val="lightGray"/>']);
  });
});
