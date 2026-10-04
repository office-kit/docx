/**
 * Color / font-size commands write only schema-valid values (ST_HexColor,
 * ST_HpsMeasure) and reject the rest before the document changes, keeping undo
 * and redo intact. Public surface only.
 */

import {
  createDocx,
  getRunFormat,
  openDocx,
  paragraphs,
  setRunFormat,
  setRunValProp,
  toUint8Array,
  type WmlRun,
} from "@office-kit/docx";
import { describe, expect, it } from "vitest";
import { commands, editorFor, type EditorModel, renderDocumentHtml, runCommand } from "./index.js";

function editorOnHello(): EditorModel {
  const model = editorFor(createDocx({ paragraphs: ["Hello"] }));
  model.setSelection({
    anchor: { block: 0, inline: 0, offset: 0 },
    focus: { block: 0, inline: 0, offset: 5 },
  });
  return model;
}

function firstRun(model: EditorModel): WmlRun {
  const run = paragraphs(model.doc)[0]?.children.find((c) => c.kind === "run");
  if (run?.kind !== "run") throw new Error("expected a run");
  return run;
}

describe("setColorCommand", () => {
  it.each([
    ["#ff0000", "ff0000"],
    ["00FF7F", "00FF7F"],
    ["auto", "auto"],
  ])("%s is saved as %s and reopens", (input, stored) => {
    const model = editorOnHello();
    runCommand(model, commands.setColorCommand, { color: input });
    const reopened = openDocx(toUint8Array(model.doc));
    const run = paragraphs(reopened)[0]?.children[0];
    expect(run?.kind === "run" && getRunFormat(run).color).toBe(stored);
  });
});

describe("setFontSizeCommand", () => {
  it.each([
    [0, 0],
    [11.25, 23], // rounded to the nearest half-point
    [1638, 3276],
  ])("%s pt is saved as %s half-points", (points, halfPoints) => {
    const model = editorOnHello();
    runCommand(model, commands.setFontSizeCommand, { points });
    expect(getRunFormat(firstRun(model)).fontSizeHalfPoints).toBe(halfPoints);
  });
});

/** Bold then undo, so a redo is pending; returns a check that nothing changed. */
function withPendingRedo(model: EditorModel): () => void {
  runCommand(model, commands.toggleBoldCommand, undefined);
  model.undo();
  const body = JSON.stringify(model.doc.document.body);
  const selection = model.selection;
  return () => {
    expect(JSON.stringify(model.doc.document.body)).toBe(body);
    expect(firstRun(model).rPr).toBeUndefined();
    expect(model.selection).toEqual(selection);
    expect(model.canUndo()).toBe(false);
    expect(model.canRedo()).toBe(true);
    model.redo();
    expect(getRunFormat(firstRun(model)).bold).toBe(true);
  };
}

describe.each(["red", "#FFF", "FF00001", ""])("setColorCommand rejects %j atomically", (color) => {
  it("throws before changing anything and keeps undo / redo", () => {
    const model = editorOnHello();
    const unchanged = withPendingRedo(model);
    expect(() => runCommand(model, commands.setColorCommand, { color })).toThrow(RangeError);
    unchanged();
  });
});

describe("wrong runtime types from untyped JS", () => {
  it("setColorCommand rejects a number instead of stringifying it", () => {
    const model = editorOnHello();
    const unchanged = withPendingRedo(model);
    const params: { color: string } = JSON.parse('{"color":123456}');
    expect(() => runCommand(model, commands.setColorCommand, params)).toThrow(RangeError);
    unchanged();
  });

  it.each(['{"points":"12"}', '{"points":null}', '{"points":true}', '{"points":[12]}'])(
    "setFontSizeCommand rejects %s instead of coercing it",
    (json) => {
      const model = editorOnHello();
      const unchanged = withPendingRedo(model);
      const params: { points: number } = JSON.parse(json);
      expect(() => runCommand(model, commands.setFontSizeCommand, params)).toThrow(RangeError);
      unchanged();
    },
  );
});

describe.each([-1, -0.2, -0.1, Number.NaN, Number.POSITIVE_INFINITY])(
  "setFontSizeCommand rejects %s pt atomically",
  (points) => {
    it("throws before changing anything and keeps undo / redo", () => {
      const model = editorOnHello();
      const unchanged = withPendingRedo(model);
      expect(() => runCommand(model, commands.setFontSizeCommand, { points })).toThrow(RangeError);
      unchanged();
    });
  },
);

describe("rendering color / font values from an untrusted file", () => {
  it("cannot break out of the style attribute", () => {
    const doc = createDocx({ paragraphs: ["Hello"] });
    const run = paragraphs(doc)[0]!.children[0]!;
    if (run.kind !== "run") throw new Error("expected a run");
    // Written raw, as a hostile file would carry them; the typed writers
    // would reject the color.
    setRunValProp(run, "color", 'x" onmouseover="alert(1)');
    setRunFormat(run, { font: 'Evil" onclick="alert(2)\\' });
    const html = renderDocumentHtml(doc);
    expect(html).not.toMatch(/onmouseover="|onclick="/);
    expect(html).not.toContain("color:#x");
    expect(html).toContain("font-family:'Evil onclick=alert(2)'");
  });

  it("renders a valid hex color and skips auto", () => {
    const model = editorOnHello();
    runCommand(model, commands.setColorCommand, { color: "C00000" });
    expect(renderDocumentHtml(model.doc)).toContain("color:#C00000");
    runCommand(model, commands.setColorCommand, { color: "auto" });
    expect(renderDocumentHtml(model.doc)).not.toContain("color:#");
  });
});
