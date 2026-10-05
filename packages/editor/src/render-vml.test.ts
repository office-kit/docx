import {
  createDocx,
  getShapeFill,
  getShapeLayout,
  getShapeWrap,
  getSmartArt,
  shapeText,
  smartArts,
  toUint8Array,
  validate,
  vmlShapes,
} from "@office-kit/docx";
import { describe, expect, it } from "vitest";
import { runCommand } from "./commands/types.js";
import {
  changeShapeCommand,
  deleteShapeCommand,
  groupShapesCommand,
  insertInkCommand,
  insertShapeCommand,
  insertSmartArtCommand,
  insertWordArtCommand,
  shapeFillCommand,
  shapeLayoutCommand,
  shapeTextRunCommand,
  shapeWrapCommand,
  smartArtNodesCommand,
  ungroupShapesCommand,
} from "./commands/shapes.js";
import { editorFor } from "./index.js";
import { renderDocumentHtml } from "./render.js";
import { caretAt } from "./selection.js";
import { vmlPathToSvg } from "./vml-path.js";

function model() {
  const m = editorFor(createDocx({ paragraphs: ["Hello", "World"] }));
  m.setSelection(caretAt({ block: 0, inline: 0, offset: 0 }));
  return m;
}

describe("vmlPathToSvg", () => {
  it("reads empty parameter slots as zero and maps to the box", () => {
    const [sub] = vmlPathToSvg("m,l,21600r21600,xe", {
      coordSize: [21600, 21600],
      width: 100,
      height: 50,
    });
    expect(sub?.d).toBe("M0 0 L0 50 L100 50 Z");
  });

  it("evaluates shape type formulas and adjust values", () => {
    const [sub] = vmlPathToSvg("m@0,l,21600r21600,xe", {
      coordSize: [21600, 21600],
      width: 216,
      height: 216,
      adj: [5400],
      formulas: ["val #0"],
    });
    expect(sub?.d.startsWith("M54 0")).toBe(true);
  });

  it("splits unfilled subpaths", () => {
    const subs = vmlPathToSvg("m0,0l10,10xnfm0,0l5,5e", {
      coordSize: [10, 10],
      width: 10,
      height: 10,
    });
    expect(subs.map((s) => s.fill)).toEqual([true, false]);
  });
});

describe("shape commands and rendering", () => {
  it("inserts a shape at the caret paragraph and renders it selectable", () => {
    const m = model();
    const at = runCommand(m, insertShapeCommand, {
      options: { preset: "rightArrow", left: 10, top: 20, width: 100, height: 50 },
    });
    expect(at).toEqual({ block: 0, inline: 1, offset: 0 });
    const html = renderDocumentHtml(m.doc);
    expect(html).toContain('data-wk-object="shape"');
    expect(html).toContain("<svg");
    expect(html).toContain("left:calc(var(--m-left) + 10pt)");
    // Floats render before the paragraph's text.
    expect(html.indexOf("wk-shape")).toBeLessThan(html.indexOf(">Hello<"));
    expect(validate(m.doc)).toEqual([]);
  });

  it("edits fill, layout, wrap and preset, then deletes", () => {
    const m = model();
    const at = runCommand(m, insertShapeCommand, {
      options: { preset: "rect", width: 10, height: 10 },
    });
    if (!at) throw new Error("not inserted");
    runCommand(m, shapeFillCommand, { at, fill: { type: "solid", color: "FF0000" } });
    runCommand(m, shapeLayoutCommand, { at, layout: { width: 30, rotation: 45 } });
    runCommand(m, shapeWrapCommand, { at, wrap: "square" });
    runCommand(m, changeShapeCommand, { at, preset: "ellipse" });
    const [shape] = vmlShapes(m.doc);
    if (!shape) throw new Error("missing");
    expect(getShapeFill(shape)).toEqual({ type: "solid", color: "FF0000" });
    expect(getShapeLayout(shape)).toMatchObject({ width: 30, rotation: 45 });
    expect(getShapeWrap(shape)).toBe("square");
    expect(renderDocumentHtml(m.doc)).toContain("float:left");
    runCommand(m, deleteShapeCommand, { at });
    expect(vmlShapes(m.doc)).toHaveLength(0);
  });

  it("renders editable text boxes and syncs typed text", () => {
    const m = model();
    const at = runCommand(m, insertShapeCommand, {
      options: { preset: "textBox", width: 100, height: 40, text: "Old" },
    });
    if (!at) throw new Error("not inserted");
    const html = renderDocumentHtml(m.doc);
    expect(html).toContain('data-wk-object="textBox"');
    expect(html).toContain('contenteditable="true"');
    expect(html).toContain('data-wk-txbx-run="0,0"');
    runCommand(m, shapeTextRunCommand, { at, para: 0, run: 0, text: "New" });
    const [shape] = vmlShapes(m.doc);
    if (!shape) throw new Error("missing");
    const run = shapeText(shape)[0]?.children[0];
    expect(run?.kind === "run" && run.pieces[0]).toMatchObject({ value: "New" });
  });

  it("renders WordArt as stretched SVG text and ink as a stroke", () => {
    const m = model();
    runCommand(m, insertWordArtCommand, { options: { text: "Big", width: 200, height: 50 } });
    runCommand(m, insertInkCommand, {
      options: {
        width: 20,
        height: 20,
        points: [
          [0, 0],
          [20, 20],
        ],
        stroke: { color: "FF0000", weight: 2 },
      },
    });
    const html = renderDocumentHtml(m.doc);
    expect(html).toContain('lengthAdjust="spacingAndGlyphs"');
    expect(html).toContain('data-wk-object="ink"');
  });

  it("groups and ungroups through commands", () => {
    const m = model();
    const a = runCommand(m, insertShapeCommand, {
      options: { preset: "rect", left: 0, top: 0, width: 10, height: 10 },
    });
    const b = runCommand(m, insertShapeCommand, {
      options: { preset: "rect", left: 20, top: 0, width: 10, height: 10 },
    });
    if (!a || !b) throw new Error("not inserted");
    const group = runCommand(m, groupShapesCommand, { ats: [a, b] });
    if (!group) throw new Error("not grouped");
    expect(vmlShapes(m.doc)).toHaveLength(1);
    expect(renderDocumentHtml(m.doc)).toContain("wk-shape-part");
    runCommand(m, ungroupShapesCommand, { at: group });
    expect(vmlShapes(m.doc)).toHaveLength(2);
    expect(toUint8Array(m.doc).length).toBeGreaterThan(0);
  });

  it("inserts, renders and edits SmartArt", () => {
    const m = model();
    const at = runCommand(m, insertSmartArtCommand, {
      options: { layout: "basicProcess", nodes: [{ text: "A" }, { text: "B" }] },
    });
    if (!at) throw new Error("not inserted");
    expect(renderDocumentHtml(m.doc)).toContain('data-wk-object="smartArt"');
    runCommand(m, smartArtNodesCommand, { at, nodes: [{ text: "X" }] });
    const [ref] = smartArts(m.doc);
    if (!ref) throw new Error("missing");
    expect(getSmartArt(m.doc, ref).nodes).toEqual([{ text: "X" }]);
    expect(validate(m.doc)).toEqual([]);
  });
});
