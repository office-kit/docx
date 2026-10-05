import {
  addImage,
  createDocx,
  type ChartSpec,
  imageDrawings,
  openDocx,
  readDrawing,
  setDrawingHidden,
  setDrawingHyperlink,
  setDrawingWrap,
  setPictureColorAdjustments,
  setPictureCrop,
  setPictureEffects,
  setPictureGeometry,
  setPictureOutline,
  toUint8Array,
  validate,
} from "@office-kit/docx";
import { describe, expect, it } from "vitest";
import { chartSvg } from "./render-chart.js";
import { renderDocumentHtml } from "./render.js";
import { drawingIndexAt, drawingPositionOf } from "./drawing-access.js";
import { caretAt } from "./selection.js";
import { editorFor } from "./index.js";
import { runCommand } from "./commands/types.js";
import { insertImageCommand } from "./commands/image.js";
import {
  insertChartCommand,
  drawingWrapCommand,
  drawingDeleteCommand,
  drawingLayoutCommand,
} from "./commands/drawing.js";

const PNG_1X1 = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
  0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0xfa, 0xcf, 0x00, 0x00,
  0x00, 0x02, 0x00, 0x01, 0xe5, 0x27, 0xde, 0xfc, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44,
  0xae, 0x42, 0x60, 0x82,
]);
const INCH = 914400;

const SPEC: ChartSpec = {
  kind: "column",
  grouping: "clustered",
  title: "Sales",
  legend: "right",
  dataLabels: true,
  categories: ["Q1", "Q2"],
  series: [{ name: "North", values: [3, 5] }],
};

describe("picture rendering", () => {
  it("draws an inline picture as SVG with its data URL and anchors", () => {
    const doc = createDocx({ paragraphs: [] });
    addImage(doc, PNG_1X1, {
      widthEmu: INCH,
      heightEmu: INCH / 2,
      name: "Logo",
      altText: "Our logo",
    });
    const html = renderDocumentHtml(doc);
    expect(html).toContain(`data-wk-object="picture"`);
    expect(html).toContain(`data-wk-drawing="0"`);
    expect(html).toContain(`aria-label="Our logo"`);
    expect(html).toContain("width:96px;height:48px");
    expect(html).toContain("data:image/png;base64,");
    // A picture-only paragraph keeps a caret placeholder and no editable run span.
    expect(html).not.toContain(`class="wk-run"`);
  });

  it("maps crop, shape, border, effects and adjustments onto SVG", () => {
    const doc = createDocx({ paragraphs: [] });
    addImage(doc, PNG_1X1, { widthEmu: INCH, heightEmu: INCH });
    setPictureCrop(doc, 0, { left: 50, top: 0, right: 0, bottom: 0 });
    setPictureGeometry(doc, 0, "ellipse");
    setPictureOutline(doc, 0, { color: "FF0000", widthEmu: 19050, dash: "dash", compound: "sng" });
    setPictureEffects(doc, 0, {
      shadow: {
        kind: "outer",
        color: "000000",
        opacity: 50,
        blurEmu: 38100,
        distanceEmu: 38100,
        directionDeg: 45,
      },
      reflection: { blurEmu: 0, startOpacity: 50, endOpacity: 0, endPosition: 50, distanceEmu: 0 },
    });
    setPictureColorAdjustments(doc, 0, {
      grayscale: true,
      transparency: 50,
      transparentColor: "FFFFFF",
    });
    const html = renderDocumentHtml(doc);
    // 50 % cropped from the left: the image is drawn twice as wide, shifted left.
    expect(html).toContain(`x="-96" y="0" width="192" height="96"`);
    expect(html).toContain(`<clipPath id="wkd0-c"><path d="M0,48A48,48`);
    expect(html).toContain(`stroke="#FF0000" stroke-width="2" stroke-dasharray="8 6"`);
    expect(html).toContain(`filter="url(#wkd0-f)"`);
    expect(html).toContain(`opacity="0.5"`);
    expect(html).toContain(`<mask id="wkd0-m"`);
    expect(html).toContain(`feComposite in="SourceGraphic" in2="keep" operator="in"`);
  });

  it("emits float data for anchored pictures and hides hidden ones", () => {
    const doc = createDocx({ paragraphs: ["Text"] });
    addImage(doc, PNG_1X1, { widthEmu: INCH, heightEmu: INCH });
    setDrawingWrap(doc, 0, "behindText", {
      horizontal: { relativeTo: "page", align: "center" },
      vertical: { relativeTo: "page", offsetEmu: INCH },
    });
    setDrawingHidden(doc, 0, true);
    setDrawingHyperlink(doc, 0, "https://example.com/");
    const html = renderDocumentHtml(doc);
    expect(html).toContain("wk-obj-float");
    expect(html).toContain(
      "&quot;relativeTo&quot;:&quot;page&quot;,&quot;align&quot;:&quot;center&quot;",
    );
    expect(html).toMatch(/data-wk-object="picture"[^>]* hidden /);
    expect(html).toMatch(/z-index:-\d+/);
    expect(html).toContain(`data-wk-href="https://example.com/"`);
  });
});

describe("chart rendering", () => {
  it("draws title, legend, bars and data labels", () => {
    const svg = chartSvg(SPEC, 400, 240);
    expect(svg).toContain(">Sales</text>");
    expect(svg).toContain(">North</text>");
    expect((svg.match(/<rect /g) ?? []).length).toBeGreaterThanOrEqual(4);
    expect(svg).toContain(">5</text>");
  });

  it.each<ChartSpec["kind"]>(["bar", "line", "area", "pie", "doughnut", "scatter", "radar"])(
    "draws a %s chart",
    (kind) => {
      const svg = chartSvg(
        { ...SPEC, kind, categories: kind === "scatter" ? ["1", "2"] : SPEC.categories },
        300,
        200,
      );
      expect(svg.startsWith("<svg")).toBe(true);
      expect(svg).not.toContain("NaN");
    },
  );
});

describe("picture and chart commands", () => {
  it("inserts a picture at the caret, splitting the run, and finds it again", () => {
    const model = editorFor(createDocx({ paragraphs: ["Hello world"] }));
    model.setSelection(caretAt({ block: 0, inline: 0, offset: 5 }));
    runCommand(model, insertImageCommand, {
      bytes: PNG_1X1,
      options: { widthEmu: INCH, heightEmu: INCH },
    });
    const para = model.doc.document.body.blocks[0];
    expect(
      para?.kind === "paragraph" &&
        para.children.map((c) => (c.kind === "run" ? c.pieces.map((p) => p.kind).join() : c.kind)),
    ).toEqual(["text", "drawing", "text"]);
    expect(model.selection?.focus).toMatchObject({ inline: 2, offset: 0 });
    expect(drawingIndexAt(model.doc, { block: 0, inline: 1 })).toBe(0);
    expect(drawingPositionOf(model.doc, 0)).toEqual({ block: 0, inline: 1 });
  });

  it("inserts a chart, floats it, deletes it, and stays valid", () => {
    const model = editorFor(createDocx({ paragraphs: [""] }));
    model.setSelection(caretAt({ block: 0, inline: 0, offset: 0 }));
    runCommand(model, insertChartCommand, { spec: SPEC });
    runCommand(model, drawingWrapCommand, { index: 0, wrap: "square" });
    const bytes = toUint8Array(model.doc);
    expect(validate(openDocx(bytes))).toEqual([]);
    expect(renderDocumentHtml(openDocx(bytes))).toContain(`data-wk-object="chart"`);
    runCommand(model, drawingDeleteCommand, { index: 0 });
    expect(imageDrawings(model.doc)).toHaveLength(0);
  });

  it("applies Layout Options in one undo step and stays valid", () => {
    const model = editorFor(createDocx({ paragraphs: ["Text"] }));
    model.setSelection(caretAt({ block: 0, inline: 0, offset: 0 }));
    runCommand(model, insertImageCommand, {
      bytes: PNG_1X1,
      options: { widthEmu: INCH, heightEmu: INCH },
    });
    runCommand(model, drawingLayoutCommand, {
      index: 0,
      wrap: "tight",
      position: {
        horizontal: { relativeTo: "page", align: "center" },
        vertical: { relativeTo: "paragraph", offsetEmu: 12700 },
      },
      options: { wrapSide: "largest", distance: { left: 0, right: 0 }, allowOverlap: false },
      size: { cxEmu: 2 * INCH, cyEmu: INCH },
      rotation: 30,
      lockAspect: false,
    });
    const doc = openDocx(toUint8Array(model.doc));
    expect(validate(doc)).toEqual([]);
    const drawing = imageDrawings(doc)[0];
    if (!drawing) throw new Error("drawing missing");
    const info = readDrawing(doc, drawing);
    expect(info).toMatchObject({
      kind: "picture",
      wrap: "tight",
      widthEmu: 2 * INCH,
      heightEmu: INCH,
      rotation: 30,
      lockAspect: false,
      anchor: {
        horizontal: { relativeTo: "page", align: "center" },
        vertical: { relativeTo: "paragraph", offsetEmu: 12700 },
        wrapSide: "largest",
        allowOverlap: false,
        distance: { left: 0, right: 0 },
      },
    });
    model.undo();
    expect(imageDrawings(model.doc).map((d) => readDrawing(model.doc, d).wrap)).toEqual(["inline"]);
  });
});
