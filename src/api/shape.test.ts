import { describe, expect, it } from "vitest";
import { getPart, writeOpcPackage } from "../internal/opc/index.js";
import { serializeXml, type XmlElement } from "../internal/xml/index.js";
import { createDocx, type Docx, openDocx, paragraphs, toUint8Array } from "./docx.js";
import {
  addDrawingCanvas,
  addShape,
  addShapeToGroup,
  addWordArt,
  changeShapePreset,
  getShapeAltText,
  getShapeFill,
  getShapeLayout,
  getShapePoints,
  getShapeShadow,
  getShapeStroke,
  getShapeWrap,
  getTextBoxLayout,
  getWordArt,
  groupMembers,
  groupShapes,
  linkedTextBox,
  linkTextBoxes,
  removeShape,
  runShape,
  SHAPE_PRESETS,
  type ShapePreset,
  setShapeAltText,
  setShapeFill,
  setShapeLayout,
  setShapeOrder,
  setShapePoints,
  setShapeShadow,
  setShapeStroke,
  setShapeText,
  setShapeWrap,
  setTextBoxLayout,
  setWordArt,
  shapeKind,
  shapePreset,
  shapeText,
  ungroupShapes,
  vmlShapes,
} from "./shape.js";
import { validatePackage } from "./validator.js";

const TINY_PNG = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
  0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0xfa, 0xcf, 0x00, 0x00,
  0x00, 0x02, 0x00, 0x01, 0xe5, 0x27, 0xde, 0xfc, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44,
  0xae, 0x42, 0x60, 0x82,
]);

function firstParagraph(doc: Docx) {
  const p = paragraphs(doc)[0];
  if (!p) throw new Error("no paragraph");
  return p;
}

/** Save, reopen, and return the reopened document plus its document.xml text. */
function roundTrip(doc: Docx): { doc: Docx; xml: string } {
  const bytes = toUint8Array(doc);
  const reopened = openDocx(bytes);
  expect(validatePackage(reopened.opc)).toEqual([]);
  const part = getPart(reopened.opc, "/word/document.xml");
  if (!part) throw new Error("no document part");
  return { doc: reopened, xml: new TextDecoder().decode(part.data) };
}

const spid = (el: XmlElement) => el.attrs.find((x) => x.name.local === "spid")?.value;

function only(doc: Docx): XmlElement {
  const [shape] = vmlShapes(doc);
  if (!shape) throw new Error("no shape");
  return shape;
}

describe("addShape", () => {
  it("writes a floating VML shape that round-trips", () => {
    const doc = createDocx({ paragraphs: ["Anchor"] });
    addShape(doc, firstParagraph(doc), {
      preset: "rightArrow",
      left: 36,
      top: 18,
      width: 144,
      height: 72,
      altText: "An arrow",
    });
    const { doc: back, xml } = roundTrip(doc);
    expect(xml).toContain("<w:pict");
    expect(xml).toContain('xmlns:v="urn:schemas-microsoft-com:vml"');
    expect(xml).toContain('o:spt="13"');
    expect(xml).toContain('alt="An arrow"');
    const shape = only(back);
    expect(shapePreset(shape)).toBe("rightArrow");
    expect(shapeKind(shape)).toBe("shape");
    expect(getShapeAltText(shape)).toBe("An arrow");
    const layout = getShapeLayout(shape);
    expect(layout).toMatchObject({ inline: false, left: 36, top: 18, width: 144, height: 72 });
    expect(layout.zIndex).toBeGreaterThan(0);
    expect(getShapeWrap(shape)).toBe("inFrontOfText");
    expect(getShapeFill(shape)).toEqual({ type: "solid", color: "4472C4" });
    expect(getShapeStroke(shape)).toMatchObject({ color: "2F528F", weight: 1 });
  });

  it("writes every gallery preset as valid VML", () => {
    const doc = createDocx({ paragraphs: ["Anchor"] });
    const presets = Object.keys(SHAPE_PRESETS) as ShapePreset[];
    for (const preset of presets)
      addShape(doc, firstParagraph(doc), { preset, width: 50, height: 50 });
    const { doc: back } = roundTrip(doc);
    const shapes = vmlShapes(back);
    expect(shapes).toHaveLength(presets.length);
    // Presets that share a VML shape type read back as the first of them.
    for (const shape of shapes) expect(shapePreset(shape)).toBeDefined();
    expect(shapes.map((s) => s.name.local)).toContain("oval");
    expect(shapes.map((s) => s.name.local)).toContain("roundrect");
  });

  it("allocates unique shape ids", () => {
    const doc = createDocx({ paragraphs: ["Anchor"] });
    const a = addShape(doc, firstParagraph(doc), { preset: "rect", width: 10, height: 10 });
    const b = addShape(doc, firstParagraph(doc), { preset: "rect", width: 10, height: 10 });
    expect(spid(a)).toBe("_x0000_s1025");
    expect(spid(b)).toBe("_x0000_s1026");
  });

  it("writes an inline shape without a position", () => {
    const doc = createDocx({ paragraphs: ["Anchor"] });
    addShape(doc, firstParagraph(doc), {
      preset: "ellipse",
      width: 20,
      height: 20,
      wrap: "inline",
    });
    const shape = only(roundTrip(doc).doc);
    expect(getShapeLayout(shape).inline).toBe(true);
    expect(getShapeWrap(shape)).toBe("inline");
  });
});

describe("fill, outline, shadow", () => {
  it("round-trips solid, gradient, pattern and picture fills", () => {
    const doc = createDocx({ paragraphs: ["Anchor"] });
    const p = firstParagraph(doc);
    const solid = addShape(doc, p, { preset: "rect", width: 10, height: 10 });
    setShapeFill(doc, solid, { type: "solid", color: "FF0000", opacity: 0.5 });
    const gradient = addShape(doc, p, { preset: "rect", width: 10, height: 10 });
    setShapeFill(doc, gradient, { type: "gradient", color: "FF0000", color2: "0000FF", angle: 90 });
    const pattern = addShape(doc, p, { preset: "rect", width: 10, height: 10 });
    setShapeFill(doc, pattern, {
      type: "pattern",
      pattern: "ltDnDiag",
      color: "000000",
      color2: "FFFFFF",
    });
    const picture = addShape(doc, p, { preset: "rect", width: 10, height: 10 });
    setShapeFill(doc, picture, { type: "picture", bytes: TINY_PNG, contentType: "image/png" });
    const none = addShape(doc, p, { preset: "rect", width: 10, height: 10 });
    setShapeFill(doc, none, { type: "none" });
    const [s, g, pt, pic, n] = vmlShapes(roundTrip(doc).doc);
    if (!s || !g || !pt || !pic || !n) throw new Error("missing shapes");
    expect(getShapeFill(s)).toEqual({ type: "solid", color: "FF0000", opacity: 0.5 });
    expect(getShapeFill(g)).toEqual({
      type: "gradient",
      color: "FF0000",
      color2: "0000FF",
      style: "linear",
      angle: 90,
    });
    expect(getShapeFill(pt)).toEqual({
      type: "pattern",
      pattern: "ltDnDiag",
      color: "000000",
      color2: "FFFFFF",
    });
    expect(getShapeFill(pic)).toMatchObject({ type: "picture", tile: false });
    expect(getShapeFill(n)).toEqual({ type: "none" });
  });

  it("round-trips outline dashes, arrows and shadows", () => {
    const doc = createDocx({ paragraphs: ["Anchor"] });
    const shape = addShape(doc, firstParagraph(doc), { preset: "line", width: 100, height: 0 });
    setShapeStroke(shape, { color: "00FF00", weight: 3, dash: "dash", endArrow: "classic" });
    setShapeShadow(shape, { color: "808080", offsetX: 3, offsetY: 3, opacity: 0.4 });
    const back = only(roundTrip(doc).doc);
    expect(getShapeStroke(back)).toEqual({
      color: "00FF00",
      weight: 3,
      dash: "dash",
      startArrow: "none",
      endArrow: "classic",
    });
    expect(getShapeShadow(back)).toEqual({ color: "808080", offsetX: 3, offsetY: 3, opacity: 0.4 });
    setShapeStroke(back, null);
    setShapeShadow(back, null);
    expect(getShapeStroke(back)).toBeNull();
    expect(getShapeShadow(back)).toBeNull();
  });
});

describe("layout, wrap and order", () => {
  it("updates position, size, rotation and flip", () => {
    const doc = createDocx({ paragraphs: ["Anchor"] });
    const shape = addShape(doc, firstParagraph(doc), { preset: "triangle", width: 10, height: 10 });
    setShapeLayout(shape, {
      left: 5,
      top: 6,
      width: 70,
      height: 80,
      rotation: 450,
      flipH: true,
      horizontalRelativeTo: "page",
      verticalRelativeTo: "margin",
    });
    const back = only(roundTrip(doc).doc);
    expect(getShapeLayout(back)).toMatchObject({
      left: 5,
      top: 6,
      width: 70,
      height: 80,
      rotation: 90,
      flipH: true,
      flipV: false,
      horizontalRelativeTo: "page",
      verticalRelativeTo: "margin",
    });
    setShapeLayout(back, { horizontalAlign: "center" });
    expect(getShapeLayout(back).horizontalAlign).toBe("center");
  });

  it("writes w10:wrap and moves behind text to the negative layer", () => {
    const doc = createDocx({ paragraphs: ["Anchor"] });
    const shape = addShape(doc, firstParagraph(doc), { preset: "rect", width: 10, height: 10 });
    setShapeWrap(doc, shape, "square");
    expect(serializeXml({ root: shape, prologue: [], epilogue: [] })).toContain(
      '<w10:wrap type="square"/>',
    );
    setShapeWrap(doc, shape, "behindText");
    const back = only(roundTrip(doc).doc);
    expect(getShapeWrap(back)).toBe("behindText");
    expect(getShapeLayout(back).zIndex).toBeLessThan(0);
    setShapeWrap(doc, back, "inline");
    expect(getShapeWrap(back)).toBe("inline");
  });

  it("brings shapes forward and sends them back", () => {
    const doc = createDocx({ paragraphs: ["Anchor"] });
    const p = firstParagraph(doc);
    const a = addShape(doc, p, { preset: "rect", width: 10, height: 10 });
    const b = addShape(doc, p, { preset: "rect", width: 10, height: 10 });
    const z = (s: XmlElement) => getShapeLayout(s).zIndex;
    expect(z(b)).toBeGreaterThan(z(a));
    setShapeOrder(doc, a, "bringToFront");
    expect(z(a)).toBeGreaterThan(z(b));
    setShapeOrder(doc, a, "sendBackward");
    expect(z(a)).toBeLessThan(z(b));
    setShapeOrder(doc, b, "sendBehindText");
    expect(z(b)).toBeLessThan(0);
  });
});

describe("text boxes", () => {
  it("writes w:txbxContent and round-trips its paragraphs", () => {
    const doc = createDocx({ paragraphs: ["Anchor"] });
    const box = addShape(doc, firstParagraph(doc), {
      preset: "textBox",
      width: 100,
      height: 50,
      text: "Hello\nWorld",
      textLayout: { direction: "vertical", anchor: "middle" },
    });
    expect(shapeKind(box)).toBe("textBox");
    const { doc: back, xml } = roundTrip(doc);
    expect(xml).toContain("<w:txbxContent>");
    expect(xml).toContain("layout-flow:vertical");
    const shape = only(back);
    expect(shapeText(shape).map((p) => p.children.length)).toEqual([1, 1]);
    expect(getTextBoxLayout(shape)).toMatchObject({ direction: "vertical", anchor: "middle" });
    setTextBoxLayout(shape, {
      direction: "vertical270",
      anchor: "bottom",
      inset: [1, 2, 3, 4],
      autoFit: true,
    });
    expect(getTextBoxLayout(shape)).toEqual({
      direction: "vertical270",
      anchor: "bottom",
      inset: [1, 2, 3, 4],
      autoFit: true,
    });
    setShapeText(shape, "One");
    expect(shapeText(shape)).toHaveLength(1);
  });

  it("links one text box to the next", () => {
    const doc = createDocx({ paragraphs: ["Anchor"] });
    const p = firstParagraph(doc);
    const a = addShape(doc, p, { preset: "textBox", width: 10, height: 10 });
    const b = addShape(doc, p, { preset: "textBox", width: 10, height: 10, name: "Text Box 9" });
    linkTextBoxes(a, b);
    const [ra] = vmlShapes(roundTrip(doc).doc);
    if (!ra) throw new Error("missing");
    expect(linkedTextBox(ra)).toBe("Text Box 9");
    linkTextBoxes(ra, undefined);
    expect(linkedTextBox(ra)).toBeUndefined();
  });
});

describe("WordArt", () => {
  it("writes a v:textpath shape and edits its text", () => {
    const doc = createDocx({ paragraphs: ["Anchor"] });
    addWordArt(doc, firstParagraph(doc), {
      text: "Title",
      width: 200,
      height: 60,
      bold: true,
      font: "Georgia",
    });
    const { doc: back, xml } = roundTrip(doc);
    expect(xml).toContain("<v:textpath");
    expect(xml).toContain('string="Title"');
    // Word draws nothing for a text path without its shape type definition.
    expect(xml).toMatch(/<v:shapetype id="_x0000_t136"[^>]*o:spt="136"/);
    expect(xml).toContain('type="#_x0000_t136"');
    const shape = only(back);
    expect(shapeKind(shape)).toBe("wordArt");
    expect(getWordArt(shape)).toEqual({
      text: "Title",
      font: "Georgia",
      size: 36,
      bold: true,
      italic: false,
    });
    setWordArt(shape, { text: "New", italic: true, bold: false });
    expect(getWordArt(shape)).toMatchObject({ text: "New", italic: true, bold: false });
  });
});

describe("freeforms and ink", () => {
  it("writes ink strokes as named, unfilled freeforms", () => {
    const doc = createDocx({ paragraphs: ["Anchor"] });
    const ink = addShape(doc, firstParagraph(doc), {
      preset: "scribble",
      ink: true,
      width: 100,
      height: 50,
      points: [
        [0, 0],
        [50, 50],
        [100, 0],
      ],
      stroke: { color: "FFFF00", weight: 6, opacity: 0.5 },
    });
    expect(shapeKind(ink)).toBe("ink");
    const back = only(roundTrip(doc).doc);
    expect(getShapeFill(back)).toEqual({ type: "none" });
    expect(getShapeStroke(back)?.opacity).toBe(0.5);
    expect(getShapePoints(back)).toEqual([
      [0, 0],
      [50, 50],
      [100, 0],
    ]);
    setShapePoints(back, [
      [0, 0],
      [100, 50],
    ]);
    expect(getShapePoints(back)).toEqual([
      [0, 0],
      [100, 50],
    ]);
  });
});

describe("change shape, delete", () => {
  it("swaps the geometry and keeps the box", () => {
    const doc = createDocx({ paragraphs: ["Anchor"] });
    const shape = addShape(doc, firstParagraph(doc), { preset: "rect", width: 30, height: 40 });
    changeShapePreset(shape, "star5");
    expect(shape.name.local).toBe("shape");
    expect(shapePreset(shape)).toBe("star5");
    changeShapePreset(shape, "ellipse");
    expect(shape.name.local).toBe("oval");
    expect(getShapeLayout(shape)).toMatchObject({ width: 30, height: 40 });
    setShapeAltText(shape, "");
    expect(getShapeAltText(shape)).toBe("");
  });

  it("removes a shape and its run", () => {
    const doc = createDocx({ paragraphs: ["Anchor"] });
    const p = firstParagraph(doc);
    const before = p.children.length;
    const shape = addShape(doc, p, { preset: "rect", width: 1, height: 1 });
    expect(removeShape(doc, shape)).toBe(true);
    expect(p.children.length).toBe(before);
  });
});

describe("groups and drawing canvases", () => {
  it("groups and ungroups floating shapes", () => {
    const doc = createDocx({ paragraphs: ["Anchor"] });
    const p = firstParagraph(doc);
    const a = addShape(doc, p, { preset: "rect", left: 10, top: 10, width: 20, height: 20 });
    const b = addShape(doc, p, { preset: "ellipse", left: 50, top: 30, width: 20, height: 20 });
    const group = groupShapes(doc, [a, b]);
    expect(shapeKind(group)).toBe("group");
    expect(getShapeLayout(group)).toMatchObject({ left: 10, top: 10, width: 60, height: 40 });
    const { doc: back, xml } = roundTrip(doc);
    expect(xml).toContain("<v:group");
    const reGroup = only(back);
    expect(groupMembers(reGroup)).toHaveLength(2);
    const members = ungroupShapes(back, reGroup);
    expect(members).toHaveLength(2);
    expect(vmlShapes(back)).toHaveLength(2);
    const [ra, rb] = vmlShapes(back);
    if (!ra || !rb) throw new Error("missing");
    expect(getShapeLayout(ra)).toMatchObject({ left: 10, top: 10, width: 20, height: 20 });
    expect(getShapeLayout(rb)).toMatchObject({ left: 50, top: 30, width: 20, height: 20 });
  });

  it("draws into a drawing canvas", () => {
    const doc = createDocx({ paragraphs: ["Anchor"] });
    const canvas = addDrawingCanvas(doc, firstParagraph(doc), { width: 300, height: 200 });
    addShapeToGroup(doc, canvas, {
      preset: "scribble",
      ink: true,
      left: 10,
      top: 10,
      width: 50,
      height: 50,
      points: [
        [0, 0],
        [50, 50],
      ],
    });
    const { doc: back, xml } = roundTrip(doc);
    expect(xml).toContain('editas="canvas"');
    const shape = only(back);
    expect(shapeKind(shape)).toBe("canvas");
    // Background plus the stroke.
    expect(groupMembers(shape)).toHaveLength(2);
  });
});

describe("reading existing VML", () => {
  it("reads Word-style shapes that reference a shape type", () => {
    const xml =
      `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">` +
      `<w:body><w:p><w:r><w:pict><v:shapetype id="_x0000_t5" coordsize="21600,21600" o:spt="5" path="m@0,l,21600r21600,xe"/>` +
      `<v:shape id="Isosceles Triangle 1" o:spid="_x0000_s1026" type="#_x0000_t5" style="position:absolute;margin-left:1in;margin-top:12pt;width:100px;height:2cm;z-index:251659264" fillcolor="red"/>` +
      `</w:pict></w:r></w:p></w:body></w:document>`;
    const doc = createDocx({ paragraphs: [] });
    toUint8Array(doc);
    const part = getPart(doc.opc, "/word/document.xml");
    if (!part) throw new Error("no part");
    part.data = new TextEncoder().encode(xml);
    const reopened = openDocx(writeOpcPackage(doc.opc));
    const shape = only(reopened);
    expect(shapePreset(shape)).toBe("triangle");
    expect(getShapeLayout(shape)).toMatchObject({ left: 72, top: 12, width: 75 });
    expect(getShapeLayout(shape).height).toBeCloseTo(56.69, 1);
    expect(getShapeFill(shape)).toEqual({ type: "solid", color: "FF0000" });
    const run = firstParagraph(reopened).children[0];
    expect(run?.kind === "run" && runShape(run)).toBe(shape);
  });
});
