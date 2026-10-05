import { describe, expect, it } from "vitest";
import { getPart, partRelationships, relationshipsByType } from "../internal/opc/index.js";
import { parseXml } from "../internal/xml/index.js";
import { createDocx, openDocx, paragraphs, toUint8Array } from "./docx.js";
import {
  addSmartArt,
  getSmartArt,
  runSmartArt,
  SMARTART_LAYOUTS,
  type SmartArtLayout,
  setSmartArtColors,
  setSmartArtLayout,
  setSmartArtNodes,
  setSmartArtSize,
  setSmartArtStyle,
  smartArts,
} from "./smartart.js";
import { validatePackage } from "./validator.js";

const NODES = [
  { text: "Plan", children: [{ text: "Scope" }] },
  { text: "Build" },
  { text: "Ship & <celebrate>" },
];

function setup(layout: SmartArtLayout = "basicBlockList") {
  const doc = createDocx({ paragraphs: ["Anchor"] });
  const p = paragraphs(doc)[0];
  if (!p) throw new Error("no paragraph");
  const run = addSmartArt(doc, p, { layout, nodes: NODES, width: 400, height: 200 });
  return { doc, run };
}

describe("addSmartArt", () => {
  it("writes the drawing, the four diagram parts and their relationships", () => {
    const { doc } = setup();
    const reopened = openDocx(toUint8Array(doc));
    expect(validatePackage(reopened.opc)).toEqual([]);
    const rels = partRelationships(reopened.opc, "/word/document.xml");
    for (const type of ["diagramData", "diagramLayout", "diagramQuickStyle", "diagramColors"]) {
      expect(
        relationshipsByType(
          rels,
          `http://schemas.openxmlformats.org/officeDocument/2006/relationships/${type}`,
        ),
      ).toHaveLength(1);
    }
    const docXml = new TextDecoder().decode(getPart(reopened.opc, "/word/document.xml")?.data);
    expect(docXml).toContain('uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"');
    expect(docXml).toContain("<dgm:relIds");
    // Every part is well-formed XML with the expected root.
    const roots = ["data1", "layout1", "quickStyle1", "colors1"].map(
      (n) =>
        parseXml(new TextDecoder().decode(getPart(reopened.opc, `/word/diagrams/${n}.xml`)?.data))
          .root.name.local,
    );
    expect(roots).toEqual(["dataModel", "layoutDef", "styleDef", "colorsDef"]);
  });

  it("reads back the bullets, layout, colours, style and size", () => {
    const { doc } = setup("basicProcess");
    const reopened = openDocx(toUint8Array(doc));
    const [ref] = smartArts(reopened);
    if (!ref) throw new Error("no SmartArt");
    expect(getSmartArt(reopened, ref)).toEqual({
      layout: "basicProcess",
      colors: "accent1_2",
      style: "simple1",
      nodes: NODES,
      width: 400,
      height: 200,
    });
  });

  it("writes every layout", () => {
    for (const layout of Object.keys(SMARTART_LAYOUTS) as SmartArtLayout[]) {
      const { doc } = setup(layout);
      const reopened = openDocx(toUint8Array(doc));
      expect(validatePackage(reopened.opc)).toEqual([]);
      const [ref] = smartArts(reopened);
      if (!ref) throw new Error("no SmartArt");
      expect(getSmartArt(reopened, ref).layout).toBe(layout);
    }
  });
});

describe("editing SmartArt", () => {
  it("rewrites the bullets, layout, colours, style and size", () => {
    const { doc, run } = setup();
    const ref = runSmartArt(doc, run);
    if (!ref) throw new Error("no SmartArt");
    setSmartArtNodes(doc, ref, [{ text: "Only" }]);
    setSmartArtLayout(doc, ref, "basicCycle");
    setSmartArtColors(doc, ref, "colorful1");
    setSmartArtStyle(doc, ref, "simple5");
    setSmartArtSize(ref, 300, 150);
    const reopened = openDocx(toUint8Array(doc));
    expect(validatePackage(reopened.opc)).toEqual([]);
    const [back] = smartArts(reopened);
    if (!back) throw new Error("no SmartArt");
    expect(getSmartArt(reopened, back)).toEqual({
      layout: "basicCycle",
      colors: "colorful1",
      style: "simple5",
      nodes: [{ text: "Only" }],
      width: 300,
      height: 150,
    });
  });
});
