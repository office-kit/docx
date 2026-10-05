import { describe, expect, it } from "vitest";
import {
  getPart,
  partRelationships,
  relationshipsByType,
  readOpcPackage,
} from "../internal/opc/index.js";
import {
  addImage,
  createDocx,
  type Docx,
  imageDrawings,
  openDocx,
  setImageSizeEmu,
  toUint8Array,
} from "./docx.js";
import {
  addChartRun,
  arrangeDrawing,
  changePicture,
  type ChartSpec,
  imageNaturalSizeEmu,
  imagePixelSize,
  readChart,
  readDrawing,
  removeDrawing,
  resetPicture,
  setChart,
  setDrawingAnchorOptions,
  setDrawingAspectLock,
  setDrawingHidden,
  setDrawingHyperlink,
  setDrawingName,
  setDrawingPosition,
  setDrawingTransform,
  setDrawingWrap,
  setPictureColorAdjustments,
  setPictureCrop,
  setPictureEffects,
  setPictureGeometry,
  setPictureOutline,
} from "./picture.js";
import { validatePackage } from "./validator.js";

// A 2x1 PNG with a pHYs chunk of 3780 px/m (96 DPI).
const PNG_2X1 = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x02, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0xf4, 0x22, 0x7f,
  0x8a, 0x00, 0x00, 0x00, 0x09, 0x70, 0x48, 0x59, 0x73, 0x00, 0x00, 0x0e, 0xc4, 0x00, 0x00, 0x0e,
  0xc4, 0x01, 0x95, 0x2b, 0x0e, 0x1b, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae, 0x42,
  0x60, 0x82,
]);

const INCH = 914400;

function docWithPicture(): Docx {
  const doc = createDocx({ paragraphs: ["Before"] });
  addImage(doc, PNG_2X1, { widthEmu: 2 * INCH, heightEmu: INCH, name: "Logo" });
  return doc;
}

/** Save, check the package, reopen. */
function roundTrip(doc: Docx): Docx {
  const bytes = toUint8Array(doc);
  expect(validatePackage(readOpcPackage(bytes))).toEqual([]);
  return openDocx(bytes);
}

function documentXml(doc: Docx): string {
  const part = getPart(readOpcPackage(toUint8Array(doc)), doc.partName);
  return part ? new TextDecoder().decode(part.data) : "";
}

function info(doc: Docx, index = 0) {
  const drawing = imageDrawings(doc)[index];
  if (!drawing) throw new Error("no drawing");
  return readDrawing(doc, drawing);
}

describe("readDrawing", () => {
  it("reads an inline picture's identity, size and image", () => {
    const d = info(roundTrip(docWithPicture()));
    expect(d.kind).toBe("picture");
    expect(d.name).toBe("Logo");
    expect(d.wrap).toBe("inline");
    expect([d.widthEmu, d.heightEmu]).toEqual([2 * INCH, INCH]);
    expect(d.lockAspect).toBe(true);
    expect(d.picture?.image?.contentType).toBe("image/png");
    expect(d.picture?.geometry).toBe("rect");
  });
});

describe("image sizes", () => {
  it("reads PNG pixels and pHYs resolution", () => {
    expect(imagePixelSize(PNG_2X1)).toEqual({ width: 2, height: 1, dpiX: 96, dpiY: 96 });
    expect(imageNaturalSizeEmu(PNG_2X1)).toEqual({ widthEmu: 19050, heightEmu: 9525 });
  });

  it("resizes only the extent and the transform, not extLst entries", () => {
    const doc = docWithPicture();
    setImageSizeEmu(doc, 0, INCH, INCH / 2);
    const xml = documentXml(roundTrip(doc));
    expect(xml).toContain(`<wp:extent cx="${INCH}" cy="${INCH / 2}"/>`);
    expect(xml).toContain(`<a:ext cx="${INCH}" cy="${INCH / 2}"/>`);
    expect(() => setImageSizeEmu(doc, 0, 0, 1)).toThrow(/Invalid size/);
  });
});

describe("wrapping and position", () => {
  it("floats an inline picture with square wrapping and back", () => {
    const doc = docWithPicture();
    setDrawingWrap(doc, 0, "square", {
      horizontal: { relativeTo: "margin", align: "right" },
      vertical: { relativeTo: "paragraph", offsetEmu: 12700 },
    });
    let reopened = roundTrip(doc);
    let d = info(reopened);
    expect(d.wrap).toBe("square");
    expect(d.anchor?.horizontal).toEqual({ relativeTo: "margin", align: "right" });
    expect(d.anchor?.vertical).toEqual({ relativeTo: "paragraph", offsetEmu: 12700 });
    const xml = documentXml(reopened);
    // CT_Anchor child order.
    expect(xml).toMatch(
      /<wp:simplePos[^>]*\/><wp:positionH.*<\/wp:positionH><wp:positionV.*<\/wp:positionV><wp:extent[^>]*\/><wp:effectExtent[^>]*\/><wp:wrapSquare wrapText="bothSides"\/><wp:docPr/,
    );
    setDrawingWrap(reopened, 0, "behindText");
    d = info(roundTrip(reopened));
    expect(d.wrap).toBe("behindText");
    expect(d.anchor?.horizontal).toEqual({ relativeTo: "margin", align: "right" });
    setDrawingWrap(reopened, 0, "inline");
    reopened = roundTrip(reopened);
    d = info(reopened);
    expect(d.wrap).toBe("inline");
    expect(d.anchor).toBeUndefined();
    expect(documentXml(reopened)).toMatch(
      /<wp:inline [^>]*><wp:extent[^>]*\/><wp:effectExtent[^>]*\/><wp:docPr/,
    );
  });

  it("writes tight / through wrap polygons, distances and flags", () => {
    const doc = docWithPicture();
    setDrawingWrap(doc, 0, "tight");
    setDrawingAnchorOptions(doc, 0, {
      wrapSide: "left",
      distance: { top: 1000, left: 2000 },
      allowOverlap: false,
      locked: true,
      wrapPolygon: [
        { x: 0, y: 0 },
        { x: 21600, y: 0 },
        { x: 10800, y: 21600 },
        { x: 0, y: 0 },
      ],
    });
    const d = info(roundTrip(doc));
    expect(d.wrap).toBe("tight");
    expect(d.anchor?.wrapSide).toBe("left");
    expect(d.anchor?.wrapPolygon).toHaveLength(4);
    expect(d.anchor?.distance).toMatchObject({ top: 1000, left: 2000 });
    expect(d.anchor?.allowOverlap).toBe(false);
    expect(d.anchor?.locked).toBe(true);
    setDrawingWrap(doc, 0, "through");
    expect(info(roundTrip(doc)).anchor?.wrapPolygon?.[2]).toEqual({ x: 10800, y: 21600 });
    setDrawingWrap(doc, 0, "topAndBottom");
    expect(info(roundTrip(doc)).wrap).toBe("topAndBottom");
  });

  it("moves a floating object and refuses to move an inline one", () => {
    const doc = docWithPicture();
    expect(() =>
      setDrawingPosition(doc, 0, { horizontal: { relativeTo: "page", offsetEmu: 1 } }),
    ).toThrow(/in line/);
    setDrawingWrap(doc, 0, "inFrontOfText");
    setDrawingPosition(doc, 0, {
      horizontal: { relativeTo: "page", offsetEmu: INCH },
      vertical: { relativeTo: "page", align: "bottom" },
    });
    const d = info(roundTrip(doc));
    expect(d.anchor?.horizontal).toEqual({ relativeTo: "page", offsetEmu: INCH });
    expect(d.anchor?.vertical).toEqual({ relativeTo: "page", align: "bottom" });
  });

  it("restacks floating objects by relativeHeight", () => {
    const doc = docWithPicture();
    addImage(doc, PNG_2X1, { widthEmu: INCH, heightEmu: INCH });
    addImage(doc, PNG_2X1, { widthEmu: INCH, heightEmu: INCH });
    for (const i of [0, 1, 2]) setDrawingWrap(doc, i, "inFrontOfText");
    const heights = (d: Docx): number[] =>
      [0, 1, 2].map((i) => info(d, i).anchor?.relativeHeight ?? -1);
    expect(heights(doc)[0]).toBeLessThan(heights(doc)[2] ?? 0);
    arrangeDrawing(doc, 0, "bringToFront");
    let h = heights(roundTrip(doc));
    expect(Math.max(...h)).toBe(h[0]);
    arrangeDrawing(doc, 0, "sendBackward");
    h = heights(doc);
    expect(h[0]).toBeLessThan(h[2] ?? 0);
    expect(h[0]).toBeGreaterThan(h[1] ?? 0);
    arrangeDrawing(doc, 0, "sendToBack");
    h = heights(doc);
    expect(Math.min(...h)).toBe(h[0]);
    arrangeDrawing(doc, 1, "sendBehindText");
    expect(info(doc, 1).wrap).toBe("behindText");
  });
});

describe("picture format", () => {
  it("rotates and flips through a:xfrm", () => {
    const doc = docWithPicture();
    setDrawingTransform(doc, 0, { rotation: -90, flipH: true });
    const d = info(roundTrip(doc));
    expect(d.rotation).toBe(270);
    expect(d.flipH).toBe(true);
    expect(d.flipV).toBe(false);
  });

  it("crops, reshapes and borders the picture in schema order", () => {
    const doc = docWithPicture();
    setPictureCrop(doc, 0, { left: 10, top: 0, right: 5.5, bottom: 0 });
    setPictureGeometry(doc, 0, "ellipse");
    setPictureOutline(doc, 0, { color: "FF0000", widthEmu: 12700, dash: "dash", compound: "dbl" });
    setPictureEffects(doc, 0, {
      shadow: {
        kind: "outer",
        color: "000000",
        opacity: 40,
        blurEmu: 50800,
        distanceEmu: 38100,
        directionDeg: 45,
        align: "tl",
      },
      reflection: {
        blurEmu: 6350,
        startOpacity: 50,
        endOpacity: 0.3,
        endPosition: 35,
        distanceEmu: 0,
      },
      glow: { radiusEmu: 63500, color: "4472C4", opacity: 60 },
      softEdgeEmu: 31750,
      bevel: { preset: "circle", widthEmu: 63500, heightEmu: 25400 },
      rotation3d: { camera: "perspectiveFront", latitude: 0, longitude: 20, revolution: 0 },
    });
    const reopened = roundTrip(doc);
    const p = info(reopened).picture;
    expect(p?.crop).toEqual({ left: 10, top: 0, right: 5.5, bottom: 0 });
    expect(p?.geometry).toBe("ellipse");
    expect(p?.outline).toEqual({ color: "FF0000", widthEmu: 12700, dash: "dash", compound: "dbl" });
    expect(p?.effects.shadow).toMatchObject({
      kind: "outer",
      opacity: 40,
      directionDeg: 45,
      align: "tl",
    });
    expect(p?.effects.reflection).toMatchObject({ startOpacity: 50, endPosition: 35 });
    expect(p?.effects.glow).toEqual({ radiusEmu: 63500, color: "4472C4", opacity: 60 });
    expect(p?.effects.softEdgeEmu).toBe(31750);
    expect(p?.effects.bevel).toEqual({ preset: "circle", widthEmu: 63500, heightEmu: 25400 });
    expect(p?.effects.rotation3d).toEqual({
      camera: "perspectiveFront",
      latitude: 0,
      longitude: 20,
      revolution: 0,
    });
    const xml = documentXml(reopened);
    expect(xml).toMatch(/<a:blip [^>]*\/><a:srcRect l="10000" t="0" r="5500" b="0"\/><a:stretch>/);
    expect(xml).toMatch(
      /<pic:spPr><a:xfrm>.*<\/a:xfrm><a:prstGeom prst="ellipse">.*<\/a:prstGeom><a:ln .*<\/a:ln><a:effectLst><a:glow.*<\/a:glow><a:outerShdw.*<\/a:outerShdw><a:reflection [^>]*\/><a:softEdge [^>]*\/><\/a:effectLst><a:scene3d>.*<\/a:scene3d><a:sp3d>.*<\/a:sp3d><\/pic:spPr>/,
    );
    setPictureOutline(doc, 0, undefined);
    setPictureEffects(doc, 0, {});
    const cleared = info(roundTrip(doc)).picture;
    expect(cleared?.outline).toBeUndefined();
    expect(cleared?.effects).toEqual({});
    expect(() => setPictureCrop(doc, 0, { left: 60, right: 50, top: 0, bottom: 0 })).toThrow();
    expect(() => setPictureGeometry(doc, 0, 'x"y')).toThrow();
  });

  it("writes inner shadows", () => {
    const doc = docWithPicture();
    setPictureEffects(doc, 0, {
      shadow: {
        kind: "inner",
        color: "000000",
        opacity: 50,
        blurEmu: 63500,
        distanceEmu: 0,
        directionDeg: 0,
      },
    });
    expect(info(roundTrip(doc)).picture?.effects.shadow?.kind).toBe("inner");
  });

  it("writes Corrections, Color and Transparency as blip effects", () => {
    const doc = docWithPicture();
    setPictureColorAdjustments(doc, 0, {
      brightness: 20,
      contrast: -40,
      saturation: 33,
      grayscale: true,
      biLevelThreshold: 25,
      duotone: ["000000", "D9C3A5"],
      transparentColor: "FFFFFF",
      transparency: 30,
    });
    const reopened = roundTrip(doc);
    expect(info(reopened).picture?.adjustments).toEqual({
      brightness: 20,
      contrast: -40,
      saturation: 33,
      grayscale: true,
      biLevelThreshold: 25,
      duotone: ["000000", "D9C3A5"],
      transparentColor: "FFFFFF",
      transparency: 30,
    });
    expect(documentXml(reopened)).toContain(`<a:alphaModFix amt="70000"/>`);
    setPictureColorAdjustments(doc, 0, { brightness: 0 });
    expect(info(roundTrip(doc)).picture?.adjustments).toEqual({});
    expect(() => setPictureColorAdjustments(doc, 0, { contrast: 101 })).toThrow();
  });

  it("changes and resets the picture", () => {
    const doc = docWithPicture();
    setPictureCrop(doc, 0, { left: 10, top: 10, right: 10, bottom: 10 });
    setDrawingTransform(doc, 0, { rotation: 30 });
    setPictureOutline(doc, 0, { color: "000000", widthEmu: 12700, dash: "solid", compound: "sng" });
    const before = info(doc).picture?.image?.partName;
    changePicture(doc, 0, PNG_2X1);
    const after = info(roundTrip(doc)).picture?.image?.partName;
    expect(after).not.toBe(before);
    resetPicture(doc, 0, { size: true });
    const d = info(roundTrip(doc));
    expect(d.picture?.crop).toEqual({ left: 0, top: 0, right: 0, bottom: 0 });
    expect(d.picture?.outline).toBeUndefined();
    expect(d.rotation).toBe(0);
    expect([d.widthEmu, d.heightEmu]).toEqual([19050, 9525]);
  });
});

describe("docPr", () => {
  it("renames, hides, links and locks", () => {
    const doc = docWithPicture();
    setDrawingName(doc, 0, "Company logo");
    setDrawingHidden(doc, 0, true);
    setDrawingHyperlink(doc, 0, "https://example.com/");
    setDrawingAspectLock(doc, 0, false);
    const reopened = roundTrip(doc);
    const d = info(reopened);
    expect(d.name).toBe("Company logo");
    expect(d.hidden).toBe(true);
    expect(d.hyperlink).toBe("https://example.com/");
    expect(d.lockAspect).toBe(false);
    expect(() => setDrawingHyperlink(doc, 0, "javascript:alert(1)")).toThrow(/scheme/);
    setDrawingHyperlink(reopened, 0, undefined);
    expect(info(roundTrip(reopened)).hyperlink).toBeUndefined();
  });

  it("removes a drawing and its run", () => {
    const doc = docWithPicture();
    removeDrawing(doc, 0);
    expect(imageDrawings(roundTrip(doc))).toHaveLength(0);
  });
});

describe("charts", () => {
  const spec: ChartSpec = {
    kind: "column",
    grouping: "clustered",
    title: "Sales",
    legend: "bottom",
    dataLabels: true,
    categories: ["Q1", "Q2", "Q3"],
    series: [
      { name: "2024", values: [1, 2.5, null] },
      { name: "2025", values: [3, 4, 5], color: "FF0000" },
    ],
  };

  function docWithChart(s: ChartSpec = spec): Docx {
    const doc = createDocx({ paragraphs: ["Chart:"] });
    const run = addChartRun(doc, s, { widthEmu: 5 * INCH, heightEmu: 3 * INCH });
    const para = doc.document.body.blocks[0];
    if (para?.kind !== "paragraph") throw new Error("no paragraph");
    para.children.push(run);
    return doc;
  }

  it("writes a chart part, an embedded workbook and the drawing", () => {
    const doc = roundTrip(docWithChart());
    const drawing = imageDrawings(doc)[0];
    if (!drawing) throw new Error("no drawing");
    const d = readDrawing(doc, drawing);
    expect(d.kind).toBe("chart");
    expect(d.chartPartName).toBe("/word/charts/chart1.xml");
    const rels = partRelationships(doc.opc, "/word/charts/chart1.xml");
    const pkgRel = relationshipsByType(
      rels,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package",
    )[0];
    expect(pkgRel?.target).toBe("../embeddings/Microsoft_Excel_Worksheet1.xlsx");
    const workbook = getPart(doc.opc, "/word/embeddings/Microsoft_Excel_Worksheet1.xlsx");
    const sheet = workbook && getPart(readOpcPackage(workbook.data), "/xl/worksheets/sheet1.xml");
    const sheetXml = sheet ? new TextDecoder().decode(sheet.data) : "";
    expect(sheetXml).toContain(
      `<c r="B1" t="inlineStr"><is><t xml:space="preserve">2024</t></is></c>`,
    );
    expect(sheetXml).toContain(`<c r="C4"><v>5</v></c>`);
    const chart = readChart(doc, drawing);
    expect(chart).toEqual({
      ...spec,
      series: [{ name: "2024", values: [1, 2.5, null], color: "4472C4" }, spec.series[1]],
    });
    const chartXml = new TextDecoder().decode(getPart(doc.opc, "/word/charts/chart1.xml")?.data);
    expect(chartXml).toMatch(
      /<c:barChart><c:barDir val="col"\/><c:grouping val="clustered"\/><c:varyColors val="0"\/><c:ser>.*<\/c:ser><c:dLbls>.*<\/c:dLbls><c:gapWidth val="150"\/><c:axId [^>]*\/><c:axId [^>]*\/><\/c:barChart><c:catAx>/,
    );
    expect(chartXml).toContain(
      `<c:externalData r:id="${pkgRel?.id}"><c:autoUpdate val="0"/></c:externalData>`,
    );
  });

  it.each<ChartSpec["kind"]>(["bar", "line", "area", "pie", "doughnut", "scatter", "radar"])(
    "round-trips a %s chart",
    (kind) => {
      const s: ChartSpec = {
        ...spec,
        kind,
        categories: kind === "scatter" ? ["1", "2", "3"] : spec.categories,
        ...(kind === "line" || kind === "scatter" || kind === "radar" ? { markers: true } : {}),
        ...(kind === "pie" || kind === "doughnut" || kind === "scatter" || kind === "radar"
          ? { grouping: undefined }
          : {}),
        series:
          kind === "pie" || kind === "doughnut"
            ? [{ name: "2024", values: [1, 2, 3] }]
            : spec.series,
      };
      const doc = roundTrip(docWithChart(s));
      const drawing = imageDrawings(doc)[0];
      if (!drawing) throw new Error("no drawing");
      const read = readChart(doc, drawing);
      expect(read?.kind).toBe(kind);
      expect(read?.series[0]?.values).toEqual(s.series[0]?.values);
      expect(read?.categories).toEqual(s.categories);
      expect(read?.markers).toBe(s.markers);
    },
  );

  it("rewrites chart data and elements in place", () => {
    const doc = docWithChart();
    setChart(doc, 0, {
      ...spec,
      kind: "line",
      grouping: "stacked",
      legend: "none",
      title: undefined,
      dataLabels: false,
      categories: ["A", "B"],
      series: [{ name: "Only", values: [7, 8] }],
    });
    const reopened = roundTrip(doc);
    const drawing = imageDrawings(reopened)[0];
    if (!drawing) throw new Error("no drawing");
    const chart = readChart(reopened, drawing);
    expect(chart).toMatchObject({
      kind: "line",
      grouping: "stacked",
      legend: "none",
      dataLabels: false,
      categories: ["A", "B"],
    });
    // A single-series chart without a deleted title shows nothing extra: autoTitleDeleted=1 was written.
    expect(chart?.title).toBeUndefined();
    expect(() => setChart(doc, 0, { ...spec, series: [{ name: "x", values: [1] }] })).toThrow(
      /values/,
    );
  });
});
