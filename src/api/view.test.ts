import { describe, expect, it } from "vitest";
import { getPart } from "../internal/opc/index.js";
import { createDocx, type Docx, openDocx, toUint8Array } from "./docx.js";
import { validatePackage } from "./validator.js";
import {
  customProperties,
  documentView,
  documentZoom,
  removeCustomProperty,
  setCustomProperty,
  setDocumentView,
  setDocumentZoom,
} from "./view.js";

function partText(doc: Docx, name: string): string {
  return new TextDecoder().decode(getPart(openDocx(toUint8Array(doc)).opc, name)?.data);
}

describe("document view and zoom", () => {
  it("round-trips w:view and w:zoom in schema order", () => {
    const doc = createDocx({ paragraphs: ["a"] });
    setDocumentZoom(doc, { percent: 150, preset: "bestFit" });
    setDocumentView(doc, "draft");
    const reopened = openDocx(toUint8Array(doc));
    expect(documentView(reopened)).toBe("draft");
    expect(documentZoom(reopened)).toEqual({ percent: 150, preset: "bestFit" });
    expect(partText(doc, "/word/settings.xml")).toContain(
      '<w:view w:val="normal"/><w:zoom w:val="bestFit" w:percent="150"/>',
    );
    setDocumentView(doc, "print");
    expect(documentView(doc)).toBe("print");
    expect(() => setDocumentZoom(doc, { percent: 5, preset: "none" })).toThrow(RangeError);
  });
});

describe("custom properties", () => {
  it("adds, changes and removes typed properties in docProps/custom.xml", () => {
    const doc = createDocx({ paragraphs: ["a"] });
    const when = new Date("2026-10-05T12:00:00Z");
    setCustomProperty(doc, "Client", "Acme");
    setCustomProperty(doc, "Budget", 12);
    setCustomProperty(doc, "Rate", 1.5);
    setCustomProperty(doc, "Approved", true);
    setCustomProperty(doc, "Due", when);
    setCustomProperty(doc, "Client", "Globex");
    const reopened = openDocx(toUint8Array(doc));
    expect(customProperties(reopened)).toEqual([
      { name: "Client", value: "Globex" },
      { name: "Budget", value: 12 },
      { name: "Rate", value: 1.5 },
      { name: "Approved", value: true },
      { name: "Due", value: when },
    ]);
    const xml = partText(doc, "/docProps/custom.xml");
    expect(xml).toContain(
      '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="Client"><vt:lpwstr>Globex</vt:lpwstr></property>',
    );
    expect(xml).toContain('pid="3" name="Budget"><vt:i4>12</vt:i4>');
    expect(xml).toContain("<vt:r8>1.5</vt:r8>");
    expect(xml).toContain("<vt:filetime>2026-10-05T12:00:00Z</vt:filetime>");
    expect(partText(doc, "/[Content_Types].xml")).toContain(
      "application/vnd.openxmlformats-officedocument.custom-properties+xml",
    );
    expect(validatePackage(reopened.opc)).toEqual([]);
    expect(removeCustomProperty(doc, "Budget")).toBe(true);
    expect(removeCustomProperty(doc, "Budget")).toBe(false);
    expect(customProperties(doc).map((p) => p.name)).toEqual(["Client", "Rate", "Approved", "Due"]);
  });
});
