/**
 * Header / footer paragraphs are rewritten through the same WML paragraph
 * parser and writer as the body whenever text in them is replaced. Their
 * `<w:p>` / `<w:r>` attributes must survive that rewrite and a save → reopen.
 */

import { describe, expect, it } from "vitest";
import {
  addFooter,
  addHeader,
  createDocx,
  footers,
  getRawPartRoot,
  headers,
  markRawPartDirty,
  openDocx,
  replaceTextEverywhere,
  toUint8Array,
  type Docx,
  type XmlAttr,
  type XmlElement,
} from "../index.js";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

function rsid(local: string, value: string): XmlAttr {
  return { name: { uri: W, local, prefix: "w" }, value, isNamespaceDecl: false };
}

function findElement(root: XmlElement, local: string): XmlElement | undefined {
  for (const child of root.children) {
    if (child.kind !== "element") continue;
    if (child.name.uri === W && child.name.local === local) return child;
    const nested = findElement(child, local);
    if (nested) return nested;
  }
  return undefined;
}

function firstElement(root: XmlElement, local: string): XmlElement {
  const el = findElement(root, local);
  if (!el) throw new Error(`no <w:${local}>`);
  return el;
}

/** Stamp rsid attributes on the first paragraph and run of a part, and save. */
function stamp(doc: Docx, partName: string): Uint8Array {
  const root = getRawPartRoot(doc, partName);
  if (!root) throw new Error(`missing part ${partName}`);
  // The raw-part editor hands out the live tree; attrs are readonly in the type
  // because ordinary callers should go through setElementAttr (unprefixed only).
  (firstElement(root, "p").attrs as XmlAttr[]).push(rsid("rsidR", "00AA0001"));
  (firstElement(root, "r").attrs as XmlAttr[]).push(rsid("rsidRPr", "00AA0002"));
  markRawPartDirty(doc, partName);
  return toUint8Array(doc);
}

describe("header / footer attribute preservation", () => {
  it("keeps <w:p>/<w:r> attributes when header and footer text is replaced", () => {
    const doc = createDocx({ paragraphs: ["Body"] });
    addHeader(doc, "Header DRAFT");
    addFooter(doc, "Footer DRAFT");
    const headerPart = headers(doc)[0]!.partName;
    const footerPart = footers(doc)[0]!.partName;
    stamp(doc, headerPart);
    const stamped = openDocx(stamp(doc, footerPart));

    expect(replaceTextEverywhere(stamped, "DRAFT", "FINAL")).toBe(2);
    const reopened = openDocx(toUint8Array(stamped));

    for (const part of [headerPart, footerPart]) {
      const root = getRawPartRoot(reopened, part)!;
      const p = firstElement(root, "p");
      const r = firstElement(root, "r");
      expect(p.attrs).toContainEqual(rsid("rsidR", "00AA0001"));
      expect(r.attrs).toContainEqual(rsid("rsidRPr", "00AA0002"));
    }
    expect(headers(reopened)[0]!.text).toContain("FINAL");
  });
});
