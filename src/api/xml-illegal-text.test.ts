import { strFromU8, unzipSync } from "fflate";
import { describe, expect, it } from "vitest";
import {
  addFooter,
  addHeader,
  createDocx,
  openDocx,
  paragraphs,
  paragraphText,
  toUint8Array,
} from "./index.js";

const DOCUMENT_PART = "word/document.xml";

// XML 1.0 forbids `]]>` in character data and has no way to represent the
// code points below, not even as character references (#37).
const NON_XML_CHARS = ["\u0000", "\u0001", "\u000B", "\u000C", "\u001F", "\uD800", "￾", "￿"];

function label(ch: string): string {
  return `U+${ch.charCodeAt(0).toString(16).toUpperCase().padStart(4, "0")}`;
}

describe("text that XML 1.0 cannot hold verbatim (#37)", () => {
  it("escapes `]]>` in body text and reads it back", () => {
    const text = "if (a[b[0]]> 1) return";
    const bytes = toUint8Array(createDocx({ paragraphs: [text] }));

    const xml = strFromU8(unzipSync(bytes)[DOCUMENT_PART]!);
    expect(xml).not.toContain("]]>");
    expect(xml).toContain("a[b[0]]&gt; 1");
    expect(paragraphText(paragraphs(openDocx(bytes))[0]!)).toBe(text);
  });

  for (const ch of NON_XML_CHARS) {
    it(`rejects ${label(ch)} in body text when saving`, () => {
      const doc = createDocx({ paragraphs: [`copied${ch}from a PDF`] });
      expect(() => toUint8Array(doc)).toThrow(
        new RangeError(`${label(ch)} is not allowed in XML 1.0`),
      );
    });
  }

  it("rejects a forbidden character in header and footer text", () => {
    const doc = createDocx();
    expect(() => addHeader(doc, "a\u000Bb")).toThrow(RangeError);
    expect(() => addFooter(doc, "a\u000Bb")).toThrow(RangeError);
  });

  it("still accepts tab, private-use and astral characters", () => {
    const text = "a\tb 😀 �";
    const bytes = toUint8Array(createDocx({ paragraphs: [text] }));
    expect(paragraphText(paragraphs(openDocx(bytes))[0]!)).toBe(text);
  });
});
