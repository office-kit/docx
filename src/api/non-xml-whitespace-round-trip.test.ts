import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { describe, expect, it } from "vitest";
import { createDocx, openDocx, paragraphs, paragraphText, toUint8Array } from "./index.js";

const DOCUMENT_PART = "word/document.xml";
const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

// French typography puts a (narrow) no-break space alone in its own run before
// `:` and inside « », written by Word without xml:space="preserve". Such a run
// must survive open → save untouched (#35).
function withBody(bodyXml: string): Uint8Array {
  const files = unzipSync(toUint8Array(createDocx()));
  files[DOCUMENT_PART] = strToU8(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<w:document xmlns:w="${W_NS}"><w:body>${bodyXml}</w:body></w:document>`,
  );
  return zipSync(files);
}

describe("non-XML white space in <w:t> (#35)", () => {
  for (const ch of [" ", " ", "　"]) {
    const label = `U+${ch.charCodeAt(0).toString(16).toUpperCase().padStart(4, "0")}`;

    it(`keeps a run holding only ${label} on read and on save`, () => {
      const doc = openDocx(
        withBody(
          `<w:p><w:r><w:t>Nom</w:t></w:r><w:r><w:t>${ch}</w:t></w:r><w:r><w:t>: Dupont</w:t></w:r></w:p>`,
        ),
      );
      expect(paragraphText(paragraphs(doc)[0]!)).toBe(`Nom${ch}: Dupont`);

      const saved = strFromU8(unzipSync(toUint8Array(doc))[DOCUMENT_PART]!);
      expect(saved).toContain(`<w:t>${ch}</w:t>`);

      const reopened = openDocx(toUint8Array(doc));
      expect(paragraphText(paragraphs(reopened)[0]!)).toBe(`Nom${ch}: Dupont`);
    });
  }
});
