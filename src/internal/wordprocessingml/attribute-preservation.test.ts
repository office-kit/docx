/**
 * Attributes on structural WML elements (`w14:paraId`, `w14:textId`,
 * `w:rsid*`, …) must survive parse → write. Word uses `w14:paraId` to anchor
 * comments-extended data and collaboration state, so dropping it on every
 * save loses data the caller never touched.
 */

import { describe, expect, it } from "vitest";
import { parseXml, serializeXml } from "../xml/index.js";
import { parseWmlDocument } from "./parser.js";
import { writeWmlDocument } from "./writer.js";

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';

function cycle(body: string, bodyAttrs = ""): string {
  const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body${bodyAttrs}>${body}</w:body></w:document>`;
  return serializeXml(writeWmlDocument(parseWmlDocument(parseXml(xml))));
}

describe("structural attribute preservation", () => {
  it("keeps paragraph and run attributes", () => {
    const out = cycle(
      '<w:p w14:paraId="1A2B3C4D" w14:textId="77777777" w:rsidR="00AB12CD">' +
        '<w:r w:rsidRPr="00EF3456"><w:t>Hi</w:t></w:r></w:p>',
    );
    expect(out).toContain('w14:paraId="1A2B3C4D"');
    expect(out).toContain('w14:textId="77777777"');
    expect(out).toContain('w:rsidR="00AB12CD"');
    expect(out).toContain('w:rsidRPr="00EF3456"');
  });

  it("keeps table, row and cell attributes, including cell paragraphs", () => {
    const out = cycle(
      '<w:tbl w:rsidR="00000001"><w:tblGrid><w:gridCol w:w="100"/></w:tblGrid>' +
        '<w:tr w14:paraId="0000AAAA" w:rsidR="00000002"><w:tc>' +
        '<w:p w14:paraId="0000BBBB"><w:r><w:t>c</w:t></w:r></w:p>' +
        "</w:tc></w:tr></w:tbl>",
    );
    expect(out).toContain('w14:paraId="0000AAAA"');
    expect(out).toContain('w14:paraId="0000BBBB"');
    expect(out).toContain('w:rsidR="00000001"');
    expect(out).toContain('w:rsidR="00000002"');
  });

  it("keeps body attributes", () => {
    const out = cycle("<w:p/>", ' w:rsidR="0000FFFF"');
    expect(out).toContain('<w:body w:rsidR="0000FFFF">');
  });

  it("is stable across two cycles (no duplicated attributes)", () => {
    const once = cycle('<w:p w14:paraId="12345678"><w:r><w:t>x</w:t></w:r></w:p>');
    const body = once.slice(once.indexOf("<w:body>") + 8, once.indexOf("</w:body>"));
    const twice = cycle(body);
    expect(twice.match(/w14:paraId=/g)).toHaveLength(1);
    expect(twice).toBe(once);
  });
});
