import { describe, expect, it } from "vitest";
import {
  fieldSwitch,
  formatFieldNumber,
  formatNumber,
  parseFieldInstruction,
  quoteFieldArg,
} from "./field-instruction.js";
import {
  blocksFromXml,
  estimatePages,
  paragraphFlow,
  scanFields,
  withOpenContentControls,
} from "./field-flow.js";

describe("parseFieldInstruction", () => {
  it("splits type, arguments and switches", () => {
    const f = parseFieldInstruction(' TOC \\o "1-3" \\h \\z \\u ');
    expect(f.type).toBe("TOC");
    expect(f.switches).toEqual([
      { name: "o", arg: "1-3" },
      { name: "h" },
      { name: "z" },
      { name: "u" },
    ]);
    const xe = parseFieldInstruction('XE "Main\\:x:Sub" \\b \\t "See \\"A\\""');
    expect(xe.args).toEqual(["Main\\:x:Sub"]);
    expect(fieldSwitch(xe, "t")?.arg).toBe('See "A"');
    expect(fieldSwitch(xe, "b")).toEqual({ name: "b" });
    const seq = parseFieldInstruction("SEQ Figure \\* ROMAN \\s 1");
    expect(seq.args).toEqual(["Figure"]);
    expect(formatFieldNumber(4, seq)).toBe("IV");
  });

  it("quotes arguments only when needed", () => {
    expect(quoteFieldArg("Figure")).toBe("Figure");
    expect(quoteFieldArg("My Label")).toBe('"My Label"');
    expect(quoteFieldArg('a"b')).toBe('"a\\"b"');
  });
});

describe("formatNumber", () => {
  it.each([
    [4, "upperRoman", "IV"],
    [1994, "lowerRoman", "mcmxciv"],
    [28, "lowerLetter", "bb"],
    [3, "upperLetter", "C"],
    [22, "ordinal", "22nd"],
    [13, "ordinal", "13th"],
    [121, "cardinalText", "one hundred twenty-one"],
    [42, "ordinalText", "forty-second"],
    [12, "ordinalText", "twelfth"],
    [6, "chicago", "††"],
    [255, "hex", "FF"],
  ] as const)("%i as %s is %s", (n, format, expected) => {
    expect(formatNumber(n, format)).toBe(expected);
  });
});

describe("field flow", () => {
  it("finds nested fields and fields inside a content control", () => {
    const body = {
      blocks: blocksFromXml(
        '<w:sdt><w:sdtPr/><w:sdtContent><w:p><w:r><w:fldChar w:fldCharType="begin"/><w:instrText>TOC \\o "1-1"</w:instrText><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>x</w:t></w:r></w:p><w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:sdtContent></w:sdt><w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>IF 1 = 1 "</w:instrText></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>PAGE</w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:instrText>"</w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>',
      ),
      extras: [],
    };
    const types = withOpenContentControls(body, () =>
      scanFields(paragraphFlow(body)).map((f) => `${f.parsed.type}@${f.depth}`),
    );
    expect(types).toEqual(["TOC@0", "IF@0", "PAGE@1"]);
    expect(body.blocks[0]?.kind).toBe("raw");
  });

  it("estimates pages from breaks and rendered page marks", () => {
    const body = {
      blocks: blocksFromXml(
        '<w:p><w:r><w:t>a</w:t></w:r></w:p><w:p><w:r><w:br w:type="page"/></w:r></w:p><w:p><w:r><w:lastRenderedPageBreak/><w:t>b</w:t></w:r></w:p><w:p><w:r><w:lastRenderedPageBreak/><w:t>c</w:t></w:r></w:p><w:p><w:pPr><w:sectPr><w:type w:val="continuous"/></w:sectPr></w:pPr></w:p><w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t>d</w:t></w:r></w:p>',
      ),
      extras: [],
    };
    const pages = estimatePages(body);
    const flow = paragraphFlow(body).map((r) => pages.get(r.paragraph));
    expect(flow).toEqual([1, 1, 2, 3, 3, 4]);
  });
});
