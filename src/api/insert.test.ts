import { describe, expect, it } from "vitest";
import { getPart } from "../internal/opc/index.js";
import {
  addFootnote,
  appendHeading,
  appendParagraph,
  createDocx,
  type Docx,
  ensureHeadingStyles,
  footers,
  headers,
  openDocx,
  paragraphs,
  setCoreProperties,
  toUint8Array,
  type WmlParagraph,
} from "./index.js";
import {
  buildEquation,
  complexFields,
  editHyperlink,
  ensureReferenceBookmark,
  equationLinear,
  formatDatePicture,
  getDropCap,
  getPageNumberFormat,
  hasCoverPage,
  headerFooterParagraphs,
  insertBookmark,
  insertCoverPage,
  insertDocumentContent,
  insertEquation,
  insertField,
  insertFormField,
  insertHyperlink,
  insertPageNumbers,
  insertSignatureLine,
  insertSymbol,
  isValidBookmarkName,
  paragraphEquations,
  paragraphHyperlinks,
  removeCoverPage,
  removeHeaderFooter,
  removeHyperlink,
  removePageNumbers,
  setDropCap,
  setEquation,
  setHeaderFooterParagraphs,
  setPageNumberFormat,
  updateFields,
} from "./insert.js";
import { validatePackage } from "./validator.js";

function partXml(doc: Docx, name: string): string {
  const part = getPart(doc.opc, name);
  return new TextDecoder().decode(part?.data ?? new Uint8Array());
}

/** Save, reopen, check the package is valid, and return the reopened doc + document.xml. */
function roundTrip(doc: Docx): { doc: Docx; xml: string } {
  const reopened = openDocx(toUint8Array(doc));
  expect(validatePackage(reopened.opc)).toEqual([]);
  return { doc: reopened, xml: partXml(reopened, "/word/document.xml") };
}

function para(doc: Docx, i = 0): WmlParagraph {
  const p = paragraphs(doc)[i];
  if (!p) throw new Error(`no paragraph ${i}`);
  return p;
}

function textOf(p: WmlParagraph): string {
  let out = "";
  for (const c of p.children) {
    if (c.kind !== "run") continue;
    for (const piece of c.pieces) if (piece.kind === "text") out += piece.value;
  }
  return out;
}

// A fixed clock: Tuesday 2026-03-03 14:05:09 local time.
const NOW = new Date(2026, 2, 3, 14, 5, 9);

describe("formatDatePicture", () => {
  it("formats Word's date and time pictures", () => {
    expect(formatDatePicture(NOW, "M/d/yyyy")).toBe("3/3/2026");
    expect(formatDatePicture(NOW, "dddd, MMMM d, yyyy")).toBe("Tuesday, March 3, 2026");
    expect(formatDatePicture(NOW, "dd-MMM-yy")).toBe("03-Mar-26");
    expect(formatDatePicture(NOW, "h:mm am/pm")).toBe("2:05 pm");
    expect(formatDatePicture(NOW, "HH:mm:ss")).toBe("14:05:09");
    expect(formatDatePicture(NOW, "yyyy'年'M'月'd'日'")).toBe("2026年3月3日");
  });
});

describe("insertField", () => {
  it("inserts a complex field at the caret with a computed DATE result", () => {
    const doc = createDocx({ paragraphs: ["Today is ."] });
    insertField(doc, para(doc), 9, 'DATE \\@ "MMMM d, yyyy"', { context: { now: NOW } });
    expect(textOf(para(doc))).toBe("Today is March 3, 2026.");
    const { xml, doc: back } = roundTrip(doc);
    expect(xml).toContain('<w:fldChar w:fldCharType="begin"/>');
    expect(xml).toContain('<w:instrText xml:space="preserve"> DATE \\@ "MMMM d, yyyy" </w:instrText>');
    expect(xml).toContain('<w:fldChar w:fldCharType="separate"/>');
    expect(xml).toContain('<w:fldChar w:fldCharType="end"/>');
    expect(complexFields(back).map((f) => f.type)).toEqual(["DATE"]);
  });

  it("computes document-property, statistics, formula and IF fields", () => {
    const doc = createDocx({ paragraphs: ["one two three", ""] });
    setCoreProperties(doc, { title: "Report", creator: "Ada" });
    const p = para(doc, 1);
    insertField(doc, p, 0, "TITLE \\* Upper");
    insertField(doc, p, textOf(p).length, "AUTHOR");
    // Counts the body words plus the results already shown ("REPORTAda").
    insertField(doc, p, textOf(p).length, "NUMWORDS");
    insertField(doc, p, textOf(p).length, "= 2 * (3 + 4) \\# 0.00");
    insertField(doc, p, textOf(p).length, 'IF 1 > 2 "yes" "no"');
    insertField(doc, p, textOf(p).length, "QUOTE \"hi\"");
    insertField(doc, p, textOf(p).length, "SYMBOL 169");
    expect(complexFields(doc).map((f) => f.result)).toEqual(["REPORT", "Ada", "4", "14.00", "no", "hi", "©"]);
    roundTrip(doc);
  });

  it("numbers SEQ fields in document order and resolves REF to bookmark text", () => {
    const doc = createDocx({ paragraphs: ["Figure ", "Figure ", "See "] });
    insertField(doc, para(doc, 0), 7, "SEQ Figure \\* ARABIC");
    insertField(doc, para(doc, 1), 7, "SEQ Figure \\* ROMAN");
    insertBookmark(doc, "Second", { paragraph: para(doc, 1), offset: 0 }, { paragraph: para(doc, 1), offset: 9 });
    insertField(doc, para(doc, 2), 4, "REF Second \\h");
    insertField(doc, para(doc, 2), textOf(para(doc, 2)).length, "REF Second \\p");
    expect(textOf(para(doc, 0))).toBe("Figure 1");
    expect(textOf(para(doc, 1))).toBe("Figure II");
    expect(textOf(para(doc, 2))).toBe("See Figure IIabove");
    roundTrip(doc);
  });

  it("updates every field and leaves locked and uncomputable ones alone", () => {
    const doc = createDocx({ paragraphs: [""] });
    const p = para(doc);
    insertField(doc, p, 0, "PAGE", { result: "9" });
    insertField(doc, p, 1, "DATE \\@ yyyy \\!", { result: "1999" });
    insertField(doc, p, 5, "TIME \\@ HH:mm", { result: "00:00" });
    insertField(doc, p, 10, "INCLUDETEXT x.docx", { result: "kept" });
    const changed = updateFields(doc, { now: NOW, page: 1 });
    expect(changed).toBe(2);
    expect(complexFields(doc).map((f) => f.result)).toEqual(["1", "1999", "14:05", "kept"]);
    expect(updateFields(doc, { now: new Date(2026, 2, 3, 9, 0), types: ["PAGE"] })).toBe(0);
  });

  it("rejects an offset outside the paragraph", () => {
    const doc = createDocx({ paragraphs: ["ab"] });
    expect(() => insertField(doc, para(doc), 3, "PAGE")).toThrow(RangeError);
  });

  it("writes legacy form fields with ffData", () => {
    const doc = createDocx({ paragraphs: ["Name: "] });
    insertFormField(doc, para(doc), 6, { kind: "text", name: "Name", defaultText: "Jane", maxLength: 20 });
    insertFormField(doc, para(doc), textOf(para(doc)).length, { kind: "checkBox", name: "Agree", checked: true });
    insertFormField(doc, para(doc), textOf(para(doc)).length, { kind: "dropDown", name: "Size", entries: ["S", "M", "L"], selected: 1 });
    const { xml } = roundTrip(doc);
    expect(xml).toContain('<w:ffData><w:name w:val="Name"/><w:enabled/><w:calcOnExit w:val="0"/><w:textInput><w:default w:val="Jane"/><w:maxLength w:val="20"/></w:textInput></w:ffData>');
    expect(xml).toContain('<w:checkBox><w:sizeAuto/><w:default w:val="1"/></w:checkBox>');
    expect(xml).toContain('<w:ddList><w:default w:val="1"/><w:listEntry w:val="S"/><w:listEntry w:val="M"/><w:listEntry w:val="L"/></w:ddList>');
    expect(xml).toContain("FORMDROPDOWN");
  });
});

describe("insertSymbol", () => {
  it("writes w:sym and the special hyphens", () => {
    const doc = createDocx({ paragraphs: ["ab"] });
    insertSymbol(doc, para(doc), 1, { font: "Wingdings", char: "f0e0" });
    insertSymbol(doc, para(doc), 2, "noBreakHyphen");
    insertSymbol(doc, para(doc), 3, "softHyphen");
    const { xml } = roundTrip(doc);
    expect(xml).toContain('<w:sym w:font="Wingdings" w:char="F0E0"/>');
    expect(xml).toContain("<w:noBreakHyphen/>");
    expect(xml).toContain("<w:softHyphen/>");
    expect(xml.indexOf("<w:t>a</w:t>")).toBeLessThan(xml.indexOf("<w:sym"));
  });

  it("rejects a malformed char code", () => {
    const doc = createDocx({ paragraphs: ["a"] });
    expect(() => insertSymbol(doc, para(doc), 0, { font: "Symbol", char: "zz" })).toThrow(RangeError);
  });
});

describe("hyperlinks", () => {
  it("inserts, edits and removes a link at the caret", () => {
    const doc = createDocx({ paragraphs: ["Visit  today"] });
    const link = insertHyperlink(doc, para(doc), 6, "our site", {
      url: "https://example.com/",
      tooltip: "Example",
      targetFrame: "_blank",
    });
    let { xml } = roundTrip(doc);
    expect(xml).toMatch(/<w:hyperlink r:id="rId\d+" w:tgtFrame="_blank" w:tooltip="Example" w:history="1"><w:r><w:rPr><w:rStyle w:val="Hyperlink"\/><\/w:rPr><w:t>our site<\/w:t><\/w:r><\/w:hyperlink>/);
    expect(paragraphHyperlinks(doc, para(doc))[0]).toMatchObject({ url: "https://example.com/", text: "our site", tooltip: "Example" });

    const edited = editHyperlink(doc, para(doc), link, { bookmark: "Top" }, "the top");
    ({ xml } = roundTrip(doc));
    expect(xml).toContain('<w:hyperlink w:anchor="Top" w:history="1">');
    expect(paragraphHyperlinks(doc, para(doc))[0]?.text).toBe("the top");

    removeHyperlink(doc, para(doc), edited);
    ({ xml } = roundTrip(doc));
    expect(xml).not.toContain("w:hyperlink");
    expect(xml).toContain("<w:t>the top</w:t>");
  });

  it("refuses unsafe schemes", () => {
    const doc = createDocx({ paragraphs: ["x"] });
    expect(() => insertHyperlink(doc, para(doc), 0, "x", { url: "javascript:alert(1)" })).toThrow(/scheme/);
  });
});

describe("bookmarks", () => {
  it("brackets a range, across paragraphs too", () => {
    const doc = createDocx({ paragraphs: ["Hello world", "second"] });
    insertBookmark(doc, "Greeting", { paragraph: para(doc, 0), offset: 0 }, { paragraph: para(doc, 0), offset: 5 });
    insertBookmark(doc, "Span", { paragraph: para(doc, 0), offset: 6 }, { paragraph: para(doc, 1), offset: 3 });
    const { xml } = roundTrip(doc);
    expect(xml).toContain('<w:bookmarkStart w:id="0" w:name="Greeting"/><w:r><w:t>Hello</w:t></w:r><w:bookmarkEnd w:id="0"/>');
    expect(xml).toContain('<w:bookmarkStart w:id="1" w:name="Span"/>');
    expect(() => insertBookmark(doc, "Greeting", { paragraph: para(doc, 0), offset: 0 })).toThrow(/exists/);
    expect(isValidBookmarkName("1abc")).toBe(false);
    expect(isValidBookmarkName("_Hidden")).toBe(true);
    expect(isValidBookmarkName("a".repeat(41))).toBe(false);
  });

  it("creates one hidden _Ref bookmark per target and reuses it", () => {
    const doc = createDocx({ paragraphs: [] });
    ensureHeadingStyles(doc);
    const heading = appendHeading(doc, "Intro", 1);
    const a = ensureReferenceBookmark(doc, heading);
    const b = ensureReferenceBookmark(doc, heading);
    expect(a).toMatch(/^_Ref\d{9}$/);
    expect(b).toBe(a);
    appendParagraph(doc, "See ");
    insertField(doc, para(doc, 1), 4, `REF ${a} \\h`);
    expect(textOf(para(doc, 1))).toBe("See Intro");
    roundTrip(doc);
  });

  it("references a footnote number through NOTEREF", () => {
    const doc = createDocx({ paragraphs: ["Claim", "As in note "] });
    addFootnote(doc, para(doc, 0), "Source.");
    const refRun = para(doc, 0).children.find((c) => c.kind === "raw");
    expect(refRun).toBeDefined();
    // Reopen so the reference run is a typed run, as in an edited document.
    const reopened = openDocx(toUint8Array(doc));
    const first = para(reopened, 0);
    const run = first.children.find(
      (c) => c.kind === "run" && c.pieces.some((p) => p.kind === "raw" && p.node.name.local === "footnoteReference"),
    );
    if (run?.kind !== "run") throw new Error("no footnote reference run");
    const name = ensureReferenceBookmark(reopened, first, run);
    insertField(reopened, para(reopened, 1), 11, `NOTEREF ${name} \\h`);
    expect(textOf(para(reopened, 1))).toBe("As in note 1");
    roundTrip(reopened);
  });
});

describe("equations", () => {
  const cases: Array<[string, string]> = [
    ["a/b", "<m:f><m:num>"],
    ["x^2+y_i", "<m:sSup>"],
    ["x_i^2", "<m:sSubSup>"],
    ["√(x+1)", '<m:radPr><m:degHide m:val="1"/></m:radPr>'],
    ["√(3&x)", "<m:deg><m:r>"],
    ["∑_(i=1)^n▒i", '<m:chr m:val="∑"/><m:limLoc m:val="undOvr"/>'],
    ["∫_0^1▒x dx", '<m:limLoc m:val="subSup"/>'],
    ["[a│b]", '<m:begChr m:val="["/><m:sepChr m:val="│"/>'],
    ["|x|", '<m:begChr m:val="|"/><m:endChr m:val="|"/>'],
    ["sin x", '<m:fName><m:r><m:rPr><m:sty m:val="p"/>'],
    ["lim_(n→∞)a_n", "<m:limLow>"],
    ["x\\hat", '<m:acc><m:accPr><m:chr m:val="̂"/>'],
    ["¯(ab)", '<m:pos m:val="top"/>'],
    ["■(a&b@c&d)", "<m:m><m:mr>"],
    ["█(x=1@y=2)", "<m:eqArr>"],
    ["▭(E=mc^2)", "<m:borderBox>"],
    ["\\alpha+\\beta", "α+β"],
  ];
  for (const [linear, fragment] of cases) {
    it(`builds OMML for ${linear}`, () => {
      const doc = createDocx({ paragraphs: ["Eq: "] });
      insertEquation(doc, para(doc), 4, linear);
      const { xml, doc: back } = roundTrip(doc);
      expect(xml).toContain('<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">');
      expect(xml).toContain(fragment);
      // Linear → OMML → linear → OMML is stable.
      const eq = paragraphEquations(para(back))[0];
      if (!eq) throw new Error("equation lost");
      const again = equationLinear(eq);
      expect(equationLinear(buildEquation(again))).toBe(again);
    });
  }

  it("writes display equations as oMathPara and replaces equations in place", () => {
    const doc = createDocx({ paragraphs: [""] });
    insertEquation(doc, para(doc), 0, "E=mc^2", { display: true });
    setEquation(doc, para(doc), 0, "a^2+b^2=c^2");
    const { xml, doc: back } = roundTrip(doc);
    expect(xml).toContain("<m:oMathPara");
    expect(equationLinear(paragraphEquations(para(back))[0] ?? buildEquation(""))).toBe("a^2+b^2=c^2");
  });
});

describe("drop caps", () => {
  it("moves the first letter into a dropCap frame and back", () => {
    const doc = createDocx({ paragraphs: ["Once upon a time"] });
    setDropCap(doc, para(doc), { position: "drop", lines: 3, distanceTwips: 144, font: "Georgia" });
    expect(paragraphs(doc)).toHaveLength(2);
    expect(textOf(para(doc, 0))).toBe("O");
    expect(textOf(para(doc, 1))).toBe("nce upon a time");
    expect(getDropCap(doc, para(doc, 1))).toMatchObject({ position: "drop", lines: 3, distanceTwips: 144, font: "Georgia" });
    const { xml } = roundTrip(doc);
    expect(xml).toContain('<w:framePr w:dropCap="drop" w:lines="3" w:wrap="around" w:vAnchor="text" w:hAnchor="text" w:hSpace="144"/>');
    expect(xml).toContain('<w:rFonts w:ascii="Georgia" w:hAnsi="Georgia"/><w:position w:val="-8"/><w:sz w:val="111"/>');
    expect(xml).toContain('<w:spacing w:after="0" w:line="827" w:lineRule="exact"/>');

    setDropCap(doc, para(doc, 1), { position: "margin", lines: 2 });
    expect(getDropCap(doc, para(doc, 1)).position).toBe("margin");
    setDropCap(doc, para(doc, 1), { position: "none" });
    expect(paragraphs(doc)).toHaveLength(1);
    expect(textOf(para(doc, 0))).toBe("Once upon a time");
    roundTrip(doc);
  });
});

// 1×1 transparent PNG.
const PNG = Uint8Array.from(
  atob("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGNgYGBgAAAABQABpfZFQAAAAABJRU5ErkJggg=="),
  (c) => c.charCodeAt(0),
);

describe("insertSignatureLine", () => {
  it("writes a VML picture with o:signatureline", () => {
    const doc = createDocx({ paragraphs: [""] });
    insertSignatureLine(doc, para(doc), 0, {
      image: PNG,
      widthPt: 192,
      heightPt: 96,
      suggestedSigner: "Ada Lovelace",
      suggestedSignerTitle: "CEO",
      allowComments: true,
    });
    const { xml } = roundTrip(doc);
    expect(xml).toContain('<w:pict xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"');
    expect(xml).toMatch(/<o:signatureline v:ext="edit" id="\{[0-9A-F-]{36}\}" provid="\{0{8}-0{4}-0{4}-0{4}-0{12}\}" o:suggestedsigner="Ada Lovelace" o:suggestedsigner2="CEO" allowcomments="t" issignatureline="t"\/>/);
    expect(xml).toMatch(/<v:imagedata r:id="rId\d+" o:title=""\/>/);
  });
});

describe("cover pages", () => {
  it("inserts, replaces and removes a cover page", () => {
    const doc = createDocx({ paragraphs: ["Body"] });
    const cover = (title: string): WmlParagraph => ({
      kind: "paragraph",
      children: [{ kind: "run", pieces: [{ kind: "text", value: title, preserveSpace: false }], extras: [] }],
      extras: [],
    });
    insertCoverPage(doc, [cover("Title A")]);
    insertCoverPage(doc, [cover("Title B"), cover("Subtitle")]);
    expect(paragraphs(doc).map(textOf)).toEqual(["Title B", "Subtitle", "", "Body"]);
    expect(hasCoverPage(doc)).toBe(true);
    const { xml, doc: back } = roundTrip(doc);
    expect(xml).toContain('w:name="_CoverPage"');
    expect(xml).toContain("<w:titlePg/>");
    expect(removeCoverPage(back)).toBe(true);
    expect(paragraphs(back).map(textOf)).toEqual(["Body"]);
    expect(removeCoverPage(back)).toBe(false);
  });
});

describe("headers, footers and page numbers", () => {
  it("replaces header content and removes it again", () => {
    const doc = createDocx({ paragraphs: ["Body"] });
    const content: WmlParagraph = {
      kind: "paragraph",
      children: [{ kind: "run", pieces: [{ kind: "text", value: "Left", preserveSpace: false }, { kind: "tab" }, { kind: "text", value: "Mid", preserveSpace: false }], extras: [] }],
      extras: [],
    };
    setHeaderFooterParagraphs(doc, "header", "default", [content]);
    setHeaderFooterParagraphs(doc, "header", "first", [content]);
    let back = roundTrip(doc).doc;
    expect(headers(back)).toHaveLength(2);
    expect(headerFooterParagraphs(back, "header", "default")?.map(textOf)).toEqual(["LeftMid"]);
    expect(partXml(back, "/word/document.xml")).toContain("<w:titlePg/>");
    expect(partXml(back, "/word/styles.xml")).toContain('w:styleId="Header"');

    expect(removeHeaderFooter(back, "header", "first")).toBe(true);
    back = roundTrip(back).doc;
    expect(headers(back)).toHaveLength(1);
  });

  it("puts page numbers in the footer or the margin and removes them", () => {
    const doc = createDocx({ paragraphs: ["Body"] });
    insertPageNumbers(doc, { position: "bottom", align: "center", style: "pageXofY" });
    insertPageNumbers(doc, { position: "margin", align: "right" });
    let back = roundTrip(doc).doc;
    const footerXml = partXml(back, footers(back)[0]?.partName ?? "");
    expect(footerXml).toContain('<w:jc w:val="center"/>');
    expect(footerXml).toContain(" PAGE ");
    expect(footerXml).toContain(" NUMPAGES ");
    const headerXml = partXml(back, headers(back)[0]?.partName ?? "");
    expect(headerXml).toContain('<w:framePr w:wrap="around" w:vAnchor="page" w:hAnchor="page" w:xAlign="right" w:yAlign="center"/>');

    expect(removePageNumbers(back)).toBe(3);
    back = roundTrip(back).doc;
    expect(partXml(back, footers(back)[0]?.partName ?? "")).not.toContain("PAGE");
    expect(partXml(back, headers(back)[0]?.partName ?? "")).not.toContain("framePr");
  });

  it("writes the page number format in schema order", () => {
    const doc = createDocx({ paragraphs: ["Body"] });
    setPageNumberFormat(doc, { format: "lowerRoman", start: 3, chapterStyle: 1, chapterSeparator: "emDash" });
    const { xml, doc: back } = roundTrip(doc);
    expect(xml).toContain('<w:pgNumType w:fmt="lowerRoman" w:start="3" w:chapStyle="1" w:chapSep="emDash"/>');
    expect(getPageNumberFormat(back)).toEqual({ format: "lowerRoman", start: 3, chapterStyle: 1, chapterSeparator: "emDash" });
    setPageNumberFormat(back, {});
    expect(roundTrip(back).xml).not.toContain("pgNumType");
  });
});

describe("insertDocumentContent", () => {
  it("copies paragraphs with their styles, lists, links, pictures and notes", () => {
    const source = createDocx({ paragraphs: [] });
    ensureHeadingStyles(source);
    appendHeading(source, "Imported", 1);
    const p = appendParagraph(source, "See ");
    insertHyperlink(source, p, 4, "site", { url: "https://example.org/" });
    addFootnote(source, p, "Note text.");
    insertSignatureLine(source, p, 0, { image: PNG, widthPt: 10, heightPt: 10 });
    insertBookmark(source, "Shared", { paragraph: p, offset: 0 });
    const target = createDocx({ paragraphs: ["Before", "After"] });
    addFootnote(target, para(target, 0), "Existing note.");
    insertBookmark(target, "Shared", { paragraph: para(target, 1), offset: 0 });

    const n = insertDocumentContent(target, openDocx(toUint8Array(source)), 1);
    expect(n).toBe(2);
    const { doc: back, xml } = roundTrip(target);
    expect(paragraphs(back).map(textOf)).toEqual(["Before", "Imported", "See ", "After"]);
    expect(partXml(back, "/word/styles.xml")).toContain('w:styleId="Heading1"');
    expect(xml).toMatch(/<w:footnoteReference w:id="2"\/>/);
    expect(partXml(back, "/word/footnotes.xml")).toContain("Note text.");
    expect(xml.match(/w:name="Shared"/g)).toHaveLength(1);
    expect(paragraphHyperlinks(back, para(back, 2))[0]?.url).toBe("https://example.org/");
  });

  it("rejects an index outside the body", () => {
    const doc = createDocx({ paragraphs: ["x"] });
    expect(() => insertDocumentContent(doc, createDocx(), 5)).toThrow(RangeError);
  });
});
