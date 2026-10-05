import { describe, expect, it } from "vitest";
import {
  appendHeading,
  appendPageBreak,
  appendParagraph,
  createDocx,
  ensureHeadingStyles,
  openDocx,
  paragraphs,
  setPageSize,
  type Docx,
  toUint8Array,
} from "./docx.js";
import {
  addCaptionLabel,
  bibliographySources,
  bibliographyStyle,
  captionLabels,
  convertNotes,
  insertBibliography,
  insertCaption,
  insertCitation,
  insertIndex,
  insertNote,
  insertTableOfAuthorities,
  insertTableOfContents,
  markAllIndexEntries,
  markAuthorityCitation,
  markIndexEntry,
  noteMarks,
  noteProperties,
  noteText,
  removeTableOfContents,
  setBibliographySources,
  setBibliographyStyle,
  setNoteProperties,
  setNoteText,
  setTocLevel,
  suggestSourceTag,
  tableOfContentsInstruction,
  tocLevel,
  updateTables,
  type BibliographySource,
} from "./references.js";
import { PAGE_SIZE_A4 } from "../internal/wordprocessingml/index.js";
import { validatePackage } from "./validator.js";

function partXml(doc: Docx, name: string): string {
  const reopened = openDocx(toUint8Array(doc));
  return new TextDecoder().decode(reopened.opc.parts.get(name)?.data ?? new Uint8Array());
}

function bodyXml(doc: Docx): string {
  return partXml(doc, "/word/document.xml");
}

/** The visible text of each paragraph (field codes dropped), one string per paragraph. */
function paragraphTexts(xml: string): string[] {
  return [...xml.matchAll(/<w:p[ >].*?<\/w:p>/g)].map((m) =>
    [
      ...m[0]
        .replace(/<w:instrText[^>]*>.*?<\/w:instrText>/g, "")
        .matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>|<w:tab\/>/g),
    ]
      .map((t) => (t[1] ?? "\t").replace(/&quot;/g, '"').replace(/&amp;/g, "&"))
      .join(""),
  );
}

function roundTrip(doc: Docx): Docx {
  const reopened = openDocx(toUint8Array(doc));
  expect(validatePackage(reopened.opc)).toEqual([]);
  return reopened;
}

function headingDoc(): Docx {
  const doc = createDocx({ paragraphs: [] });
  ensureHeadingStyles(doc, 3);
  appendHeading(doc, "Introduction", 1);
  appendParagraph(doc, "Body text.");
  appendHeading(doc, "Background", 2);
  appendPageBreak(doc);
  appendHeading(doc, "Method", 1);
  appendHeading(doc, "Deep detail", 3);
  return doc;
}

describe("table of contents", () => {
  it("builds the Word instruction from the dialog options", () => {
    expect(tableOfContentsInstruction()).toBe('TOC \\o "1-3" \\h \\z \\u');
    expect(
      tableOfContentsInstruction({
        levels: { from: 1, to: 2 },
        pageNumbers: false,
        hyperlinks: false,
      }),
    ).toBe('TOC \\o "1-2" \\z \\u \\n');
    expect(tableOfContentsInstruction({ captionLabel: "Figure" })).toBe('TOC \\h \\z \\c "Figure"');
    expect(() => tableOfContentsInstruction({ levels: { from: 3, to: 1 } })).toThrow(/Invalid TOC/);
  });

  it("computes entries, page numbers and bookmarks, and round-trips", () => {
    const doc = headingDoc();
    insertTableOfContents(doc, 0, { title: "Contents" });
    const reopened = roundTrip(doc);
    const xml = bodyXml(reopened);
    expect(xml).toContain('TOC \\o "1-3" \\h \\z \\u');
    expect(xml).toContain('<w:pStyle w:val="TOCHeading"/>');
    expect(xml).toContain('<w:pStyle w:val="TOC1"/>');
    expect(xml).toContain('<w:pStyle w:val="TOC3"/>');
    // Method is after the page break: page 2.
    const method = xml.indexOf(">Method<");
    const pageref = xml.indexOf("PAGEREF", method);
    expect(xml.slice(pageref, pageref + 400)).toMatch(/<w:t>2<\/w:t>/);
    expect(xml).toMatch(/<w:bookmarkStart w:id="\d+" w:name="_Toc\d{9}"\/>/);
    expect(xml).toContain('w:leader="dot"');
    const styles = partXml(reopened, "/word/styles.xml");
    expect(styles).toContain('w:styleId="TOC2"');
  });

  it("ends the leader at the right margin, with Word's 1 in default margins", () => {
    const doc = headingDoc();
    // A4 width without w:pgMar: Word lays the section out with 1 in margins.
    setPageSize(doc, PAGE_SIZE_A4);
    const sectPr = doc.document.body.sectPr;
    if (sectPr) {
      sectPr.children = sectPr.children.filter(
        (c) => !(c.kind === "element" && c.name.local === "pgMar"),
      );
    }
    insertTableOfContents(doc, 0, {});
    expect(bodyXml(doc)).toContain(`w:pos="${11906 - 2 * 1440}"`);
  });

  it("uses the page-number provider when given", () => {
    const doc = headingDoc();
    insertTableOfContents(doc, 0, { pageOf: () => 7 });
    expect(paragraphTexts(bodyXml(doc)).slice(0, 4)).toEqual([
      "Introduction\t7",
      "Background\t7",
      "Method\t7",
      "Deep detail\t7",
    ]);
  });

  it("updates after headings change and keeps one bookmark per heading", () => {
    const doc = headingDoc();
    insertTableOfContents(doc, 0, {});
    appendHeading(doc, "Conclusion", 1);
    expect(updateTables(doc, { types: ["TOC"] })).toBe(1);
    updateTables(doc);
    const xml = bodyXml(doc);
    expect(xml).toContain(">Conclusion<");
    const marks = xml.match(/w:name="_Toc\d+"/g) ?? [];
    expect(marks.length).toBe(5);
  });

  it("Add Text excludes or promotes paragraphs", () => {
    const doc = headingDoc();
    const background = paragraphs(doc)[2];
    const body = paragraphs(doc)[1];
    if (!background || !body) throw new Error("fixture");
    setTocLevel(doc, background, "none");
    setTocLevel(doc, body, 2);
    expect(tocLevel(doc, background)).toBe("none");
    expect(tocLevel(doc, body)).toBe(2);
    insertTableOfContents(doc, 0, {});
    const xml = bodyXml(doc);
    expect(paragraphTexts(xml).slice(0, 4)).toEqual([
      "Introduction\t1",
      "Body text.\t1",
      "Method\t2",
      "Deep detail\t2",
    ]);
    expect(xml).toContain('<w:outlineLvl w:val="9"/>');
  });

  it("removes the table and its title", () => {
    const doc = headingDoc();
    const before = doc.document.body.blocks.length;
    insertTableOfContents(doc, 0, { title: "Contents" });
    expect(removeTableOfContents(doc)).toBe(1);
    expect(doc.document.body.blocks.length).toBe(before);
    expect(bodyXml(doc)).not.toContain("TOC \\o");
  });

  it("reports when there are no entries", () => {
    const doc = createDocx({ paragraphs: ["plain"] });
    insertTableOfContents(doc, 0, {});
    expect(bodyXml(doc)).toContain("No table of contents entries found.");
  });
});

describe("captions and table of figures", () => {
  it("numbers captions in document order and renumbers later ones", () => {
    const doc = createDocx({ paragraphs: ["a", "b", "c"] });
    insertCaption(doc, 2, "below", { label: "Figure", text: ": Last" });
    insertCaption(doc, 0, "below", { label: "Figure", text: ": First" });
    const xml = bodyXml(roundTrip(doc));
    expect(xml).toContain("SEQ Figure \\* ARABIC");
    const first = xml.indexOf(": First");
    const last = xml.indexOf(": Last");
    expect(xml.slice(xml.lastIndexOf("SEQ", first), first)).toContain("<w:t>1</w:t>");
    expect(xml.slice(xml.lastIndexOf("SEQ", last), last)).toContain("<w:t>2</w:t>");
    expect(partXml(doc, "/word/styles.xml")).toContain('w:styleId="Caption"');
  });

  it("supports formats, excluded labels and chapter numbers", () => {
    const doc = createDocx({ paragraphs: [] });
    ensureHeadingStyles(doc, 1);
    appendHeading(doc, "One", 1);
    appendParagraph(doc, "x");
    appendHeading(doc, "Two", 1);
    appendParagraph(doc, "y");
    insertCaption(doc, 3, "below", {
      label: "Table",
      numberFormat: "upperLetter",
      chapter: { headingLevel: 1, separator: "-" },
    });
    insertCaption(doc, 1, "above", { label: "Equation", excludeLabel: true });
    const xml = bodyXml(doc);
    expect(xml).toContain("SEQ Table \\* ALPHABETIC \\s 1");
    expect(xml).toMatch(/STYLEREF 1 \\s.*?<w:t>2<\/w:t>.*?<w:t>-<\/w:t>.*?<w:t>A<\/w:t>/);
    expect(xml).not.toContain(">Equation <");
  });

  it("builds a table of figures from captions", () => {
    const doc = createDocx({ paragraphs: ["pic"] });
    insertCaption(doc, 0, "below", { label: "Figure", text: ": A chart" });
    insertTableOfContents(doc, 0, { captionLabel: "Figure" });
    const xml = bodyXml(roundTrip(doc));
    expect(xml).toContain('TOC \\h \\z \\c "Figure"');
    expect(xml).toMatch(/TableofFigures.*?Figure 1: A chart/);
    insertTableOfContents(doc, 0, { captionLabel: "Figure", includeLabelAndNumber: false });
    const texts = paragraphTexts(bodyXml(doc));
    expect(bodyXml(doc)).toContain('TOC \\h \\z \\a "Figure"');
    expect(texts[0]).toBe("A chart\t1");
  });

  it("stores new labels in settings", () => {
    const doc = createDocx();
    addCaptionLabel(doc, "Photo");
    addCaptionLabel(doc, "photo");
    expect(captionLabels(doc)).toEqual(["Figure", "Table", "Equation", "Photo"]);
    expect(partXml(roundTrip(doc), "/word/settings.xml")).toContain(
      '<w:caption w:name="Photo" w:pos="below"/>',
    );
  });
});

describe("index", () => {
  it("marks entries and builds an indented two-column index", () => {
    const doc = createDocx({ paragraphs: ["Apples are red.", "Pears and apples."] });
    appendPageBreak(doc);
    appendParagraph(doc, "More apples.");
    const [p1, p2] = paragraphs(doc);
    if (!p1 || !p2) throw new Error("fixture");
    markIndexEntry(doc, p1, 6, { main: "Apple", subentry: "red", bold: true });
    markIndexEntry(doc, p2, 5, { main: "Pear" });
    markIndexEntry(doc, p2, 5, { main: "Fruit", crossReference: "Apple" });
    expect(markAllIndexEntries(doc, "apples", { main: "Apple" })).toBe(2);
    insertIndex(doc, doc.document.body.blocks.length, { headings: true });
    const reopened = roundTrip(doc);
    const xml = bodyXml(reopened);
    expect(xml).toContain('XE "Apple:red" \\b');
    expect(xml).toContain('XE "Fruit" \\t "Apple"');
    expect(xml).toContain('INDEX \\h "A" \\c "2" \\z "1033"');
    const texts = paragraphTexts(xml);
    const start = texts.indexOf("A");
    expect(texts.slice(start, start + 6)).toEqual([
      "A",
      "Apple, 1, 2",
      "red, 1",
      "F",
      "Fruit. See Apple",
      "P",
    ]);
    expect(xml).toMatch(/<w:r><w:rPr><w:b\/><w:bCs\/><\/w:rPr><w:t>1<\/w:t>/);
    expect(xml).toContain('<w:cols w:num="2" w:space="720"/>');
    expect(xml).toContain('<w:type w:val="continuous"/>');
  });

  it("run-in index with right-aligned numbers", () => {
    const doc = createDocx({ paragraphs: ["x"] });
    const [p] = paragraphs(doc);
    if (!p) throw new Error("fixture");
    markIndexEntry(doc, p, 1, { main: "Main", subentry: "Sub" });
    insertIndex(doc, 1, { type: "runIn", columns: 1, rightAlignPageNumbers: true });
    const xml = bodyXml(doc);
    expect(xml).toContain('\\e "\t" \\r \\c "1"');
    expect(paragraphTexts(xml)).toContain("Main: Sub\t1");
    expect(xml).not.toContain("w:cols");
  });
});

describe("table of authorities", () => {
  it("lists marked citations per category with passim", () => {
    const doc = createDocx({ paragraphs: ["See Roe.", "Again."] });
    const [p1, p2] = paragraphs(doc);
    if (!p1 || !p2) throw new Error("fixture");
    markAuthorityCitation(doc, p1, 3, {
      longCitation: "Roe v. Wade, 410 U.S. 113 (1973)",
      shortCitation: "Roe",
      category: 1,
    });
    markAuthorityCitation(doc, p2, 0, { longCitation: "U.S. Const. art. I", category: 7 });
    insertTableOfAuthorities(doc, 2, { category: "all" });
    const xml = bodyXml(roundTrip(doc));
    expect(xml).toContain('TA \\l "Roe v. Wade, 410 U.S. 113 (1973)" \\s "Roe" \\c 1');
    expect(xml).toContain('TOA \\h \\c "1" \\p');
    expect(xml).toContain('TOA \\h \\c "7" \\p');
    const texts = paragraphTexts(xml);
    expect(texts.slice(2)).toEqual([
      "Cases",
      "Roe v. Wade, 410 U.S. 113 (1973)\t1",
      "",
      "Constitutional Provisions",
      "U.S. Const. art. I\t1",
      "",
    ]);
  });
});

describe("footnotes and endnotes", () => {
  it("inserts notes at the caret with Word's styles and numbers them", () => {
    const doc = createDocx({ paragraphs: ["Hello world"] });
    const [p] = paragraphs(doc);
    if (!p) throw new Error("fixture");
    insertNote(doc, p, 5, "footnote", { text: "First." });
    insertNote(doc, p, 11, "footnote", { customMark: "*" });
    insertNote(doc, p, 0, "endnote", { text: "End." });
    const reopened = roundTrip(doc);
    const xml = bodyXml(reopened);
    expect(xml).toMatch(
      /<w:t>Hello<\/w:t><\/w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"\/><\/w:rPr><w:footnoteReference w:id="1"\/>/,
    );
    expect(xml).toContain('<w:footnoteReference w:customMarkFollows="1" w:id="2"/><w:t>*</w:t>');
    expect(partXml(reopened, "/word/endnotes.xml")).toContain('<w:pStyle w:val="EndnoteText"/>');
    expect(noteMarks(reopened, "footnote").map((m) => m.mark)).toEqual(["1", "*"]);
    expect(noteMarks(reopened, "endnote").map((m) => m.mark)).toEqual(["i"]);
  });

  it("reads and replaces a note's text, keeping its mark", () => {
    const doc = createDocx({ paragraphs: ["Hello"] });
    const [p] = paragraphs(doc);
    if (!p) throw new Error("fixture");
    const id = insertNote(doc, p, 5, "footnote", { text: "Draft." });
    expect(noteText(doc, "footnote", id)).toBe("Draft.");
    setNoteText(doc, "footnote", id, "Final line one.\nLine two.");
    const reopened = roundTrip(doc);
    expect(noteText(reopened, "footnote", id)).toBe("Final line one.\nLine two.");
    const notes = partXml(reopened, "/word/footnotes.xml");
    expect(notes).toContain(
      '<w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> Final line one.</w:t></w:r></w:p>',
    );
    expect(notes.match(/<w:pStyle w:val="FootnoteText"\/>/g)).toHaveLength(2);
  });

  it("writes footnotePr to settings and sections and honours it in marks", () => {
    const doc = createDocx({ paragraphs: ["a"] });
    const [p] = paragraphs(doc);
    if (!p) throw new Error("fixture");
    insertNote(doc, p, 1, "footnote");
    insertNote(doc, p, 1, "footnote");
    setNoteProperties(doc, "footnote", {
      numberFormat: "upperRoman",
      startAt: 3,
      position: "beneathText",
    });
    expect(noteProperties(doc, "footnote")).toEqual({
      position: "beneathText",
      numberFormat: "upperRoman",
      startAt: 3,
      restart: "continuous",
    });
    expect(noteMarks(doc, "footnote").map((m) => m.mark)).toEqual(["III", "IV"]);
    const reopened = roundTrip(doc);
    expect(partXml(reopened, "/word/settings.xml")).toContain(
      '<w:footnotePr><w:pos w:val="beneathText"/><w:numFmt w:val="upperRoman"/><w:numStart w:val="3"/></w:footnotePr>',
    );
    expect(bodyXml(reopened)).toMatch(
      /<w:sectPr><w:footnotePr><w:pos w:val="beneathText"\/><w:numFmt w:val="upperRoman"\/><w:numStart w:val="3"\/><\/w:footnotePr><w:pgSz/,
    );
    expect(() => setNoteProperties(doc, "endnote", { position: "pageBottom" })).toThrow();
  });

  it("converts footnotes to endnotes and swaps", () => {
    const doc = createDocx({ paragraphs: ["a"] });
    const [p] = paragraphs(doc);
    if (!p) throw new Error("fixture");
    insertNote(doc, p, 1, "footnote", { text: "note" });
    expect(convertNotes(doc, "footnotesToEndnotes")).toBe(1);
    let reopened = roundTrip(doc);
    expect(bodyXml(reopened)).toContain('<w:endnoteReference w:id="1"/>');
    expect(partXml(reopened, "/word/endnotes.xml")).toMatch(
      /<w:endnote w:id="1"><w:p><w:pPr><w:pStyle w:val="EndnoteText"\/><\/w:pPr><w:r><w:rPr><w:rStyle w:val="EndnoteReference"\/><\/w:rPr><w:endnoteRef\/>/,
    );
    insertNote(reopened, paragraphs(reopened)[0] as never, 0, "footnote", { text: "other" });
    expect(convertNotes(reopened, "swap")).toBe(2);
    reopened = roundTrip(reopened);
    expect(noteMarks(reopened, "footnote")).toHaveLength(1);
    expect(noteMarks(reopened, "endnote")).toHaveLength(1);
    expect(partXml(reopened, "/word/footnotes.xml")).toContain("note");
  });
});

const BOOK: BibliographySource = {
  tag: "Smi20",
  type: "Book",
  authors: [{ last: "Smith", first: "John", middle: "Paul" }],
  fields: { Title: "A Book", Year: "2020", City: "Boston", Publisher: "Acme" },
};
const ARTICLE: BibliographySource = {
  tag: "Doe19",
  type: "JournalArticle",
  authors: [
    { last: "Doe", first: "Jane" },
    { last: "Roe", first: "Rick" },
  ],
  fields: {
    Title: "An Article",
    Year: "2019",
    JournalName: "Journal",
    Volume: "3",
    Issue: "2",
    Pages: "1-9",
  },
};

describe("citations and bibliography", () => {
  it("stores sources in a customXml part per §22.6 and round-trips them", () => {
    const doc = createDocx({ paragraphs: ["Claim."] });
    setBibliographySources(doc, [BOOK, ARTICLE]);
    const reopened = roundTrip(doc);
    expect(bibliographySources(reopened)).toEqual([BOOK, ARTICLE]);
    expect(bibliographyStyle(reopened)).toBe("APA");
    const item = partXml(reopened, "/customXml/item1.xml");
    expect(item).toContain('StyleName="APA"');
    expect(item).toContain("<b:Tag>Smi20</b:Tag>");
    expect(partXml(reopened, "/customXml/itemProps1.xml")).toContain(
      'ds:uri="http://schemas.openxmlformats.org/officeDocument/2006/bibliography"',
    );
    expect(() => setBibliographySources(doc, [BOOK, BOOK])).toThrow(/Duplicate/);
  });

  it("formats citations and the bibliography in APA and MLA", () => {
    const doc = createDocx({ paragraphs: ["Claim one.", "Claim two."] });
    setBibliographySources(doc, [BOOK, ARTICLE]);
    const [p1, p2] = paragraphs(doc);
    if (!p1 || !p2) throw new Error("fixture");
    insertCitation(doc, p1, 9, "Smi20", { pages: "12" });
    insertCitation(doc, p2, 9, "Doe19");
    insertBibliography(doc, 2, { title: "References" });
    let xml = bodyXml(roundTrip(doc));
    expect(xml).toContain("CITATION Smi20 \\l 1033 \\p 12");
    expect(paragraphTexts(xml)).toEqual([
      "Claim one(Smith, 2020, p. 12).",
      "Claim two(Doe & Roe, 2019).",
      "References",
      "Doe, J., & Roe, R. (2019). An Article. Journal, 3(2), 1-9.",
      "Smith, J. P. (2020). A Book. Boston: Acme.",
      "",
    ]);
    expect(xml).toContain("<w:i/><w:iCs/></w:rPr><w:t>Journal, 3</w:t>");
    expect(xml).toContain("<w:i/><w:iCs/></w:rPr><w:t>A Book.</w:t>");
    setBibliographyStyle(doc, "MLA");
    xml = bodyXml(doc);
    expect(paragraphTexts(xml).slice(0, 5)).toEqual([
      "Claim one(Smith 12).",
      "Claim two(Doe and Roe).",
      "References",
      'Doe, Jane and Rick Roe. "An Article." Journal 3.2 (2019): 1-9. Print.',
      "Smith, John Paul. A Book. Boston: Acme, 2020. Print.",
    ]);
    setBibliographyStyle(doc, "IEEE");
    xml = bodyXml(doc);
    expect(paragraphTexts(xml).slice(0, 2)).toEqual(["Claim one[1, 12].", "Claim two[2]."]);
    expect(partXml(doc, "/customXml/item1.xml")).toContain(
      'SelectedStyle="\\IEEE2006OfficeOnline.xsl"',
    );
  });

  it("suggests Word-style tags", () => {
    expect(
      suggestSourceTag([BOOK], {
        type: "Book",
        authors: [{ last: "Smith" }],
        fields: { Year: "2020" },
      }),
    ).toBe("Smi201");
    expect(suggestSourceTag([], { type: "Misc" })).toBe("Src");
  });
});
