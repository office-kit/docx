import { describe, expect, it } from "vitest";
import { getPart, writeOpcPackage } from "../internal/opc/index.js";
import {
  addComment,
  createDocx,
  type Docx,
  openDocx,
  paragraphs,
  setCoreProperties,
  text,
  toUint8Array,
} from "./docx.js";
import {
  acceptRevisions,
  checkAccessibility,
  comments,
  contrastRatio,
  deleteTrackedText,
  getRunLanguage,
  insertTrackedText,
  paragraphMarkRevision,
  rejectRevisions,
  removeComment,
  revisions,
  setCommentText,
  setRunLanguage,
  trackParagraphFormatChange,
  trackParagraphMark,
  trackRunFormatChange,
  wordCount,
} from "./review.js";
import { validatePackage } from "./validator.js";
import {
  acceptAllRevisions,
  rejectAllRevisions,
  setParagraphValProp,
  setRunOnOff,
} from "./index.js";

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const WHO = { author: "Ann", date: "2026-10-05T00:00:00Z" };

function docWithBody(body: string): Docx {
  const seed = createDocx({ paragraphs: [] });
  const part = getPart(seed.opc, "/word/document.xml");
  if (!part) throw new Error("no document part");
  part.data = new TextEncoder().encode(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${body}</w:body></w:document>`,
  );
  return openDocx(writeOpcPackage(seed.opc));
}

function documentXml(doc: Docx): string {
  const reopened = openDocx(toUint8Array(doc));
  const part = getPart(reopened.opc, "/word/document.xml");
  return new TextDecoder().decode(part?.data);
}

function firstParagraph(doc: Docx) {
  const para = paragraphs(doc)[0];
  if (!para) throw new Error("no paragraph");
  return para;
}

describe("run revisions in the AST", () => {
  it("lifts <w:ins>/<w:del> onto runs and writes them back unchanged", () => {
    const body =
      '<w:p><w:r><w:t>a</w:t></w:r><w:ins w:id="1" w:author="X" w:date="2026-01-01T00:00:00Z"><w:r><w:t>b</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>c</w:t></w:r></w:ins><w:del w:id="2" w:author="X"><w:r><w:delText>d</w:delText></w:r></w:del></w:p>';
    const doc = docWithBody(body);
    const runs = firstParagraph(doc).children;
    expect(runs.map((r) => r.kind)).toEqual(["run", "run", "run", "run"]);
    expect(runs[1]?.kind === "run" && runs[1].revision?.kind).toBe("ins");
    expect(documentXml(doc)).toContain(body);
  });

  it("keeps a wrapper holding more than runs as raw", () => {
    const body =
      '<w:p><w:ins w:id="1" w:author="X"><w:bookmarkStart w:id="5" w:name="b"/><w:r><w:t>x</w:t></w:r></w:ins></w:p>';
    const doc = docWithBody(body);
    expect(firstParagraph(doc).children[0]?.kind).toBe("raw");
    expect(documentXml(doc)).toContain(body);
  });
});

describe("listing and resolving revisions", () => {
  const body =
    '<w:p><w:pPr><w:rPr><w:ins w:id="7" w:author="B"/></w:rPr></w:pPr><w:r><w:rPr><w:b/><w:rPrChange w:id="3" w:author="B"><w:rPr/></w:rPrChange></w:rPr><w:t>bold</w:t></w:r><w:ins w:id="1" w:author="A"><w:r><w:t> new</w:t></w:r></w:ins><w:del w:id="2" w:author="A"><w:r><w:delText> old</w:delText></w:r></w:del></w:p><w:p><w:r><w:t>next</w:t></w:r></w:p>';

  it("lists every kind with author and text", () => {
    const list = revisions(docWithBody(body));
    expect(list.map((r) => [r.kind, r.id, r.author, r.text])).toEqual([
      ["format", "3", "B", ""],
      ["insert", "1", "A", " new"],
      ["delete", "2", "A", " old"],
      ["paragraphInsert", "7", "B", ""],
    ]);
  });

  it("accepts one revision by id", () => {
    const doc = docWithBody(body);
    expect(acceptRevisions(doc, ["2"])).toBe(1);
    expect(text(doc)).toBe("bold new\nnext");
    expect(revisions(doc).map((r) => r.id)).toEqual(["3", "1", "7"]);
  });

  it("rejects a formatting change back to the old properties", () => {
    const doc = docWithBody(body);
    rejectRevisions(doc, ["3"]);
    expect(documentXml(doc)).toContain("<w:r><w:rPr></w:rPr><w:t>bold</w:t></w:r>");
  });

  it("rejecting an inserted paragraph mark joins the paragraphs", () => {
    const doc = docWithBody(body);
    rejectRevisions(doc, ["7"]);
    expect(paragraphs(doc)).toHaveLength(1);
    expect(text(doc)).toContain("next");
  });

  it("accept all / reject all resolve every kind", () => {
    const accepted = docWithBody(body);
    expect(acceptAllRevisions(accepted)).toBe(4);
    expect(revisions(accepted)).toEqual([]);
    expect(text(accepted)).toBe("bold new\nnext");
    const rejected = docWithBody(body);
    rejectAllRevisions(rejected);
    expect(text(rejected)).toBe("bold oldnext");
  });

  it("resolves table row insertions and deletions", () => {
    const doc = docWithBody(
      '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>keep</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:trPr><w:ins w:id="4" w:author="A"/></w:trPr><w:tc><w:p><w:r><w:t>added</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>',
    );
    expect(revisions(doc).map((r) => r.kind)).toEqual(["rowInsert"]);
    rejectRevisions(doc, ["4"]);
    expect(text(doc)).not.toContain("added");
  });
});

describe("recording tracked changes", () => {
  it("records typing as one insertion that grows", () => {
    const doc = createDocx({ paragraphs: ["Hello world"] });
    const para = firstParagraph(doc);
    let at = insertTrackedText(doc, para, 5, ",", WHO);
    at = insertTrackedText(doc, para, at, " dear", WHO);
    expect(at).toBe(11);
    expect(text(doc)).toBe("Hello, dear world");
    const list = revisions(doc);
    expect(list).toHaveLength(1);
    expect(list[0]).toMatchObject({ kind: "insert", author: "Ann", text: ", dear" });
    expect(documentXml(doc)).toContain(
      '<w:ins w:id="0" w:author="Ann" w:date="2026-10-05T00:00:00Z"><w:r><w:t>, dear</w:t></w:r></w:ins>',
    );
    expect(validatePackage(openDocx(toUint8Array(doc)).opc)).toEqual([]);
  });

  it("marks deleted text and removes the author's own insertion outright", () => {
    const doc = createDocx({ paragraphs: ["abcdef"] });
    const para = firstParagraph(doc);
    insertTrackedText(doc, para, 3, "XY", WHO);
    expect(deleteTrackedText(doc, para, 2, 6, WHO)).toBe(4);
    // "c" and "d" are deleted; "XY" (just typed) is gone.
    expect(revisions(doc).map((r) => [r.kind, r.text])).toEqual([["delete", "cd"]]);
    expect(documentXml(doc)).toContain("<w:delText>c</w:delText>");
    rejectAllRevisions(doc);
    expect(text(doc)).toBe("abcdef");
  });

  it("types after deleted text, as Word does", () => {
    const doc = createDocx({ paragraphs: ["abc"] });
    const para = firstParagraph(doc);
    deleteTrackedText(doc, para, 2, 3, WHO);
    insertTrackedText(doc, para, 2, "x", WHO);
    const order = para.children.map((c) => (c.kind === "run" ? (c.revision?.kind ?? "-") : "raw"));
    expect(order).toEqual(["-", "del", "ins"]);
  });

  it("records paragraph marks first in the mark's rPr", () => {
    const doc = createDocx({ paragraphs: ["a"] });
    const para = firstParagraph(doc);
    trackParagraphMark(doc, para, "ins", WHO);
    trackParagraphMark(doc, para, "del", WHO);
    expect(paragraphMarkRevision(para)).toEqual({ kind: "del", author: "Ann" });
    expect(documentXml(doc)).toMatch(
      /<w:pPr><w:rPr><w:ins [^>]*\/><w:del [^>]*\/><\/w:rPr><\/w:pPr>/,
    );
  });

  it("records a run format change with the original properties, last in rPr", () => {
    const doc = createDocx({ paragraphs: ["a"] });
    const run = firstParagraph(doc).children[0];
    if (run?.kind !== "run") throw new Error("no run");
    const before = structuredClone(run.rPr);
    setRunOnOff(run, "b", true);
    trackRunFormatChange(doc, run, before, WHO);
    expect(documentXml(doc)).toContain(
      '<w:rPr><w:b/><w:rPrChange w:id="0" w:author="Ann" w:date="2026-10-05T00:00:00Z"><w:rPr/></w:rPrChange></w:rPr>',
    );
    // Undoing the change by hand drops the record.
    const again = structuredClone(run.rPr);
    setRunOnOff(run, "b", false);
    trackRunFormatChange(doc, run, again, WHO);
    expect(documentXml(doc)).not.toContain("rPrChange");
  });

  it("records a paragraph format change", () => {
    const doc = createDocx({ paragraphs: ["a"] });
    const para = firstParagraph(doc);
    const before = structuredClone(para.pPr);
    setParagraphValProp(para, "jc", "center");
    trackParagraphFormatChange(doc, para, before, WHO);
    expect(documentXml(doc)).toContain('<w:jc w:val="center"/><w:pPrChange');
    rejectAllRevisions(doc);
    expect(documentXml(doc)).not.toContain("center");
  });
});

describe("comments", () => {
  it("anchors a comment to a character range and lists it", () => {
    const doc = createDocx({ paragraphs: ["one two three"] });
    const para = firstParagraph(doc);
    const id = addComment(doc, para, {
      author: "Ann",
      initials: "A",
      text: "Why?",
      range: { start: 4, end: 7 },
    });
    const xml = documentXml(doc);
    expect(xml).toMatch(
      /<w:t xml:space="preserve">one <\/w:t><\/w:r><w:commentRangeStart w:id="0"\/><w:r><w:t>two<\/w:t><\/w:r><w:commentRangeEnd w:id="0"\/>/,
    );
    expect(comments(doc)).toEqual([{ id, author: "Ann", initials: "A", text: "Why?", block: 0 }]);
    expect(setCommentText(doc, id, "Because")).toBe(true);
    expect(comments(openDocx(toUint8Array(doc)))[0]?.text).toBe("Because");
    expect(validatePackage(openDocx(toUint8Array(doc)).opc)).toEqual([]);
  });

  it("removes one comment with its anchors", () => {
    const doc = createDocx({ paragraphs: ["a", "b"] });
    const [p1, p2] = paragraphs(doc);
    if (!p1 || !p2) throw new Error("paragraphs");
    addComment(doc, p1, { author: "A", text: "first" });
    const second = addComment(doc, p2, { author: "A", text: "second" });
    expect(removeComment(doc, second)).toBe(true);
    const reopened = openDocx(toUint8Array(doc));
    expect(comments(reopened).map((c) => c.text)).toEqual(["first"]);
    expect(documentXml(doc)).not.toContain(`w:id="${second}"`);
    expect(validatePackage(reopened.opc)).toEqual([]);
  });
});

describe("proofing language", () => {
  it("writes w:lang with each script's language", () => {
    const doc = createDocx({ paragraphs: ["a"] });
    const run = firstParagraph(doc).children[0];
    if (run?.kind !== "run") throw new Error("no run");
    setRunLanguage(run, { latin: "en-GB", eastAsia: "ja-JP" });
    expect(getRunLanguage(run)).toEqual({ latin: "en-GB", eastAsia: "ja-JP" });
    expect(documentXml(doc)).toContain('<w:lang w:val="en-GB" w:eastAsia="ja-JP"/>');
    expect(() => setRunLanguage(run, { latin: "not a tag" })).toThrow(RangeError);
    setRunLanguage(run, {});
    expect(documentXml(doc)).not.toContain("w:lang");
  });
});

describe("wordCount", () => {
  it("counts like Word: East Asian characters are words, deleted text is not", () => {
    const doc = docWithBody(
      '<w:p><w:r><w:t xml:space="preserve">Hello, world  again</w:t></w:r></w:p><w:p/><w:p><w:r><w:t>日本語です</w:t></w:r><w:del w:id="1" w:author="A"><w:r><w:delText>gone</w:delText></w:r></w:del></w:p>',
    );
    expect(wordCount(doc)).toEqual({
      words: 3 + 5,
      characters: 16 + 5,
      charactersWithSpaces: 19 + 5,
      paragraphs: 2,
    });
  });
});

describe("checkAccessibility", () => {
  it("reports headers, heading order, contrast and blanks", () => {
    const doc = docWithBody(
      [
        '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Top</w:t></w:r></w:p>',
        '<w:p><w:pPr><w:pStyle w:val="Heading3"/></w:pPr><w:r><w:t>Deep</w:t></w:r></w:p>',
        '<w:p><w:r><w:rPr><w:color w:val="DDDDDD"/></w:rPr><w:t>faint</w:t></w:r></w:p>',
        "<w:p/><w:p/><w:p/>",
        "<w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>",
      ].join(""),
    );
    const kinds = checkAccessibility(doc).map((i) => [i.kind, i.severity, i.block]);
    expect(kinds).toEqual([
      ["skippedHeadingLevel", "warning", 1],
      ["lowContrast", "warning", 2],
      ["repeatedBlankParagraphs", "tip", 5],
      ["missingTableHeader", "error", 6],
      ["missingTitle", "tip", undefined],
    ]);
    setCoreProperties(doc, { title: "T" });
    expect(checkAccessibility(doc).some((i) => i.kind === "missingTitle")).toBe(false);
  });

  it("computes WCAG contrast", () => {
    expect(contrastRatio("000000", "FFFFFF")).toBeCloseTo(21);
    expect(contrastRatio("777777", "FFFFFF")).toBeCloseTo(4.48, 2);
  });
});
