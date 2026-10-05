import { describe, expect, it } from "vitest";
import { getPart } from "../internal/opc/index.js";
import {
  addFootnote,
  addHeader,
  appendParagraph,
  appendSectionBreak,
  createDocx,
  getRawPartRoot,
  markRawPartDirty,
  openDocx,
  paragraphText,
  setPageSize,
  splitParagraphAt,
  toUint8Array,
  validatePackage,
} from "./index.js";
import {
  ensureHeaderFooter,
  isHeaderFooterLinked,
  parseStoryKey,
  resolveHeaderFooter,
  sectionProperties,
  setHeaderFooterLinked,
  storyBody,
  storyKey,
  storyView,
} from "./story.js";

function partXml(bytes: Uint8Array, partName: string): string {
  const part = getPart(openDocx(bytes).opc, partName);
  return new TextDecoder().decode(part?.data);
}

function firstParagraphText(
  doc: ReturnType<typeof createDocx>,
  ref: Parameters<typeof storyBody>[1],
): string {
  const block = storyBody(doc, ref)?.blocks[0];
  return block?.kind === "paragraph" ? paragraphText(block) : "";
}

describe("storyBody", () => {
  it("parses a header part into typed blocks and saves edits to it", () => {
    const doc = createDocx({ paragraphs: ["Body"] });
    addHeader(doc, "Old header");
    const ref = { kind: "header", partName: "/word/header1.xml" } as const;
    const body = storyBody(doc, ref);
    const p = body?.blocks[0];
    expect(p?.kind).toBe("paragraph");
    if (p?.kind !== "paragraph") return;
    const run = p.children[0];
    if (run?.kind !== "run") throw new Error("expected a run");
    run.pieces = [{ kind: "text", value: "New header", preserveSpace: false }];

    const bytes = toUint8Array(doc);
    expect(partXml(bytes, "/word/header1.xml")).toContain("<w:t>New header</w:t>");
    expect(partXml(bytes, "/word/header1.xml")).toContain("<w:hdr");
    expect(validatePackage(openDocx(bytes).opc)).toEqual([]);
    expect(firstParagraphText(openDocx(bytes), ref)).toBe("New header");
  });

  it("parses footnotes and saves their edits", () => {
    const doc = createDocx({ paragraphs: ["Body"] });
    const para = doc.document.body.blocks[0];
    if (para?.kind !== "paragraph") throw new Error("expected a paragraph");
    const id = addFootnote(doc, para, "Note text");
    const ref = { kind: "footnote", id } as const;
    const body = storyBody(doc, ref);
    expect(body?.blocks.length).toBeGreaterThan(0);
    body?.blocks.push({
      kind: "paragraph",
      children: [
        {
          kind: "run",
          pieces: [{ kind: "text", value: "Second", preserveSpace: false }],
          extras: [],
        },
      ],
      extras: [],
    });
    const bytes = toUint8Array(doc);
    expect(partXml(bytes, "/word/footnotes.xml")).toContain("<w:t>Second</w:t>");
    expect(validatePackage(openDocx(bytes).opc)).toEqual([]);
    expect(storyBody(openDocx(bytes), ref)?.blocks.length).toBe(body?.blocks.length);
  });

  it("returns undefined for a missing story", () => {
    const doc = createDocx();
    expect(storyBody(doc, { kind: "footnote", id: 42 })).toBeUndefined();
    expect(storyBody(doc, { kind: "header", partName: "/word/header9.xml" })).toBeUndefined();
  });

  it("is discarded by a raw edit of its part, which wins", () => {
    const doc = createDocx();
    addHeader(doc, "A");
    const ref = { kind: "header", partName: "/word/header1.xml" } as const;
    const first = storyBody(doc, ref);
    const root = getRawPartRoot(doc, ref.partName);
    if (!root) throw new Error("no root");
    (root.children as unknown[]).length = 0;
    markRawPartDirty(doc, ref.partName);
    const second = storyBody(doc, ref);
    expect(second).not.toBe(first);
    expect(second?.blocks).toEqual([]);
    expect(partXml(toUint8Array(doc), ref.partName)).not.toContain("<w:t>");
  });
});

describe("storyView", () => {
  it("lets the body functions edit a story, and saves the real document", () => {
    const doc = createDocx({ paragraphs: ["Body text"] });
    addHeader(doc, "HeaderText");
    const ref = { kind: "header", partName: "/word/header1.xml" } as const;
    const view = storyView(doc, ref);
    if (!view) throw new Error("no view");
    expect(splitParagraphAt(view, 0, 0, 6)).toBe(1);
    expect(storyBody(doc, ref)?.blocks.length).toBe(2);
    // The view's body is the story, its section properties the document's.
    expect(view.document.body.sectPr).toBe(doc.document.body.sectPr);
    setPageSize(view, { widthTwips: 11906, heightTwips: 16838 });
    expect(storyBody(doc, ref)?.blocks.some((b) => b.kind === "raw")).toBe(false);

    const bytes = toUint8Array(view);
    const reopened = openDocx(bytes);
    expect(reopened.document.body.blocks.length).toBe(1);
    expect(partXml(bytes, "/word/document.xml")).toContain('w:w="11906"');
    expect(partXml(bytes, ref.partName)).toContain("<w:t>Header</w:t>");
    expect(partXml(bytes, ref.partName)).toContain("<w:t>Text</w:t>");
    expect(validatePackage(openDocx(bytes).opc)).toEqual([]);
  });
});

describe("story keys", () => {
  it("round-trip", () => {
    for (const ref of [
      { kind: "header", partName: "/word/header1.xml" },
      { kind: "footnote", id: 3 },
      { kind: "comment", id: 0 },
    ] as const) {
      expect(parseStoryKey(storyKey(ref))).toEqual(ref);
    }
    expect(parseStoryKey("nonsense")).toBeUndefined();
    expect(parseStoryKey("footnote:x")).toBeUndefined();
  });
});

describe("section headers and footers", () => {
  function twoSections() {
    const doc = createDocx({ paragraphs: ["Section 1"] });
    appendSectionBreak(doc, "nextPage");
    appendParagraph(doc, "Section 2");
    return doc;
  }

  it("lists one sectPr per section", () => {
    const doc = twoSections();
    expect(sectionProperties(doc)).toHaveLength(2);
  });

  it("inherits from the previous section, creates on demand, and links / unlinks", () => {
    const doc = twoSections();
    expect(resolveHeaderFooter(doc, 1, "header", "default")).toBeUndefined();
    const part = ensureHeaderFooter(doc, 0, "header", "default");
    expect(resolveHeaderFooter(doc, 1, "header", "default")).toEqual({
      partName: part,
      section: 0,
    });
    expect(isHeaderFooterLinked(doc, 1, "header", "default")).toBe(true);
    // Editing a linked header edits the one it comes from.
    expect(ensureHeaderFooter(doc, 1, "header", "default")).toBe(part);

    const body = storyBody(doc, { kind: "header", partName: part });
    const p = body?.blocks[0];
    if (p?.kind !== "paragraph") throw new Error("expected a paragraph");
    p.children = [
      {
        kind: "run",
        pieces: [{ kind: "text", value: "Shared", preserveSpace: false }],
        extras: [],
      },
    ];

    setHeaderFooterLinked(doc, 1, "header", "default", false);
    const own = resolveHeaderFooter(doc, 1, "header", "default");
    expect(own?.section).toBe(1);
    expect(own?.partName).not.toBe(part);
    if (!own) return;
    // The unlinked copy starts with what the section showed.
    expect(firstParagraphText(doc, { kind: "header", partName: own.partName })).toBe("Shared");

    let bytes = toUint8Array(doc);
    expect(validatePackage(openDocx(bytes).opc)).toEqual([]);

    setHeaderFooterLinked(doc, 1, "header", "default", true);
    expect(resolveHeaderFooter(doc, 1, "header", "default")).toEqual({
      partName: part,
      section: 0,
    });
    bytes = toUint8Array(doc);
    expect(validatePackage(openDocx(bytes).opc)).toEqual([]);
    expect(getPart(openDocx(bytes).opc, own.partName)).toBeUndefined();
  });

  it("refuses to link the first section", () => {
    expect(() => setHeaderFooterLinked(twoSections(), 0, "footer", "default", true)).toThrow();
  });
});
