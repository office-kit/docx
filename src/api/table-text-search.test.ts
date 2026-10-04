import { describe, expect, it } from "vitest";
import {
  addTable,
  appendParagraph,
  createDocx,
  findText,
  findTextEverywhere,
  getTableCellText,
  openDocx,
  paragraphText,
  paragraphs,
  replaceText,
  replaceTextEverywhere,
  tables,
  toUint8Array,
  validate,
} from "./index.js";

// findText / replaceText promise "every paragraph" / "every occurrence"; a
// paragraph inside a table cell is a paragraph of the document too.
function docWithTable(): ReturnType<typeof createDocx> {
  const doc = createDocx({ paragraphs: ["DRAFT before"] });
  addTable(doc, [
    ["DRAFT a", "b"],
    ["c", "x DRAFT DRAFT"],
  ]);
  appendParagraph(doc, "after DRAFT");
  return doc;
}

describe("find / replace reach table cells", () => {
  it("findText returns cell matches in document order, pointing at the cell paragraph", () => {
    const doc = docWithTable();
    const matches = findText(doc, "DRAFT");
    expect(matches).toHaveLength(5);
    const cellPara = tables(doc)[0]!.rows[1]!.cells[1]!.paragraphs[0]!;
    expect(matches.map((m) => paragraphText(m.paragraph))).toEqual([
      "DRAFT before",
      "DRAFT a",
      "x DRAFT DRAFT",
      "x DRAFT DRAFT",
      "after DRAFT",
    ]);
    expect(matches[2]!.paragraph).toBe(cellPara);
  });

  it("replaceText counts and replaces cell text, and it survives save → reopen", () => {
    const doc = docWithTable();
    expect(replaceText(doc, "DRAFT", "FINAL")).toBe(5);
    const reopened = openDocx(toUint8Array(doc));
    const table = tables(reopened)[0]!;
    expect(getTableCellText(table, 0, 0)).toBe("FINAL a");
    expect(getTableCellText(table, 1, 1)).toBe("x FINAL FINAL");
    expect(paragraphs(reopened).map(paragraphText)).toEqual(["FINAL before", "after FINAL"]);
    expect(validate(reopened)).toHaveLength(0);
    expect(findText(reopened, "DRAFT")).toHaveLength(0);
  });

  it("findTextEverywhere / replaceTextEverywhere include body table cells", () => {
    const doc = docWithTable();
    const body = findTextEverywhere(doc, "DRAFT").find((e) => e.partName === doc.partName);
    expect(body?.matches).toHaveLength(5);
    expect(replaceTextEverywhere(doc, "DRAFT", "FINAL")).toBe(5);
    expect(getTableCellText(tables(doc)[0]!, 1, 1)).toBe("x FINAL FINAL");
  });
});
