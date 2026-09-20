// The document on the landing page. `+page.server.ts` runs this at build
// time and draws what it reads back from the saved bytes; the Download
// button runs the same function in the visitor's browser. Keep it
// self-contained so the snippet shown on the page is the whole program.

import {
  addBulletList,
  addTable,
  appendHeading,
  appendParagraph,
  appendTextRun,
  createDocx,
  type Docx,
  ensureHeadingStyles,
  MARGINS_NORMAL,
  PAGE_SIZE_A4,
  setPageMargins,
  setPageSize,
  setTableBorders,
  setTableCellShading,
  setTableRowAsHeader,
} from "@office-kit/docx";

export function buildHeroDocument(): Docx {
  const doc = createDocx({ paragraphs: [] });
  setPageSize(doc, PAGE_SIZE_A4);
  setPageMargins(doc, MARGINS_NORMAL);
  ensureHeadingStyles(doc);

  appendHeading(doc, "Quarterly review", 1);
  const intro = appendParagraph(doc, "Revenue grew 23% on last quarter. ");
  appendTextRun(intro, "Churn hit a record low.", { bold: true });

  appendHeading(doc, "Highlights", 2);
  addBulletList(doc, [
    "Shipped the self-serve onboarding flow",
    "Cut median support response time to 2 hours",
    "Opened the Osaka office",
  ]);

  appendHeading(doc, "Numbers", 2);
  const table = addTable(doc, [
    ["Metric", "Q2", "Q3"],
    ["Revenue", "$1.20M", "$1.48M"],
    ["Customers", "412", "497"],
    ["Churn", "3.1%", "2.4%"],
  ]);
  setTableBorders(table, { color: "BFC5D2" });
  const header = table.rows[0]!;
  setTableRowAsHeader(header);
  for (const cell of header.cells) setTableCellShading(cell, { fill: "E9EFFC" });

  return doc;
}
