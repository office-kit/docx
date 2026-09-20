// The starter code of the /repl page. Everything from the "Every
// @office-kit/docx function" line down is what the visitor sees in the
// editor; the imports and the `doc` declaration above it exist so
// svelte-check type-checks the starter against the live library, exactly
// like the other example files.
//
// The REPL runs the code as plain JavaScript, so nothing below the marker
// may use TypeScript-only syntax (annotations, `!`, `as`).

import {
  addBulletList,
  addFootnote,
  addHeader,
  addHyperlink,
  addNumberedList,
  addPageNumberFooter,
  addStyle,
  addTable,
  appendHeading,
  appendParagraph,
  appendTextRun,
  type Docx,
  ensureHeadingStyles,
  MARGINS_NORMAL,
  PAGE_SIZE_A4,
  setPageMargins,
  setPageSize,
  setParagraphAlignment,
  setParagraphStyle,
  setTableBorders,
  setTableCellShading,
  setTableRowAsHeader,
  setTitle,
} from "@office-kit/docx";

declare const doc: Docx;

// Every @office-kit/docx function is in scope. No imports needed.
// `doc` is a new, empty document from createDocx({ paragraphs: [] }).

const ACCENT = "2B63D9";
const MUTED = "5B616E";

setPageSize(doc, PAGE_SIZE_A4);
setPageMargins(doc, MARGINS_NORMAL);
setTitle(doc, "Q3 business review");
ensureHeadingStyles(doc);

// A custom paragraph style, applied by id.
addStyle(doc, {
  type: "paragraph",
  styleId: "ReportTitle",
  name: "Report Title",
  basedOn: "Normal",
  bold: true,
  fontSizeHalfPoints: 56,
  color: ACCENT,
});

addHeader(doc, "Acme Inc. Confidential");
addPageNumberFooter(doc, "Page ");

const cover = appendParagraph(doc, "Q3 business review");
setParagraphStyle(cover, "ReportTitle");
const byline = appendParagraph(doc, "");
appendTextRun(byline, "Finance team, October 2026", { italic: true, color: MUTED });

appendHeading(doc, "Summary", 1);
const summary = appendParagraph(doc, "Revenue reached ");
appendTextRun(summary, "$1.48M", { bold: true });
appendTextRun(summary, ", up 23% on last quarter, and churn fell to ");
appendTextRun(summary, "2.4%", { bold: true, color: ACCENT });
appendTextRun(summary, ", the lowest since launch.");
addFootnote(doc, summary, "Churn is measured on paying accounts, monthly.");

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
table.rows.slice(0, 1).forEach((header) => {
  setTableRowAsHeader(header);
  header.cells.forEach((cell) => setTableCellShading(cell, { fill: "E9EFFC" }));
});

appendHeading(doc, "Next steps", 2);
addNumberedList(doc, [
  "Renegotiate the hosting contract",
  "Ship annual billing",
  "Hire two support engineers",
]);

const source = addHyperlink(
  doc,
  "https://office-kit.github.io/docx/",
  "Built with @office-kit/docx",
);
setParagraphAlignment(source, "right");
