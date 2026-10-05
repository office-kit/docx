/**
 * A minimal SpreadsheetML (ECMA-376 Part 1, §18) workbook: one sheet of
 * strings and numbers. Embedded next to a chart part so Word's Edit Data opens
 * the chart's numbers in a spreadsheet; the chart itself renders from the
 * caches in the chart part, not from this workbook.
 */

import {
  addPart,
  addRelationship,
  emptyOpcPackage,
  packageRelationships,
  partRelationships,
  setContentTypeDefault,
  writeOpcPackage,
} from "../opc/index.js";
import { escapeXml } from "./xml.js";

const SML_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const OFFICE_DOCUMENT_REL =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const WORKSHEET_REL =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
const WORKBOOK_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
const WORKSHEET_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
const RELS_TYPE = "application/vnd.openxmlformats-package.relationships+xml";
const XML_DECL = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n`;

/** The sheet name chart formulas (`Sheet1!$B$2:$B$5`) refer to. */
export const SHEET_NAME = "Sheet1";

export type CellValue = string | number | null;

/** Column letters for a zero-based column index (0 → A, 26 → AA). */
export function columnName(index: number): string {
  let n = index + 1;
  let out = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

function cellMarkup(ref: string, value: CellValue): string {
  if (value === null) return "";
  if (typeof value === "number") return `<c r="${ref}"><v>${value}</v></c>`;
  return `<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${escapeXml(value)}</t></is></c>`;
}

/** Write rows (row-major, first row first) as a one-sheet `.xlsx`. */
export function writeWorkbook(rows: ReadonlyArray<ReadonlyArray<CellValue>>): Uint8Array {
  const sheetRows = rows
    .map((row, r) => {
      const cells = row.map((v, c) => cellMarkup(`${columnName(c)}${r + 1}`, v)).join("");
      return `<row r="${r + 1}">${cells}</row>`;
    })
    .join("");
  const enc = new TextEncoder();
  const pkg = emptyOpcPackage();
  setContentTypeDefault(pkg.contentTypes, "rels", RELS_TYPE);
  setContentTypeDefault(pkg.contentTypes, "xml", "application/xml");
  addPart(pkg, {
    name: "/xl/workbook.xml",
    contentType: WORKBOOK_TYPE,
    data: enc.encode(
      `${XML_DECL}<workbook xmlns="${SML_NS}" xmlns:r="${R_NS}"><sheets><sheet name="${SHEET_NAME}" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    ),
  });
  addPart(pkg, {
    name: "/xl/worksheets/sheet1.xml",
    contentType: WORKSHEET_TYPE,
    data: enc.encode(
      `${XML_DECL}<worksheet xmlns="${SML_NS}"><sheetData>${sheetRows}</sheetData></worksheet>`,
    ),
  });
  addRelationship(packageRelationships(pkg), {
    type: OFFICE_DOCUMENT_REL,
    target: "xl/workbook.xml",
  });
  addRelationship(partRelationships(pkg, "/xl/workbook.xml"), {
    id: "rId1",
    type: WORKSHEET_REL,
    target: "worksheets/sheet1.xml",
  });
  return writeOpcPackage(pkg);
}
