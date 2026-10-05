/**
 * Word's Mailings tab: envelopes and labels, the mail merge settings
 * (`w:mailMerge`, §17.14), merge fields and rules, previewing a recipient's
 * values in the merge fields, and merging to a new document.
 *
 * The recipient list itself is external to the document, as in Word: the
 * document records where it came from (an external `mailMergeSource`
 * relationship) and which recipients are excluded (`recipientData.xml`), and
 * callers pass the parsed records in when previewing or merging.
 */

import type { XmlElement } from "../internal/xml/index.js";
import {
  addPart,
  addRelationship,
  allRelationships,
  getPart,
  hasPart,
  partRelationships,
  relationshipsByType,
  removePart,
  removeRelationship,
} from "../internal/opc/index.js";
import type { WmlBlock, WmlInline, WmlParagraph } from "../internal/wordprocessingml/index.js";
import {
  type FieldInstruction,
  fieldSwitch,
  formatFieldNumber,
  quotedFieldArg,
  quoteFieldArg,
} from "../internal/wordprocessingml/field-code.js";
import {
  blocksFromXml,
  childElement,
  elementFromXml,
  elementXml,
  escapeXml,
  fieldRunsXml,
  type FlowField,
  inlinesFromXml,
  paragraphFlow,
  paragraphsFromXml,
  resultlessFieldRunsXml,
  scanFields,
  setInlineResult,
  textRunXml,
  wAttr,
  withOpenContentControls,
} from "../internal/wordprocessingml/field-flow.js";
import {
  findChild,
  MAIL_MERGE_ORDER,
  ODSO_ORDER,
  SECT_PR_ORDER,
  SETTINGS_ORDER,
  setOrderedChild,
} from "../internal/wordprocessingml/ordered.js";
import { createDocx, type Docx, openDocx, toUint8Array } from "./docx.js";
import {
  editSettings,
  ensureBodySectPr,
  ensureBuiltInStyles,
  insertInlinesAt,
  readSettings,
} from "./part-helpers.js";

const SETTINGS_PART_NAME = "/word/settings.xml";
const MAIL_MERGE_SOURCE_REL =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/mailMergeSource";
const RECIPIENT_DATA_REL =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/recipientData";
const RECIPIENT_DATA_CT =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.mailMergeRecipientData+xml";
const RECIPIENT_DATA_PART = "/word/recipientData.xml";

// --- Main document type --------------------------------------------------------

/** ST_MailMergeDocType (§17.18.54): Start Mail Merge's document kinds. */
export type MailMergeDocumentType =
  | "formLetters"
  | "email"
  | "envelopes"
  | "mailingLabels"
  | "catalog"
  | "fax";

const DOCUMENT_TYPES: ReadonlySet<string> = new Set([
  "formLetters",
  "email",
  "envelopes",
  "mailingLabels",
  "catalog",
  "fax",
]);

/** The mail merge state the document's settings record. */
export interface MailMergeSettings {
  readonly type: MailMergeDocumentType;
  /** The recipient list's file path, as recorded in the data source relationship. */
  readonly dataSource?: string;
  /** Preview Results is on (`w:viewMergedData`). */
  readonly preview: boolean;
  /** The previewed record, 1-based (`w:activeRecord`). */
  readonly activeRecord?: number;
}

/** The document's mail merge settings, or `undefined` for a normal Word document. */
export function mailMergeSettings(doc: Docx): MailMergeSettings | undefined {
  const mm = findChild(readSettings(doc), "mailMerge");
  if (!mm) return undefined;
  const type = wAttr(findChild(mm, "mainDocumentType"), "val") ?? "formLetters";
  const relId = wAttr(findChild(mm, "dataSource"), "id");
  const rel = relId
    ? allRelationships(partRelationships(doc.opc, SETTINGS_PART_NAME)).find((r) => r.id === relId)
    : undefined;
  const view = findChild(mm, "viewMergedData");
  const active = wAttr(findChild(mm, "activeRecord"), "val");
  return {
    type: DOCUMENT_TYPES.has(type) ? (type as MailMergeDocumentType) : "formLetters",
    ...(rel ? { dataSource: decodeFileUrl(rel.target) } : {}),
    preview: !!view && wAttr(view, "val") !== "0" && wAttr(view, "val") !== "false",
    ...(active ? { activeRecord: Number(active) } : {}),
  };
}

/**
 * Start Mail Merge: make the document a merge main document of `type`, or a
 * normal Word document again with `"normal"` (which drops the merge settings
 * and the recipient list link).
 */
export function setMailMergeDocumentType(
  doc: Docx,
  type: MailMergeDocumentType | "normal",
): void {
  if (type === "normal") {
    editSettings(doc, (root) => {
      const children = root.children as XmlElement[];
      const i = children.findIndex((c) => c.kind === "element" && c.name.local === "mailMerge");
      if (i >= 0) children.splice(i, 1);
    });
    dropMergeRelationships(doc);
    return;
  }
  if (!DOCUMENT_TYPES.has(type)) throw new Error(`Unknown mail merge document type ${type}.`);
  editMailMerge(doc, (mm) => {
    setOrderedChild(mm, elementFromXml(`<w:mainDocumentType w:val="${type}"/>`), MAIL_MERGE_ORDER);
    // A merge main document without a data source still needs w:dataType.
    if (!findChild(mm, "dataType")) {
      setOrderedChild(mm, elementFromXml('<w:dataType w:val="textFile"/>'), MAIL_MERGE_ORDER);
    }
  });
}

function editMailMerge(doc: Docx, edit: (mm: XmlElement) => void): void {
  editSettings(doc, (root) => {
    let mm = findChild(root, "mailMerge");
    if (!mm) {
      mm = elementFromXml('<w:mailMerge><w:mainDocumentType w:val="formLetters"/></w:mailMerge>');
      setOrderedChild(root, mm, SETTINGS_ORDER);
    }
    edit(mm);
  });
}

function dropMergeRelationships(doc: Docx): void {
  const rels = partRelationships(doc.opc, SETTINGS_PART_NAME);
  const doomed = allRelationships(rels)
    .filter((rel) => rel.type === MAIL_MERGE_SOURCE_REL || rel.type === RECIPIENT_DATA_REL)
    .map((rel) => rel.id);
  for (const id of doomed) removeRelationship(rels, id);
  if (hasPart(doc.opc, RECIPIENT_DATA_PART)) removePart(doc.opc, RECIPIENT_DATA_PART);
}

// --- Recipient list --------------------------------------------------------------

/** A recipient list: column names and one value per column for each record. */
export interface RecipientList {
  readonly columns: readonly string[];
  readonly records: ReadonlyArray<Readonly<Record<string, string>>>;
}

const MAX_CSV_BYTES = 10 * 1024 * 1024;

/**
 * Parse a CSV recipient list (RFC 4180: quoted fields, doubled quotes, CRLF
 * or LF). The first row holds the column names; blank lines are skipped.
 * Rejects input over 10 MB, a header-less file, and duplicate column names.
 */
export function parseRecipientCsv(text: string, delimiter = ","): RecipientList {
  if (text.length > MAX_CSV_BYTES) throw new Error("The recipient list is too large (over 10 MB).");
  const rows: string[][] = [];
  let row: string[] = [];
  let cell = "";
  let quoted = false;
  const input = text.replace(/^﻿/, "");
  for (let i = 0; i < input.length; i++) {
    const ch = input[i] as string;
    if (quoted) {
      if (ch === '"' && input[i + 1] === '"') {
        cell += '"';
        i++;
      } else if (ch === '"') {
        quoted = false;
      } else {
        cell += ch;
      }
      continue;
    }
    if (ch === '"' && cell === "") quoted = true;
    else if (ch === delimiter) {
      row.push(cell);
      cell = "";
    } else if (ch === "\n" || ch === "\r") {
      if (ch === "\r" && input[i + 1] === "\n") i++;
      row.push(cell);
      rows.push(row);
      row = [];
      cell = "";
    } else {
      cell += ch;
    }
  }
  if (cell !== "" || row.length > 0) {
    row.push(cell);
    rows.push(row);
  }
  const nonEmpty = rows.filter((r) => r.some((c) => c.trim() !== ""));
  const header = nonEmpty[0]?.map((c) => c.trim());
  if (!header || header.every((c) => c === "")) throw new Error("The recipient list has no header row.");
  const seen = new Set<string>();
  for (const name of header) {
    if (!name) throw new Error("Every column of the recipient list needs a name.");
    if (seen.has(name.toLowerCase())) throw new Error(`Duplicate column ${JSON.stringify(name)}.`);
    seen.add(name.toLowerCase());
  }
  const records = nonEmpty.slice(1).map((r) =>
    Object.fromEntries(header.map((name, i) => [name, r[i] ?? ""])),
  );
  return { columns: header, records };
}

function csvQuote(value: string): string {
  return /[",\r\n]/.test(value) ? `"${value.replace(/"/g, '""')}"` : value;
}

/** Serialize a recipient list as CSV (Type a New List ▸ Save). */
export function recipientListToCsv(list: RecipientList): string {
  const lines = [list.columns.map(csvQuote).join(",")];
  for (const r of list.records) lines.push(list.columns.map((c) => csvQuote(r[c] ?? "")).join(","));
  return `${lines.join("\r\n")}\r\n`;
}

function fileUrl(path: string): string {
  if (/^[a-z][a-z0-9+.-]*:/i.test(path) && !/^[a-z]:\\/i.test(path)) return path;
  const slashed = path.replace(/\\/g, "/");
  return `file:///${encodeURI(slashed.replace(/^\/+/, ""))}`;
}

function decodeFileUrl(target: string): string {
  if (!target.startsWith("file:///")) return target;
  const path = decodeURI(target.slice("file:///".length));
  return /^[a-z]:\//i.test(path) ? path : `/${path}`;
}

/** ASCII code of the column delimiter Word records in `w:colDelim`. */
const COMMA_CODE = 44;

/**
 * Select Recipients ▸ Use an Existing List: link a delimited text recipient
 * list (CSV) the way Word does — `w:dataType textFile`, a `SELECT * FROM`
 * query, an external `mailMergeSource` relationship to the file, and the
 * ODSO description with the header-row flag and the field mapping.
 */
export function attachRecipientList(
  doc: Docx,
  path: string,
  list: RecipientList,
  fieldMap: Readonly<Record<string, string>> = autoFieldMap(list.columns),
): void {
  const name = path.trim();
  if (!name) throw new Error("A recipient list needs a file name.");
  dropMergeRelationships(doc);
  if (!mailMergeSettings(doc)) setMailMergeDocumentType(doc, "formLetters");
  const rels = partRelationships(doc.opc, SETTINGS_PART_NAME);
  const source = addRelationship(rels, {
    type: MAIL_MERGE_SOURCE_REL,
    target: fileUrl(name),
    targetMode: "External",
  });
  const odsoSource = addRelationship(rels, {
    type: MAIL_MERGE_SOURCE_REL,
    target: fileUrl(name),
    targetMode: "External",
  });
  const table = name.split(/[\\/]/).pop() ?? name;
  editMailMerge(doc, (mm) => {
    setOrderedChild(mm, elementFromXml("<w:linkToQuery/>"), MAIL_MERGE_ORDER);
    setOrderedChild(mm, elementFromXml('<w:dataType w:val="textFile"/>'), MAIL_MERGE_ORDER);
    setOrderedChild(mm, elementFromXml('<w:connectString w:val=""/>'), MAIL_MERGE_ORDER);
    setOrderedChild(
      mm,
      elementFromXml(`<w:query w:val="SELECT * FROM ${escapeXml(name)}"/>`),
      MAIL_MERGE_ORDER,
    );
    setOrderedChild(mm, elementFromXml(`<w:dataSource r:id="${source.id}"/>`), MAIL_MERGE_ORDER);
    const odso = elementFromXml(
      `<w:odso><w:udl w:val=""/><w:table w:val="${escapeXml(table)}"/><w:src r:id="${odsoSource.id}"/><w:colDelim w:val="${COMMA_CODE}"/><w:fHdr/></w:odso>`,
    );
    setOrderedChild(mm, odso, MAIL_MERGE_ORDER);
  });
  setMailMergeFieldMap(doc, list.columns, fieldMap);
}

// --- Field mapping (Match Fields) ----------------------------------------------------

/** Word's address fields (Match Fields), in the order ODSO field map entries take. */
export const ADDRESS_FIELDS = [
  "Unique Identifier",
  "Courtesy Title",
  "First Name",
  "Middle Name",
  "Last Name",
  "Suffix",
  "Nickname",
  "Job Title",
  "Company",
  "Address 1",
  "Address 2",
  "City",
  "State",
  "Postal Code",
  "Country or Region",
  "Business Phone",
  "Business Fax",
  "Home Phone",
  "Home Fax",
  "E-mail Address",
  "Web Page",
  "Spouse/Partner Courtesy Title",
  "Spouse/Partner First Name",
  "Spouse/Partner Middle Name",
  "Spouse/Partner Last Name",
  "Spouse/Partner Nickname",
  "Rural Route",
  "Address 3",
  "Department",
] as const;
export type AddressField = (typeof ADDRESS_FIELDS)[number];

// Column names commonly used for each address field, matched loosely
// (case, spaces, dashes and underscores ignored) as Word's auto-match does.
const SYNONYMS: Readonly<Partial<Record<AddressField, readonly string[]>>> = {
  "Unique Identifier": ["id", "uniqueid"],
  "Courtesy Title": ["title", "courtesytitle", "salutation", "prefix"],
  "First Name": ["firstname", "first", "givenname", "forename"],
  "Middle Name": ["middlename", "middle", "middleinitial"],
  "Last Name": ["lastname", "last", "surname", "familyname"],
  Suffix: ["suffix"],
  Nickname: ["nickname"],
  "Job Title": ["jobtitle", "position"],
  Company: ["company", "companyname", "organization", "organisation"],
  "Address 1": ["address1", "address", "addressline1", "street", "streetaddress"],
  "Address 2": ["address2", "addressline2"],
  City: ["city", "town"],
  State: ["state", "province", "region", "stateprovince", "county"],
  "Postal Code": ["postalcode", "zip", "zipcode", "postcode"],
  "Country or Region": ["countryorregion", "country", "countryregion"],
  "Business Phone": ["businessphone", "workphone", "phone"],
  "Home Phone": ["homephone"],
  "E-mail Address": ["emailaddress", "email", "mail"],
  "Web Page": ["webpage", "website", "url"],
  Department: ["department"],
};

const normalizeName = (s: string): string => s.toLowerCase().replace(/[\s_\-./]/g, "");

/** Match Fields' automatic mapping: address field → recipient-list column. */
export function autoFieldMap(columns: readonly string[]): Record<string, string> {
  const byNorm = new Map(columns.map((c) => [normalizeName(c), c]));
  const out: Record<string, string> = {};
  for (const field of ADDRESS_FIELDS) {
    const candidates = [normalizeName(field), ...(SYNONYMS[field] ?? [])];
    const hit = candidates.map((c) => byNorm.get(c)).find((c) => c !== undefined);
    if (hit) out[field] = hit;
  }
  return out;
}

/**
 * Store the Match Fields mapping in `w:odso` (one `w:fieldMapData` per
 * address field, in Word's order).
 */
export function setMailMergeFieldMap(
  doc: Docx,
  columns: readonly string[],
  map: Readonly<Record<string, string>>,
): void {
  const entries = ADDRESS_FIELDS.map((field) => {
    const column = map[field];
    const index = column === undefined ? -1 : columns.indexOf(column);
    if (index < 0) return '<w:fieldMapData><w:column w:val="0"/><w:lid w:val="en-US"/></w:fieldMapData>';
    return `<w:fieldMapData><w:type w:val="dbColumn"/><w:name w:val="${escapeXml(column as string)}"/><w:mappedName w:val="${escapeXml(field)}"/><w:column w:val="${index}"/><w:lid w:val="en-US"/></w:fieldMapData>`;
  });
  editMailMerge(doc, (mm) => {
    let odso = findChild(mm, "odso");
    if (!odso) {
      odso = elementFromXml("<w:odso/>");
      setOrderedChild(mm, odso, MAIL_MERGE_ORDER);
    }
    const kept = odso.children.filter(
      (c): c is XmlElement => c.kind === "element" && c.name.local !== "fieldMapData",
    );
    const rebuilt = elementFromXml(`<w:odso>${entries.join("")}</w:odso>`);
    for (const el of kept) setOrderedChild(rebuilt, el, ODSO_ORDER);
    setOrderedChild(mm, rebuilt, MAIL_MERGE_ORDER);
  });
}

/** The stored Match Fields mapping (address field → column). */
export function mailMergeFieldMap(doc: Docx): Record<string, string> {
  const odso = findChild(findChild(readSettings(doc), "mailMerge"), "odso");
  const out: Record<string, string> = {};
  for (const el of odso?.children ?? []) {
    if (el.kind !== "element" || el.name.local !== "fieldMapData") continue;
    const mapped = wAttr(findChild(el, "mappedName"), "val");
    const name = wAttr(findChild(el, "name"), "val");
    if (mapped && name) out[mapped] = name;
  }
  return out;
}

/**
 * Edit Recipient List: record which recipients are excluded (unchecked) in
 * `recipientData.xml` (§17.14.27). Each record is identified by the value of
 * its first column, as a base64 `w:uniqueTag`. Pass every record's inclusion
 * flag in list order.
 */
export function setRecipientInclusion(doc: Docx, list: RecipientList, included: readonly boolean[]): void {
  const keyColumn = list.columns[0];
  if (keyColumn === undefined) throw new Error("The recipient list has no columns.");
  const entries = list.records.map((r, i) => {
    const active = included[i] ?? true;
    const tag = base64(r[keyColumn] ?? "");
    return `<w:recipientData><w:active w:val="${active ? "1" : "0"}"/><w:column w:val="0"/><w:uniqueTag>${tag}</w:uniqueTag></w:recipientData>`;
  });
  const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:recipients xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${entries.join("")}</w:recipients>`;
  const data = new TextEncoder().encode(xml);
  const existing = getPart(doc.opc, RECIPIENT_DATA_PART);
  let relId: string;
  if (existing) {
    existing.data = data;
    const rels = partRelationships(doc.opc, SETTINGS_PART_NAME);
    relId = relationshipsByType(rels, RECIPIENT_DATA_REL)[0]?.id ?? "";
  } else {
    addPart(doc.opc, { name: RECIPIENT_DATA_PART, contentType: RECIPIENT_DATA_CT, data });
    relId = addRelationship(partRelationships(doc.opc, SETTINGS_PART_NAME), {
      type: RECIPIENT_DATA_REL,
      target: "recipientData.xml",
    }).id;
  }
  editMailMerge(doc, (mm) => {
    let odso = findChild(mm, "odso");
    if (!odso) {
      odso = elementFromXml("<w:odso/>");
      setOrderedChild(mm, odso, MAIL_MERGE_ORDER);
    }
    setOrderedChild(odso, elementFromXml(`<w:recipientData r:id="${relId}"/>`), ODSO_ORDER);
  });
}

/** Which records the document's recipientData marks included (all, when there is none). */
export function recipientInclusion(doc: Docx, list: RecipientList): boolean[] {
  const part = getPart(doc.opc, RECIPIENT_DATA_PART);
  const keyColumn = list.columns[0];
  if (!part || keyColumn === undefined) return list.records.map(() => true);
  const xml = new TextDecoder("utf-8").decode(part.data);
  const excluded = new Set<string>();
  for (const m of xml.matchAll(/<w:recipientData>(.*?)<\/w:recipientData>/gs)) {
    const body = m[1] ?? "";
    const active = /<w:active w:val="(0|false|off)"\/>/.test(body);
    const tag = /<w:uniqueTag>(.*?)<\/w:uniqueTag>/.exec(body)?.[1];
    if (active && tag !== undefined) excluded.add(tag);
  }
  return list.records.map((r) => !excluded.has(base64(r[keyColumn] ?? "")));
}

function base64(text: string): string {
  const bytes = new TextEncoder().encode(text);
  let binary = "";
  for (const b of bytes) binary += String.fromCharCode(b);
  return btoa(binary);
}

// --- Merge fields and rules ---------------------------------------------------------------

/** The MERGEFIELD instruction for a column name. */
export function mergeFieldInstruction(name: string): string {
  if (!name.trim()) throw new Error("A merge field needs a column name.");
  return `MERGEFIELD ${quoteFieldArg(name.trim())}`;
}

/**
 * Insert Merge Field: a MERGEFIELD at a character offset, showing «Name»
 * until results are previewed.
 */
export function insertMergeField(doc: Docx, paragraph: WmlParagraph, offset: number, name: string): void {
  insertInlinesAt(
    paragraph,
    offset,
    inlinesFromXml(fieldRunsXml(mergeFieldInstruction(name), textRunXml(`«${name.trim()}»`))),
  );
  doc.dirty = true;
}

/** Address Block options (Insert Address Block dialog). */
export interface AddressBlockOptions {
  /** The recipient-name format, as tokens: e.g. `<<_TITLE0_ >><<_FIRST0_>><< _LAST0_>>`. */
  readonly nameFormat?: string;
  /** Insert company name. Default true. */
  readonly company?: boolean;
  /** Insert postal address. Default true. */
  readonly postalAddress?: boolean;
  /** Country/region: never, always, or only when different from `excludedCountry`. Default "ifDifferent". */
  readonly country?: "never" | "always" | "ifDifferent";
  readonly excludedCountry?: string;
  /** Format the address according to the destination country/region. Default true. */
  readonly formatByCountry?: boolean;
}

/** Word's recipient-name formats for the Address Block, keyed by Word's example. */
export const ADDRESS_NAME_FORMATS = {
  Joshua: "<<_FIRST0_>>",
  "Joshua Randall Jr.": "<<_FIRST0_>><< _LAST0_>><< _SUFFIX0_>>",
  "Joshua Q. Randall Jr.": "<<_FIRST0_>><< _MIDDLE0_>><< _LAST0_>><< _SUFFIX0_>>",
  "Mr. Randall": "<<_TITLE0_>><< _LAST0_>>",
  "Mr. Josh Randall Jr.": "<<_TITLE0_ >><<_NICK0_>><< _LAST0_>><< _SUFFIX0_>>",
  "Mr. Joshua Randall Jr.": "<<_TITLE0_ >><<_FIRST0_>><< _LAST0_>><< _SUFFIX0_>>",
  "Mr. Joshua Q. Randall Jr.": "<<_TITLE0_ >><<_FIRST0_>><< _MIDDLE0_>><< _LAST0_>><< _SUFFIX0_>>",
} as const;

const DEFAULT_NAME_FORMAT = ADDRESS_NAME_FORMATS["Mr. Joshua Randall Jr."];
const DEFAULT_COUNTRY = "United States";
const LCID_EN_US = 1033;

/** The ADDRESSBLOCK instruction for these options. */
export function addressBlockInstruction(options: AddressBlockOptions = {}): string {
  let format = options.nameFormat ?? DEFAULT_NAME_FORMAT;
  if (options.company ?? true) format += "\n<<_COMPANY_\n>>";
  else format += "\n";
  if (options.postalAddress ?? true) {
    format += "<<_STREET1_\n>><<_STREET2_\n>><<_CITY_>><<, _STATE_>><< _POSTAL_>><<\n_COUNTRY_>>";
  }
  const country = options.country ?? "ifDifferent";
  const code = country === "never" ? 0 : country === "always" ? 1 : 2;
  const parts = ["ADDRESSBLOCK", "\\f", quotedFieldArg(format), "\\l", String(LCID_EN_US), "\\c", String(code)];
  if (country === "ifDifferent") parts.push("\\e", quotedFieldArg(options.excludedCountry ?? DEFAULT_COUNTRY));
  if (options.formatByCountry ?? true) parts.push("\\d");
  return parts.join(" ");
}

/** Insert an Address Block (ADDRESSBLOCK field) at a character offset. */
export function insertAddressBlock(
  doc: Docx,
  paragraph: WmlParagraph,
  offset: number,
  options: AddressBlockOptions = {},
): void {
  insertInlinesAt(
    paragraph,
    offset,
    inlinesFromXml(fieldRunsXml(addressBlockInstruction(options), textRunXml("«AddressBlock»"))),
  );
  doc.dirty = true;
}

/** Greeting Line options (Insert Greeting Line dialog). */
export interface GreetingLineOptions {
  /** "Dear", "To", or "" for none. Default "Dear". */
  readonly salutation?: string;
  /** Name tokens, e.g. `<<_TITLE0_ >><<_LAST0_>>` (Mr. Randall). */
  readonly nameFormat?: string;
  /** ",", ":", or "" for none. Default ",". */
  readonly punctuation?: string;
  /** Greeting for recipients without a name. Default "Dear Sir or Madam,". */
  readonly invalidNameGreeting?: string;
}

/** Word's Greeting Line name formats, keyed by Word's example. */
export const GREETING_NAME_FORMATS = {
  "Mr. Randall": "<<_TITLE0_ >><<_LAST0_>>",
  "Mr. Joshua Randall": "<<_TITLE0_ >><<_FIRST0_ >><<_LAST0_>>",
  "Mr. Josh Randall": "<<_TITLE0_ >><<_NICK0_ >><<_LAST0_>>",
  Joshua: "<<_FIRST0_>>",
  Josh: "<<_NICK0_>>",
  "Joshua Randall": "<<_FIRST0_ >><<_LAST0_>>",
  "Josh Randall": "<<_NICK0_ >><<_LAST0_>>",
} as const;

/** The GREETINGLINE instruction for these options. */
export function greetingLineInstruction(options: GreetingLineOptions = {}): string {
  const salutation = options.salutation ?? "Dear";
  const name = options.nameFormat ?? GREETING_NAME_FORMATS["Mr. Randall"];
  const punctuation = options.punctuation ?? ",";
  const format = `${salutation ? `<<_BEFORE_ ${salutation} >>` : ""}${name}${punctuation ? `<<_AFTER_ ${punctuation}>>` : ""}`;
  return [
    "GREETINGLINE",
    "\\f",
    quotedFieldArg(format),
    "\\l",
    String(LCID_EN_US),
    "\\e",
    quotedFieldArg(options.invalidNameGreeting ?? "Dear Sir or Madam,"),
  ].join(" ");
}

/** Insert a Greeting Line (GREETINGLINE field) at a character offset. */
export function insertGreetingLine(
  doc: Docx,
  paragraph: WmlParagraph,
  offset: number,
  options: GreetingLineOptions = {},
): void {
  insertInlinesAt(
    paragraph,
    offset,
    inlinesFromXml(fieldRunsXml(greetingLineInstruction(options), textRunXml("«GreetingLine»"))),
  );
  doc.dirty = true;
}

/** The comparisons of Word's rule dialogs ("Is blank" is `= ""`). */
export type MergeComparison = "=" | "<>" | "<" | "<=" | ">" | ">=" | "isBlank" | "isNotBlank";

/** Rules ▸ … : one Word mail merge rule. */
export type MergeRule =
  | { readonly kind: "ask"; readonly bookmark: string; readonly prompt: string; readonly defaultText?: string; readonly askOnce?: boolean }
  | { readonly kind: "fillIn"; readonly prompt: string; readonly defaultText?: string; readonly askOnce?: boolean }
  | {
      readonly kind: "if";
      readonly field: string;
      readonly comparison: MergeComparison;
      readonly value?: string;
      readonly trueText: string;
      readonly falseText: string;
    }
  | { readonly kind: "mergeRecord" }
  | { readonly kind: "mergeSequence" }
  | { readonly kind: "nextRecord" }
  | { readonly kind: "nextRecordIf"; readonly field: string; readonly comparison: MergeComparison; readonly value?: string }
  | { readonly kind: "setBookmark"; readonly bookmark: string; readonly value: string }
  | { readonly kind: "skipRecordIf"; readonly field: string; readonly comparison: MergeComparison; readonly value?: string };

const BOOKMARK_NAME = /^[\p{L}_][\p{L}\p{N}_]{0,39}$/u;

function comparisonXml(field: string, comparison: MergeComparison, value: string | undefined): string {
  const op = comparison === "isBlank" ? "=" : comparison === "isNotBlank" ? "<>" : comparison;
  const right = comparison === "isBlank" || comparison === "isNotBlank" ? "" : (value ?? "");
  // The compared field is nested, so its current value takes part in the test.
  return `${fieldRunsXml(mergeFieldInstruction(field), textRunXml(`«${field}»`))}<w:r><w:instrText xml:space="preserve"> ${escapeXml(op)} ${escapeXml(quotedFieldArg(right))} </w:instrText></w:r>`;
}

function nestedFieldXml(head: string, middleXml: string, tail: string, resultXml: string): string {
  return [
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r>',
    `<w:r><w:instrText xml:space="preserve"> ${escapeXml(head)} </w:instrText></w:r>`,
    middleXml,
    tail ? `<w:r><w:instrText xml:space="preserve">${escapeXml(tail)} </w:instrText></w:r>` : "",
    '<w:r><w:fldChar w:fldCharType="separate"/></w:r>',
    resultXml,
    '<w:r><w:fldChar w:fldCharType="end"/></w:r>',
  ].join("");
}

function checkBookmark(name: string): void {
  if (!BOOKMARK_NAME.test(name)) {
    throw new Error(`${JSON.stringify(name)} is not a valid bookmark name (letters, digits, underscores; up to 40).`);
  }
}

/** The runs a rule inserts. */
function ruleXml(rule: MergeRule): string {
  switch (rule.kind) {
    case "ask": {
      checkBookmark(rule.bookmark);
      const parts = ["ASK", rule.bookmark, quotedFieldArg(rule.prompt)];
      if (rule.defaultText) parts.push("\\d", quotedFieldArg(rule.defaultText));
      if (rule.askOnce) parts.push("\\o");
      return fieldRunsXml(parts.join(" "), "");
    }
    case "fillIn": {
      const parts = ["FILLIN", quotedFieldArg(rule.prompt)];
      if (rule.defaultText) parts.push("\\d", quotedFieldArg(rule.defaultText));
      if (rule.askOnce) parts.push("\\o");
      return fieldRunsXml(parts.join(" "), textRunXml(rule.defaultText ?? ""));
    }
    case "if":
      return nestedFieldXml(
        "IF",
        comparisonXml(rule.field, rule.comparison, rule.value),
        `${quotedFieldArg(rule.trueText)} ${quotedFieldArg(rule.falseText)}`,
        textRunXml(rule.falseText),
      );
    case "mergeRecord":
      return fieldRunsXml("MERGEREC", textRunXml("«Merge Record #»"));
    case "mergeSequence":
      return fieldRunsXml("MERGESEQ", textRunXml("«Merge Sequence #»"));
    case "nextRecord":
      return fieldRunsXml("NEXT", "");
    case "nextRecordIf":
      return nestedFieldXml("NEXTIF", comparisonXml(rule.field, rule.comparison, rule.value), "", "");
    case "setBookmark":
      checkBookmark(rule.bookmark);
      return fieldRunsXml(`SET ${rule.bookmark} ${quotedFieldArg(rule.value)}`, "");
    case "skipRecordIf":
      return nestedFieldXml("SKIPIF", comparisonXml(rule.field, rule.comparison, rule.value), "", "");
  }
}

/** Insert a mail merge rule (Rules menu) at a character offset. */
export function insertMergeRule(doc: Docx, paragraph: WmlParagraph, offset: number, rule: MergeRule): void {
  insertInlinesAt(paragraph, offset, inlinesFromXml(ruleXml(rule)));
  doc.dirty = true;
}

// --- Evaluation ------------------------------------------------------------------------

/** Field types the merge evaluates (and Edit Individual Documents turns into text). */
const MERGE_FIELD_TYPES: ReadonlySet<string> = new Set([
  "MERGEFIELD",
  "ADDRESSBLOCK",
  "GREETINGLINE",
  "IF",
  "MERGEREC",
  "MERGESEQ",
  "NEXT",
  "NEXTIF",
  "SKIPIF",
  "SET",
  "ASK",
  "FILLIN",
]);

/** One recipient: its 1-based position in the full list, and its values. */
export interface MergeRecord {
  readonly number: number;
  readonly values: Readonly<Record<string, string>>;
}

interface MergeState {
  readonly records: readonly MergeRecord[];
  /** Index into `records` of the current record. */
  cursor: number;
  /** MERGESEQ: the sequence number of the first record of the current copy. */
  sequence: number;
  /** The cursor at the start of the current copy. */
  copyStart: number;
  readonly fieldMap: Readonly<Record<string, string>>;
  readonly bookmarks: Map<string, string>;
  readonly answers: Readonly<Record<string, string>>;
  skip: boolean;
}

function lookup(values: Readonly<Record<string, string>>, name: string): string | undefined {
  if (Object.hasOwn(values, name)) return values[name];
  const target = normalizeName(name);
  for (const [k, v] of Object.entries(values)) if (normalizeName(k) === target) return v;
  return undefined;
}

// ADDRESSBLOCK / GREETINGLINE tokens → Word address fields.
const TOKEN_FIELDS: Readonly<Record<string, AddressField>> = {
  TITLE0: "Courtesy Title",
  FIRST0: "First Name",
  MIDDLE0: "Middle Name",
  LAST0: "Last Name",
  SUFFIX0: "Suffix",
  NICK0: "Nickname",
  COMPANY: "Company",
  STREET1: "Address 1",
  STREET2: "Address 2",
  CITY: "City",
  STATE: "State",
  POSTAL: "Postal Code",
  COUNTRY: "Country or Region",
};

function addressValue(state: MergeState, field: AddressField): string {
  const record = state.records[state.cursor];
  if (!record) return "";
  const column = state.fieldMap[field];
  return (column !== undefined ? lookup(record.values, column) : undefined) ?? lookup(record.values, field) ?? "";
}

/**
 * Expand an ADDRESSBLOCK / GREETINGLINE format: each `<<before_TOKEN_after>>`
 * group prints only when its token has a value; `_BEFORE_` / `_AFTER_` groups
 * print when the name in between printed anything.
 */
function expandFormat(format: string, value: (token: string) => string | undefined): { text: string; named: boolean } {
  let named = false;
  const groups = [...format.matchAll(/<<(.*?)>>|([^<]+|<)/gs)];
  const pieces: Array<{ text: string; kind: "before" | "after" | "plain" }> = [];
  for (const g of groups) {
    if (g[1] === undefined) {
      pieces.push({ text: g[2] ?? "", kind: "plain" });
      continue;
    }
    const m = /^(.*?)_([A-Z0-9]+)_(.*)$/s.exec(g[1]);
    if (!m) continue;
    const token = m[2] as string;
    if (token === "BEFORE" || token === "AFTER") {
      pieces.push({ text: `${m[1]}${m[3]}`.trim(), kind: token === "BEFORE" ? "before" : "after" });
      continue;
    }
    const v = value(token);
    if (v) {
      if (token in TOKEN_FIELDS && /^(TITLE0|FIRST0|MIDDLE0|LAST0|SUFFIX0|NICK0)$/.test(token)) named = true;
      pieces.push({ text: `${m[1]}${v}${m[3]}`, kind: "plain" });
    }
  }
  const before = pieces.filter((p) => p.kind === "before").map((p) => p.text).join(" ");
  const after = pieces.filter((p) => p.kind === "after").map((p) => p.text).join("");
  const body = pieces.filter((p) => p.kind === "plain").map((p) => p.text).join("");
  const text = `${before ? `${before} ` : ""}${body}${after}`;
  return { text: text.replace(/\n+$/, ""), named };
}

const NUMERIC = /^-?\d+(\.\d+)?$/;

/** Word's IF comparison; `=` / `<>` with a quoted pattern accept `?` and `*` wildcards. */
function compare(left: string, op: string, right: string): boolean {
  if ((op === "=" || op === "<>") && /[?*]/.test(right)) {
    const re = new RegExp(`^${right.replace(/[.+^${}()|[\]\\]/g, "\\$&").replace(/\?/g, ".").replace(/\*/g, ".*")}$`, "s");
    return op === "=" ? re.test(left) : !re.test(left);
  }
  const numeric = NUMERIC.test(left.trim()) && NUMERIC.test(right.trim());
  const a = numeric ? Number(left) : left;
  const b = numeric ? Number(right) : right;
  switch (op) {
    case "=":
      return a === b;
    case "<>":
      return a !== b;
    case "<":
      return a < b;
    case "<=":
      return a <= b;
    case ">":
      return a > b;
    case ">=":
      return a >= b;
    default:
      return false;
  }
}

function applyTextFormat(text: string, parsed: FieldInstruction): string {
  let out = text;
  for (const s of parsed.switches) {
    if (s.name !== "*" || !s.arg) continue;
    switch (s.arg.toLowerCase()) {
      case "upper":
        out = out.toUpperCase();
        break;
      case "lower":
        out = out.toLowerCase();
        break;
      case "caps":
        out = out.replace(/\b\p{L}/gu, (c) => c.toUpperCase());
        break;
      case "firstcap":
        out = out.charAt(0).toUpperCase() + out.slice(1);
        break;
    }
  }
  return out;
}

/** The text a merge field shows for the state, or `undefined` to leave it alone. */
function evaluate(field: FlowField, state: MergeState | undefined): string | undefined {
  const p = field.parsed;
  const record = state?.records[state.cursor];
  switch (p.type) {
    case "MERGEFIELD": {
      const name = p.args[0] ?? "";
      if (!state) return `«${name}»`;
      // Past the last recipient (a label sheet with records left over) fields are blank.
      if (!record) return "";
      const value = lookup(record.values, name) ?? (name in state.fieldMap ? addressValue(state, name as AddressField) : "");
      if (!value) return "";
      return `${fieldSwitch(p, "b")?.arg ?? ""}${applyTextFormat(value, p)}${fieldSwitch(p, "f")?.arg ?? ""}`;
    }
    case "ADDRESSBLOCK": {
      if (!state) return "«AddressBlock»";
      if (!record) return "";
      const code = Number(fieldSwitch(p, "c")?.arg ?? 2);
      const excluded = fieldSwitch(p, "e")?.arg ?? "";
      return expandFormat(fieldSwitch(p, "f")?.arg ?? DEFAULT_NAME_FORMAT, (token) => {
        const f = TOKEN_FIELDS[token];
        if (!f) return undefined;
        const v = addressValue(state, f);
        if (token === "COUNTRY") {
          if (code === 0) return undefined;
          if (code === 2 && v.toLowerCase() === excluded.toLowerCase()) return undefined;
        }
        return v;
      }).text;
    }
    case "GREETINGLINE": {
      if (!state) return "«GreetingLine»";
      if (!record) return "";
      const { text, named } = expandFormat(fieldSwitch(p, "f")?.arg ?? "", (token) => {
        const f = TOKEN_FIELDS[token];
        return f ? addressValue(state, f) : undefined;
      });
      return named ? text : (fieldSwitch(p, "e")?.arg ?? "");
    }
    case "MERGEREC":
      if (!state) return "«Merge Record #»";
      return record ? formatFieldNumber(record.number, p) : "";
    case "MERGESEQ":
      if (!state) return "«Merge Sequence #»";
      return record ? formatFieldNumber(state.sequence + state.cursor - state.copyStart, p) : "";
    case "IF": {
      const [left = "", op = "=", right = "", whenTrue = "", whenFalse = ""] = p.args;
      return compare(left, op, right) ? whenTrue : whenFalse;
    }
    case "FILLIN": {
      const prompt = p.args[0] ?? "";
      return state?.answers[prompt] ?? fieldSwitch(p, "d")?.arg ?? "";
    }
    case "ASK": {
      const [bookmark = "", prompt = ""] = p.args;
      state?.bookmarks.set(bookmark, state.answers[prompt] ?? fieldSwitch(p, "d")?.arg ?? "");
      return "";
    }
    case "SET": {
      const [bookmark = "", value = ""] = p.args;
      state?.bookmarks.set(bookmark, value);
      return "";
    }
    case "NEXT":
      if (state) advance(state);
      return "";
    case "NEXTIF": {
      const [left = "", op = "=", right = ""] = p.args;
      if (state && compare(left, op, right)) advance(state);
      return "";
    }
    case "SKIPIF": {
      const [left = "", op = "=", right = ""] = p.args;
      if (state && compare(left, op, right)) state.skip = true;
      return "";
    }
    default:
      return undefined;
  }
}

function advance(state: MergeState): void {
  if (state.cursor < state.records.length) state.cursor++;
}

/**
 * Evaluate every merge field in the body in document order, innermost first
 * (so an IF sees its nested MERGEFIELD's value). NEXT / NEXTIF move the state's
 * cursor as they are reached.
 */
function evaluateMergeFields(body: Docx["document"]["body"], state: MergeState | undefined): void {
  withOpenContentControls(body, () => {
    const all = scanFields(paragraphFlow(body)).filter((f) => MERGE_FIELD_TYPES.has(f.parsed.type));
    const maxDepth = all.reduce((m, f) => Math.max(m, f.depth), 0);
    for (let depth = maxDepth; depth >= 0; depth--) {
      const fields = scanFields(paragraphFlow(body)).filter(
        (f) => f.depth === depth && MERGE_FIELD_TYPES.has(f.parsed.type),
      );
      // Results are computed in document order (NEXT advances the record for
      // what follows) and written back last to first, so indexes stay valid.
      const results = fields.map((f) => ({ f, text: evaluate(f, state) }));
      for (const { f, text } of results.toReversed()) {
        if (text === undefined || f.begin.ref !== f.end.ref) continue;
        setInlineResult(f, text ? inlinesFromXml(textRunXml(text)) : []);
      }
    }
  });
}

export interface PreviewOptions {
  /** Match Fields mapping; defaults to the document's stored mapping, else automatic. */
  readonly fieldMap?: Readonly<Record<string, string>>;
  /** Answers to ASK / FILLIN prompts, by prompt text (their defaults otherwise). */
  readonly answers?: Readonly<Record<string, string>>;
}

/**
 * Preview Results: show recipient `index` (0-based into `records`) in every
 * merge field, or the «Field» placeholders again with `index` undefined.
 * Records the state in `w:viewMergedData` / `w:activeRecord` when the document
 * is a merge main document.
 */
export function previewMailMerge(
  doc: Docx,
  list: RecipientList | undefined,
  index: number | undefined,
  options: PreviewOptions = {},
): void {
  const records = (list?.records ?? []).map((values, i) => ({ number: i + 1, values }));
  const state: MergeState | undefined =
    index === undefined || !list
      ? undefined
      : {
          records,
          cursor: Math.max(0, Math.min(index, records.length - 1)),
          sequence: 1,
          copyStart: Math.max(0, Math.min(index, records.length - 1)),
          fieldMap: options.fieldMap ?? fieldMapFor(doc, list.columns),
          bookmarks: new Map(),
          answers: options.answers ?? {},
          skip: false,
        };
  evaluateMergeFields(doc.document.body, state);
  doc.dirty = true;
  if (mailMergeSettings(doc)) {
    editMailMerge(doc, (mm) => {
      if (state) {
        setOrderedChild(mm, elementFromXml("<w:viewMergedData/>"), MAIL_MERGE_ORDER);
        setOrderedChild(mm, elementFromXml(`<w:activeRecord w:val="${state.cursor + 1}"/>`), MAIL_MERGE_ORDER);
      } else {
        const children = mm.children as XmlElement[];
        for (let i = children.length - 1; i >= 0; i--) {
          const c = children[i];
          if (c?.kind === "element" && (c.name.local === "viewMergedData" || c.name.local === "activeRecord")) {
            children.splice(i, 1);
          }
        }
      }
    });
  }
}

function fieldMapFor(doc: Docx, columns: readonly string[]): Readonly<Record<string, string>> {
  const stored = mailMergeFieldMap(doc);
  return Object.keys(stored).length > 0 ? stored : autoFieldMap(columns);
}

/** Check for Errors: MERGEFIELD names that are not columns of the recipient list. */
export function mergeFieldErrors(doc: Docx, list: RecipientList): string[] {
  const columns = new Set(list.columns.map(normalizeName));
  const mappable = new Set(ADDRESS_FIELDS.map(normalizeName));
  const missing = new Set<string>();
  withOpenContentControls(doc.document.body, () => {
    for (const f of scanFields(paragraphFlow(doc.document.body))) {
      if (f.parsed.type !== "MERGEFIELD") continue;
      const name = f.parsed.args[0] ?? "";
      const n = normalizeName(name);
      if (!columns.has(n) && !mappable.has(n)) missing.add(name);
    }
  });
  return [...missing];
}

/** The MERGEFIELD names the document uses, in order of first use. */
export function mergeFieldNames(doc: Docx): string[] {
  const names: string[] = [];
  withOpenContentControls(doc.document.body, () => {
    for (const f of scanFields(paragraphFlow(doc.document.body))) {
      const name = f.parsed.args[0];
      if (f.parsed.type === "MERGEFIELD" && name && !names.includes(name)) names.push(name);
    }
  });
  return names;
}

/** Turn every merge field into its result text (the merged document keeps no merge codes). */
function unlinkMergeFields(body: Docx["document"]["body"]): void {
  withOpenContentControls(body, () => {
    // Top-level fields only (nested codes go with their outer field), last to
    // first so earlier positions stay valid.
    const fields = scanFields(paragraphFlow(body)).filter(
      (f) => f.depth === 0 && MERGE_FIELD_TYPES.has(f.parsed.type) && f.begin.ref === f.end.ref,
    );
    for (const field of fields.toReversed()) {
      const children = field.begin.ref.paragraph.children;
      const result =
        field.separate && field.separate.ref === field.end.ref
          ? children.slice(field.separate.inline + 1, field.end.inline)
          : [];
      children.splice(field.begin.inline, field.end.inline - field.begin.inline + 1, ...result);
    }
  });
}

export interface MergeOptions extends PreviewOptions {
  /** First and last record numbers to merge (1-based, inclusive); Merge Range. */
  readonly from?: number;
  readonly to?: number;
  /** Which records take part (Edit Recipient List checkboxes); all by default. */
  readonly included?: readonly boolean[];
}

/**
 * Finish & Merge ▸ Edit Individual Documents: a new document with one copy of
 * the main document per recipient (per sheet of labels for label documents),
 * merge fields replaced by the recipients' values, and each copy in its own
 * section starting on a new page. A Directory (`catalog`) runs the copies on
 * without breaks. Records a SKIPIF rule skips are left out.
 */
export function mergeToNewDocument(doc: Docx, list: RecipientList, options: MergeOptions = {}): Docx {
  const from = Math.max(1, options.from ?? 1);
  const to = Math.min(list.records.length, options.to ?? list.records.length);
  const records: MergeRecord[] = [];
  for (let n = from; n <= to; n++) {
    if (options.included && options.included[n - 1] === false) continue;
    const values = list.records[n - 1];
    if (values) records.push({ number: n, values });
  }
  const template = openDocx(toUint8Array(doc));
  const type = mailMergeSettings(template)?.type ?? "formLetters";
  const sectPr = ensureBodySectPr(template);
  const state: MergeState = {
    records,
    cursor: 0,
    sequence: 1,
    copyStart: 0,
    fieldMap: options.fieldMap ?? fieldMapFor(doc, list.columns),
    bookmarks: new Map(),
    answers: options.answers ?? {},
    skip: false,
  };
  const out: WmlBlock[] = [];
  while (state.cursor < records.length) {
    const start = state.cursor;
    const body = { ...template.document.body, blocks: structuredClone(template.document.body.blocks) };
    state.skip = false;
    state.copyStart = start;
    evaluateMergeFields(body, state);
    unlinkMergeFields(body);
    // The copy used the records up to the cursor (NEXT fields move it).
    state.cursor++;
    if (state.skip) continue;
    state.sequence += state.cursor - start;
    if (out.length > 0 && type !== "catalog") {
      // A next-page section break ends the previous copy.
      const brk = structuredClone(sectPr);
      setOrderedChild(brk, elementFromXml('<w:type w:val="nextPage"/>'), SECT_PR_ORDER);
      out.push(...paragraphsFromXml(`<w:p><w:pPr>${elementXml(brk)}</w:pPr></w:p>`));
    }
    out.push(...body.blocks);
  }
  const merged = openDocx(toUint8Array(template));
  merged.document.body.blocks = out;
  setMailMergeDocumentType(merged, "normal");
  merged.dirty = true;
  return merged;
}

// --- Envelopes ------------------------------------------------------------------------

/** Envelope sizes (Envelope Options ▸ Envelope size), width × height in twips, portrait. */
export const ENVELOPE_SIZES = {
  "Size 10": { width: 5940, height: 13680 },
  "Size 9": { width: 5580, height: 12780 },
  "Size 6 3/4": { width: 5220, height: 9360 },
  Monarch: { width: 5580, height: 10800 },
  DL: { width: 6237, height: 12474 },
  C4: { width: 12984, height: 18371 },
  C5: { width: 9185, height: 12984 },
  C6: { width: 6464, height: 9185 },
  B5: { width: 9979, height: 14175 },
} as const;
export type EnvelopeSize = keyof typeof ENVELOPE_SIZES;

export interface EnvelopeOptions {
  /** Delivery address; lines separated by `\n`. */
  readonly deliveryAddress: string;
  /** Return address; lines separated by `\n`. */
  readonly returnAddress?: string;
  readonly omitReturnAddress?: boolean;
  readonly size?: EnvelopeSize;
}

const ENVELOPE_STYLES = {
  EnvelopeAddress:
    '<w:style w:type="paragraph" w:styleId="EnvelopeAddress"><w:name w:val="envelope address"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:pPr><w:framePr w:w="7920" w:h="1980" w:hRule="exact" w:hSpace="180" w:wrap="auto" w:hAnchor="page" w:xAlign="center" w:yAlign="bottom"/><w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:ind w:left="2880"/></w:pPr><w:rPr><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:style>',
  EnvelopeReturn:
    '<w:style w:type="paragraph" w:styleId="EnvelopeReturn"><w:name w:val="envelope return"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style>',
};

const ENVELOPE_MARGINS =
  '<w:pgMar w:top="360" w:right="720" w:bottom="720" w:left="576" w:header="720" w:footer="720" w:gutter="0"/>';

/**
 * Envelopes ▸ Add to Document: an envelope section at the start of the
 * document (landscape page of the envelope size, return address top left,
 * delivery address in Word's Envelope Address frame), replacing an envelope
 * added before. The rest of the document starts on a new page after it.
 */
export function addEnvelope(doc: Docx, options: EnvelopeOptions): void {
  const delivery = options.deliveryAddress.trim();
  if (!delivery) throw new Error("The envelope needs a delivery address.");
  const size = ENVELOPE_SIZES[options.size ?? "Size 10"];
  ensureBuiltInStyles(doc, ENVELOPE_STYLES);
  ensureBodySectPr(doc);
  removeEnvelope(doc);
  const lines = (text: string, style: string) =>
    text
      .split(/\r?\n/)
      .map((line) => `<w:p><w:pPr><w:pStyle w:val="${style}"/></w:pPr>${textRunXml(line)}</w:p>`)
      .join("");
  const returnXml = options.omitReturnAddress || !options.returnAddress?.trim() ? "" : lines(options.returnAddress.trim(), "EnvelopeReturn");
  const sectPr = `<w:sectPr><w:pgSz w:w="${size.height}" w:h="${size.width}" w:orient="landscape"/>${ENVELOPE_MARGINS}<w:cols w:space="720"/></w:sectPr>`;
  const blocks = paragraphsFromXml(`${returnXml}${lines(delivery, "EnvelopeAddress")}<w:p><w:pPr>${sectPr}</w:pPr></w:p>`);
  doc.document.body.blocks.unshift(...blocks);
  doc.dirty = true;
}

/** Remove an envelope section added by {@link addEnvelope}. Returns whether one was found. */
export function removeEnvelope(doc: Docx): boolean {
  const blocks = doc.document.body.blocks;
  const end = blocks.findIndex((b) => b.kind === "paragraph" && !!childElement(b.pPr, "sectPr"));
  if (end < 0) return false;
  const first = blocks.slice(0, end + 1);
  const isEnvelope = first.some(
    (b) => b.kind === "paragraph" && wAttr(childElement(b.pPr, "pStyle"), "val") === "EnvelopeAddress",
  );
  if (!isEnvelope) return false;
  blocks.splice(0, end + 1);
  doc.dirty = true;
  return true;
}

// --- Labels ---------------------------------------------------------------------------

/** A label sheet: page, label size, pitch and margins, in twips. */
export interface LabelProduct {
  readonly name: string;
  readonly page: { readonly width: number; readonly height: number };
  readonly across: number;
  readonly down: number;
  readonly labelWidth: number;
  readonly labelHeight: number;
  readonly horizontalPitch: number;
  readonly verticalPitch: number;
  readonly topMargin: number;
  readonly sideMargin: number;
}

const LETTER = { width: 12240, height: 15840 };
const A4 = { width: 11906, height: 16838 };
const mm = (v: number) => Math.round((v * 1440) / 25.4);
const inch = (v: number) => Math.round(v * 1440);

/** Common label products (Avery US Letter and A4 sheets). */
export const LABEL_PRODUCTS: readonly LabelProduct[] = [
  { name: "Avery 5160 Address Labels", page: LETTER, across: 3, down: 10, labelWidth: inch(2.625), labelHeight: inch(1), horizontalPitch: inch(2.75), verticalPitch: inch(1), topMargin: inch(0.5), sideMargin: inch(0.1875) },
  { name: "Avery 5161 Address Labels", page: LETTER, across: 2, down: 10, labelWidth: inch(4), labelHeight: inch(1), horizontalPitch: inch(4.1875), verticalPitch: inch(1), topMargin: inch(0.5), sideMargin: inch(0.15625) },
  { name: "Avery 5162 Address Labels", page: LETTER, across: 2, down: 7, labelWidth: inch(4), labelHeight: 1920, horizontalPitch: inch(4.1875), verticalPitch: 1920, topMargin: 1200, sideMargin: inch(0.15625) },
  { name: "Avery 5163 Shipping Labels", page: LETTER, across: 2, down: 5, labelWidth: inch(4), labelHeight: inch(2), horizontalPitch: inch(4.1875), verticalPitch: inch(2), topMargin: inch(0.5), sideMargin: inch(0.15625) },
  { name: "Avery 5164 Shipping Labels", page: LETTER, across: 2, down: 3, labelWidth: inch(4), labelHeight: 4800, horizontalPitch: inch(4.1875), verticalPitch: 4800, topMargin: inch(0.5), sideMargin: inch(0.15625) },
  { name: "Avery 5167 Return Address Labels", page: LETTER, across: 4, down: 20, labelWidth: inch(1.75), labelHeight: inch(0.5), horizontalPitch: inch(2.0625), verticalPitch: inch(0.5), topMargin: inch(0.5), sideMargin: inch(0.28125) },
  { name: "Avery L7160 Address Labels", page: A4, across: 3, down: 7, labelWidth: mm(63.5), labelHeight: mm(38.1), horizontalPitch: mm(66.04), verticalPitch: mm(38.1), topMargin: mm(15.15), sideMargin: mm(7.25) },
  { name: "Avery L7161 Address Labels", page: A4, across: 3, down: 6, labelWidth: mm(63.5), labelHeight: mm(46.6), horizontalPitch: mm(66.04), verticalPitch: mm(46.6), topMargin: mm(8.8), sideMargin: mm(7.25) },
  { name: "Avery L7163 Address Labels", page: A4, across: 2, down: 7, labelWidth: mm(99.1), labelHeight: mm(38.1), horizontalPitch: mm(101.6), verticalPitch: mm(38.1), topMargin: mm(15.15), sideMargin: mm(4.65) },
  { name: "Avery L7173 Shipping Labels", page: A4, across: 2, down: 5, labelWidth: mm(99.1), labelHeight: mm(57), horizontalPitch: mm(101.6), verticalPitch: mm(57), topMargin: mm(6), sideMargin: mm(4.65) },
  { name: "Avery L7651 Mini Labels", page: A4, across: 5, down: 13, labelWidth: mm(38.1), labelHeight: mm(21.2), horizontalPitch: mm(40.6), verticalPitch: mm(21.2), topMargin: mm(10.7), sideMargin: mm(4.75) },
];

export interface LabelOptions {
  /** Label text; lines separated by `\n`. Ignored when `mailMerge` is set. */
  readonly text: string;
  readonly product: LabelProduct;
  /** Single label: only this label (1-based row and column) is printed. */
  readonly single?: { readonly row: number; readonly column: number };
  /**
   * Labels for a mail merge: the first label holds `text` (merge fields can
   * be added there) and every other label starts with a NEXT field.
   */
  readonly mailMerge?: boolean;
}

// Word's label cells keep the text clear of the label edge.
const LABEL_TEXT_INDENT = 95;

/**
 * Labels ▸ New Document: a document whose single table lays out one sheet of
 * the label product — fixed column widths with spacer columns for the gaps,
 * exact row heights — with the label text in every label (or one).
 */
export function createLabelDocument(options: LabelOptions): Docx {
  const p = options.product;
  if (p.across < 1 || p.down < 1) throw new Error("A label sheet needs at least one label.");
  if (options.single && (options.single.row < 1 || options.single.row > p.down || options.single.column < 1 || options.single.column > p.across)) {
    throw new Error(`Label position must be within ${p.down} rows and ${p.across} columns.`);
  }
  const doc = createDocx({ paragraphs: [] });
  const gap = Math.max(0, p.horizontalPitch - p.labelWidth);
  const vGap = Math.max(0, p.verticalPitch - p.labelHeight);
  const right = Math.max(0, p.page.width - p.sideMargin - (p.across - 1) * p.horizontalPitch - p.labelWidth);
  const grid: number[] = [];
  for (let c = 0; c < p.across; c++) {
    grid.push(p.labelWidth);
    if (c < p.across - 1 && gap > 0) grid.push(gap);
  }
  const textXml = (lines: string): string =>
    lines
      .split(/\r?\n/)
      .map((line) => `<w:p><w:pPr><w:spacing w:before="0" w:after="0"/><w:ind w:left="${LABEL_TEXT_INDENT}" w:right="${LABEL_TEXT_INDENT}"/></w:pPr>${textRunXml(line)}</w:p>`)
      .join("");
  const emptyP = '<w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr></w:p>';
  const rows: string[] = [];
  for (let r = 1; r <= p.down; r++) {
    const cells: string[] = [];
    for (let c = 1; c <= p.across; c++) {
      const first = r === 1 && c === 1;
      let content: string;
      if (options.mailMerge) {
        content = first ? textXml(options.text) : `<w:p><w:pPr><w:spacing w:before="0" w:after="0"/><w:ind w:left="${LABEL_TEXT_INDENT}" w:right="${LABEL_TEXT_INDENT}"/></w:pPr>${resultlessFieldRunsXml("NEXT")}</w:p>`;
      } else if (!options.single || (options.single.row === r && options.single.column === c)) {
        content = textXml(options.text);
      } else {
        content = emptyP;
      }
      cells.push(`<w:tc><w:tcPr><w:tcW w:w="${p.labelWidth}" w:type="dxa"/></w:tcPr>${content}</w:tc>`);
      if (c < p.across && gap > 0) {
        cells.push(`<w:tc><w:tcPr><w:tcW w:w="${gap}" w:type="dxa"/></w:tcPr>${emptyP}</w:tc>`);
      }
    }
    rows.push(`<w:tr><w:trPr><w:trHeight w:val="${p.labelHeight}" w:hRule="exact"/></w:trPr>${cells.join("")}</w:tr>`);
    if (r < p.down && vGap > 0) {
      const spacer = grid.map((w) => `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/></w:tcPr>${emptyP}</w:tc>`).join("");
      rows.push(`<w:tr><w:trPr><w:trHeight w:val="${vGap}" w:hRule="exact"/></w:trPr>${spacer}</w:tr>`);
    }
  }
  const table = `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="15" w:type="dxa"/><w:right w:w="15" w:type="dxa"/></w:tblCellMar><w:tblLook w:val="0000"/></w:tblPr><w:tblGrid>${grid.map((w) => `<w:gridCol w:w="${w}"/>`).join("")}</w:tblGrid>${rows.join("")}</w:tbl>`;
  doc.document.body.blocks = blocksFromXml(`${table}<w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr></w:p>`);
  doc.document.body.sectPr = elementFromXml(
    `<w:sectPr><w:pgSz w:w="${p.page.width}" w:h="${p.page.height}"/><w:pgMar w:top="${p.topMargin}" w:right="${right}" w:bottom="0" w:left="${p.sideMargin}" w:header="720" w:footer="720" w:gutter="0"/><w:cols w:space="720"/></w:sectPr>`,
  );
  if (options.mailMerge) setMailMergeDocumentType(doc, "mailingLabels");
  doc.dirty = true;
  return doc;
}

/**
 * Update Labels: copy the first label's content into every other label of
 * the label table, each preceded by a NEXT field so it takes the next
 * recipient. Returns the number of labels updated.
 */
export function updateLabels(doc: Docx): number {
  const table = doc.document.body.blocks.find((b) => b.kind === "table");
  if (!table || table.kind !== "table") return 0;
  const firstCell = table.rows[0]?.cells[0];
  if (!firstCell) return 0;
  const width = wAttr(childElement(firstCell.tcPr, "tcW"), "w");
  let count = 0;
  table.rows.forEach((row, r) => {
    row.cells.forEach((cell, c) => {
      if (r === 0 && c === 0) return;
      // Spacer columns and rows are narrower or empty-height; only cells of
      // the label width are labels.
      if (wAttr(childElement(cell.tcPr, "tcW"), "w") !== width) return;
      if (row.cells.length > 0 && isSpacerRow(row.trPr, table.rows[0]?.trPr)) return;
      const copy = structuredClone(firstCell.paragraphs);
      const [head] = copy;
      const next = inlinesFromXml(resultlessFieldRunsXml("NEXT"));
      if (head) head.children = [...next, ...stripNext(head.children)];
      cell.paragraphs = copy;
      count++;
    });
  });
  doc.dirty = true;
  return count;
}

function isSpacerRow(trPr: XmlElement | undefined, labelRow: XmlElement | undefined): boolean {
  const height = (el: XmlElement | undefined) => wAttr(childElement(el, "trHeight"), "val");
  return height(trPr) !== height(labelRow);
}

function stripNext(inlines: readonly WmlInline[]): WmlInline[] {
  const p: WmlParagraph = { kind: "paragraph", children: [...inlines], extras: [] };
  const fields = scanFields([{ list: [p], index: 0, paragraph: p }]).filter((f) => f.parsed.type === "NEXT");
  for (const f of fields.toReversed()) p.children.splice(f.begin.inline, f.end.inline - f.begin.inline + 1);
  return p.children;
}
