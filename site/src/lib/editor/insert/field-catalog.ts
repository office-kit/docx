/**
 * Every field type ECMA-376 defines (Part 1 §17.16.5), grouped into Word's
 * Field dialog categories, with what the dialog offers for each: the syntax
 * line, the argument it asks for, and its field-specific switches. Anything
 * else can be typed into the dialog's field-code box.
 */

export type FieldCategory =
  | "dateTime"
  | "automation"
  | "info"
  | "formulas"
  | "index"
  | "links"
  | "mailMerge"
  | "numbering"
  | "user"
  | "forms";

export const FIELD_CATEGORIES: readonly FieldCategory[] = [
  "dateTime",
  "automation",
  "info",
  "formulas",
  "index",
  "links",
  "mailMerge",
  "numbering",
  "user",
  "forms",
];

/** What the field's first argument is (picked from a list, or typed). */
export type FieldArgument =
  | { readonly kind: "bookmark" }
  | { readonly kind: "property"; readonly choices: readonly string[] }
  | { readonly kind: "style" }
  | {
      readonly kind: "text";
      readonly label: "identifier" | "formula" | "text" | "code";
      /** Written as-is rather than quoted (expressions, formulas). */
      readonly raw?: boolean;
    }
  | { readonly kind: "form"; readonly form: "text" | "checkBox" | "dropDown" };

/** Which `\*` / `\#` / `\@` choices apply to the result. */
export type ResultKind = "date" | "number" | "text" | "none";

export interface FieldDef {
  readonly name: string;
  readonly category: FieldCategory;
  readonly syntax: string;
  readonly arg?: FieldArgument;
  /** Field-specific switches (without values; the user completes them in the code box). */
  readonly switches?: readonly string[];
  readonly result: ResultKind;
}

// DOCPROPERTY names Word lists (core, extended and statistics properties).
const DOC_PROPERTIES = [
  "Author",
  "Category",
  "Characters",
  "CharactersWithSpaces",
  "Comments",
  "Company",
  "CreateTime",
  "Keywords",
  "LastSavedBy",
  "LastSavedTime",
  "Lines",
  "Manager",
  "Pages",
  "Paragraphs",
  "RevisionNumber",
  "Subject",
  "Template",
  "Title",
  "Words",
] as const;

// INFO's info types (§17.16.5.27).
const INFO_TYPES = [
  "Author",
  "Comments",
  "CreateDate",
  "EditTime",
  "FileName",
  "FileSize",
  "Keywords",
  "LastSavedBy",
  "NumChars",
  "NumPages",
  "NumWords",
  "PrintDate",
  "RevNum",
  "SaveDate",
  "Subject",
  "Template",
  "Title",
] as const;

export const FIELDS: readonly FieldDef[] = [
  // Date and Time
  {
    name: "CREATEDATE",
    category: "dateTime",
    syntax: 'CREATEDATE [\\@ "Date-Time Picture"]',
    result: "date",
  },
  {
    name: "DATE",
    category: "dateTime",
    syntax: 'DATE [\\@ "Date-Time Picture"] [\\h | \\l | \\s]',
    switches: ["\\l", "\\h", "\\s"],
    result: "date",
  },
  { name: "EDITTIME", category: "dateTime", syntax: "EDITTIME", result: "number" },
  {
    name: "PRINTDATE",
    category: "dateTime",
    syntax: 'PRINTDATE [\\@ "Date-Time Picture"]',
    result: "date",
  },
  {
    name: "SAVEDATE",
    category: "dateTime",
    syntax: 'SAVEDATE [\\@ "Date-Time Picture"]',
    result: "date",
  },
  { name: "TIME", category: "dateTime", syntax: 'TIME [\\@ "Date-Time Picture"]', result: "date" },
  // Document Automation
  {
    name: "COMPARE",
    category: "automation",
    syntax: "COMPARE Expression1 Operator Expression2",
    arg: { kind: "text", label: "formula", raw: true },
    result: "number",
  },
  {
    name: "DOCVARIABLE",
    category: "automation",
    syntax: 'DOCVARIABLE "Name"',
    arg: { kind: "text", label: "identifier" },
    result: "text",
  },
  {
    name: "GOTOBUTTON",
    category: "automation",
    syntax: "GOTOBUTTON Destination DisplayText",
    arg: { kind: "text", label: "text", raw: true },
    result: "none",
  },
  {
    name: "IF",
    category: "automation",
    syntax: 'IF Expression1 Operator Expression2 "TrueText" "FalseText"',
    arg: { kind: "text", label: "formula", raw: true },
    result: "text",
  },
  {
    name: "MACROBUTTON",
    category: "automation",
    syntax: "MACROBUTTON MacroName DisplayText",
    arg: { kind: "text", label: "text", raw: true },
    result: "none",
  },
  {
    name: "PRINT",
    category: "automation",
    syntax: 'PRINT "PrinterInstructions"',
    arg: { kind: "text", label: "text" },
    switches: ["\\p"],
    result: "none",
  },
  { name: "PRIVATE", category: "automation", syntax: "PRIVATE", result: "none" },
  // Document Information
  { name: "AUTHOR", category: "info", syntax: 'AUTHOR ["NewName"]', result: "text" },
  { name: "COMMENTS", category: "info", syntax: 'COMMENTS ["NewComments"]', result: "text" },
  {
    name: "DOCPROPERTY",
    category: "info",
    syntax: 'DOCPROPERTY "Name"',
    arg: { kind: "property", choices: DOC_PROPERTIES },
    result: "text",
  },
  {
    name: "FILENAME",
    category: "info",
    syntax: "FILENAME [\\p]",
    switches: ["\\p"],
    result: "text",
  },
  {
    name: "FILESIZE",
    category: "info",
    syntax: "FILESIZE [\\k | \\m]",
    switches: ["\\k", "\\m"],
    result: "number",
  },
  {
    name: "INFO",
    category: "info",
    syntax: 'INFO [InfoType] ["NewValue"]',
    arg: { kind: "property", choices: INFO_TYPES },
    result: "text",
  },
  { name: "KEYWORDS", category: "info", syntax: 'KEYWORDS ["NewKeywords"]', result: "text" },
  { name: "LASTSAVEDBY", category: "info", syntax: "LASTSAVEDBY", result: "text" },
  { name: "NUMCHARS", category: "info", syntax: "NUMCHARS", result: "number" },
  { name: "NUMPAGES", category: "info", syntax: "NUMPAGES", result: "number" },
  { name: "NUMWORDS", category: "info", syntax: "NUMWORDS", result: "number" },
  { name: "SUBJECT", category: "info", syntax: 'SUBJECT ["NewSubject"]', result: "text" },
  {
    name: "TEMPLATE",
    category: "info",
    syntax: "TEMPLATE [\\p]",
    switches: ["\\p"],
    result: "text",
  },
  { name: "TITLE", category: "info", syntax: 'TITLE ["NewTitle"]', result: "text" },
  // Equations and Formulas
  {
    name: "=",
    category: "formulas",
    syntax: "= Formula",
    arg: { kind: "text", label: "formula", raw: true },
    result: "number",
  },
  {
    name: "ADVANCE",
    category: "formulas",
    syntax: "ADVANCE [\\d n] [\\l n] [\\r n] [\\u n] [\\x n] [\\y n]",
    switches: ["\\d", "\\l", "\\r", "\\u", "\\x", "\\y"],
    result: "none",
  },
  {
    name: "EQ",
    category: "formulas",
    syntax: "EQ Instructions",
    arg: { kind: "text", label: "formula", raw: true },
    result: "none",
  },
  {
    name: "SYMBOL",
    category: "formulas",
    syntax: 'SYMBOL CharNum [\\f "Font"] [\\s Size] [\\a | \\j | \\u] [\\h]',
    arg: { kind: "text", label: "code", raw: true },
    switches: ["\\f", "\\s", "\\a", "\\j", "\\u", "\\h"],
    result: "none",
  },
  // Index and Tables
  {
    name: "BIBLIOGRAPHY",
    category: "index",
    syntax: "BIBLIOGRAPHY [\\l LCID]",
    switches: ["\\l"],
    result: "none",
  },
  {
    name: "CITATION",
    category: "index",
    syntax: "CITATION Tag [\\l LCID] [\\p Page]",
    arg: { kind: "text", label: "identifier" },
    switches: ["\\l", "\\p", "\\s", "\\v", "\\f", "\\m", "\\n", "\\t", "\\y"],
    result: "none",
  },
  {
    name: "INDEX",
    category: "index",
    syntax: "INDEX [Switches]",
    switches: [
      "\\b",
      "\\c",
      "\\d",
      "\\e",
      "\\f",
      "\\g",
      "\\h",
      "\\k",
      "\\l",
      "\\p",
      "\\r",
      "\\s",
      "\\y",
      "\\z",
    ],
    result: "none",
  },
  {
    name: "TA",
    category: "index",
    syntax: 'TA [\\l "Long"] [\\s "Short"] [\\c Category] [\\b] [\\i] [\\r Bookmark]',
    switches: ["\\l", "\\s", "\\c", "\\b", "\\i", "\\r"],
    result: "none",
  },
  {
    name: "TC",
    category: "index",
    syntax: 'TC "Text" [\\f Type] [\\l Level] [\\n]',
    arg: { kind: "text", label: "text" },
    switches: ["\\f", "\\l", "\\n"],
    result: "none",
  },
  {
    name: "TOA",
    category: "index",
    syntax: "TOA \\c Category [Switches]",
    switches: ["\\c", "\\b", "\\d", "\\e", "\\f", "\\g", "\\h", "\\l", "\\p", "\\s"],
    result: "none",
  },
  {
    name: "TOC",
    category: "index",
    syntax: "TOC [Switches]",
    switches: [
      "\\o",
      "\\h",
      "\\z",
      "\\u",
      "\\t",
      "\\b",
      "\\c",
      "\\d",
      "\\f",
      "\\l",
      "\\n",
      "\\p",
      "\\s",
      "\\w",
      "\\x",
      "\\a",
    ],
    result: "none",
  },
  {
    name: "XE",
    category: "index",
    syntax: 'XE "Text" [\\b] [\\f Type] [\\i] [\\r Bookmark] [\\t "Text"] [\\y "Yomi"]',
    arg: { kind: "text", label: "text" },
    switches: ["\\b", "\\f", "\\i", "\\r", "\\t", "\\y"],
    result: "none",
  },
  // Links and References
  {
    name: "AUTOTEXT",
    category: "links",
    syntax: "AUTOTEXT AutoTextEntry",
    arg: { kind: "text", label: "identifier" },
    result: "text",
  },
  {
    name: "AUTOTEXTLIST",
    category: "links",
    syntax: 'AUTOTEXTLIST "LiteralText" [\\s StyleName] [\\t "TipText"]',
    arg: { kind: "text", label: "text" },
    switches: ["\\s", "\\t"],
    result: "text",
  },
  {
    name: "HYPERLINK",
    category: "links",
    syntax: 'HYPERLINK "Filename" [\\l "Location"] [\\m] [\\n] [\\o "ScreenTip"] [\\t "Frame"]',
    arg: { kind: "text", label: "text" },
    switches: ["\\l", "\\m", "\\n", "\\o", "\\t"],
    result: "text",
  },
  {
    name: "INCLUDEPICTURE",
    category: "links",
    syntax: 'INCLUDEPICTURE "FileName" [\\c Converter] [\\d]',
    arg: { kind: "text", label: "text" },
    switches: ["\\c", "\\d"],
    result: "none",
  },
  {
    name: "INCLUDETEXT",
    category: "links",
    syntax: 'INCLUDETEXT "FileName" [Bookmark] [\\c Converter] [\\!]',
    arg: { kind: "text", label: "text" },
    switches: ["\\c", "\\!"],
    result: "text",
  },
  {
    name: "LINK",
    category: "links",
    syntax: 'LINK ProgId "FileName" [PlaceReference] [Switches]',
    arg: { kind: "text", label: "text", raw: true },
    switches: ["\\a", "\\b", "\\d", "\\f", "\\h", "\\p", "\\r", "\\t", "\\u"],
    result: "none",
  },
  {
    name: "NOTEREF",
    category: "links",
    syntax: "NOTEREF Bookmark [\\f] [\\h] [\\p]",
    arg: { kind: "bookmark" },
    switches: ["\\f", "\\h", "\\p"],
    result: "number",
  },
  {
    name: "PAGEREF",
    category: "links",
    syntax: "PAGEREF Bookmark [\\h] [\\p]",
    arg: { kind: "bookmark" },
    switches: ["\\h", "\\p"],
    result: "number",
  },
  {
    name: "QUOTE",
    category: "links",
    syntax: 'QUOTE "LiteralText"',
    arg: { kind: "text", label: "text" },
    result: "text",
  },
  {
    name: "REF",
    category: "links",
    syntax: "REF Bookmark [\\d Separator] [\\f] [\\h] [\\n] [\\p] [\\r] [\\t] [\\w]",
    arg: { kind: "bookmark" },
    switches: ["\\d", "\\f", "\\h", "\\n", "\\p", "\\r", "\\t", "\\w"],
    result: "text",
  },
  {
    name: "STYLEREF",
    category: "links",
    syntax: 'STYLEREF "StyleIdentifier" [\\l] [\\n] [\\p] [\\r] [\\t] [\\w]',
    arg: { kind: "style" },
    switches: ["\\l", "\\n", "\\p", "\\r", "\\t", "\\w"],
    result: "text",
  },
  // Mail Merge
  {
    name: "ADDRESSBLOCK",
    category: "mailMerge",
    syntax: "ADDRESSBLOCK [Switches]",
    switches: ["\\c", "\\d", "\\e", "\\f", "\\l"],
    result: "none",
  },
  {
    name: "ASK",
    category: "mailMerge",
    syntax: 'ASK Bookmark "Prompt" [\\d "Default"] [\\o]',
    arg: { kind: "text", label: "text", raw: true },
    switches: ["\\d", "\\o"],
    result: "none",
  },
  {
    name: "DATABASE",
    category: "mailMerge",
    syntax: "DATABASE [Switches]",
    switches: ["\\b", "\\c", "\\d", "\\f", "\\h", "\\l", "\\o", "\\s", "\\t"],
    result: "none",
  },
  {
    name: "FILLIN",
    category: "mailMerge",
    syntax: 'FILLIN ["Prompt"] [\\d "Default"] [\\o]',
    arg: { kind: "text", label: "text" },
    switches: ["\\d", "\\o"],
    result: "text",
  },
  {
    name: "GREETINGLINE",
    category: "mailMerge",
    syntax: "GREETINGLINE [Switches]",
    switches: ["\\e", "\\f", "\\l"],
    result: "none",
  },
  {
    name: "MERGEFIELD",
    category: "mailMerge",
    syntax: 'MERGEFIELD FieldName [\\b "Text"] [\\f "Text"] [\\m] [\\v]',
    arg: { kind: "text", label: "identifier" },
    switches: ["\\b", "\\f", "\\m", "\\v"],
    result: "text",
  },
  { name: "MERGEREC", category: "mailMerge", syntax: "MERGEREC", result: "number" },
  { name: "MERGESEQ", category: "mailMerge", syntax: "MERGESEQ", result: "number" },
  { name: "NEXT", category: "mailMerge", syntax: "NEXT", result: "none" },
  {
    name: "NEXTIF",
    category: "mailMerge",
    syntax: "NEXTIF Expression1 Operator Expression2",
    arg: { kind: "text", label: "formula", raw: true },
    result: "none",
  },
  {
    name: "SET",
    category: "mailMerge",
    syntax: 'SET Bookmark "Text"',
    arg: { kind: "text", label: "text", raw: true },
    result: "none",
  },
  {
    name: "SKIPIF",
    category: "mailMerge",
    syntax: "SKIPIF Expression1 Operator Expression2",
    arg: { kind: "text", label: "formula", raw: true },
    result: "none",
  },
  // Numbering
  {
    name: "AUTONUM",
    category: "numbering",
    syntax: "AUTONUM [\\s Separator]",
    switches: ["\\s"],
    result: "number",
  },
  {
    name: "AUTONUMLGL",
    category: "numbering",
    syntax: "AUTONUMLGL [\\e] [\\s Separator]",
    switches: ["\\e", "\\s"],
    result: "number",
  },
  { name: "AUTONUMOUT", category: "numbering", syntax: "AUTONUMOUT", result: "number" },
  {
    name: "BARCODE",
    category: "numbering",
    syntax: 'BARCODE "Text" [\\b] [\\f "Code"] [\\u]',
    arg: { kind: "text", label: "text" },
    switches: ["\\b", "\\f", "\\u"],
    result: "none",
  },
  { name: "BIDIOUTLINE", category: "numbering", syntax: "BIDIOUTLINE", result: "number" },
  {
    name: "LISTNUM",
    category: "numbering",
    syntax: 'LISTNUM ["Name"] [\\l Level] [\\s StartAt]',
    switches: ["\\l", "\\s"],
    result: "number",
  },
  { name: "PAGE", category: "numbering", syntax: "PAGE", result: "number" },
  { name: "REVNUM", category: "numbering", syntax: "REVNUM", result: "number" },
  { name: "SECTION", category: "numbering", syntax: "SECTION", result: "number" },
  { name: "SECTIONPAGES", category: "numbering", syntax: "SECTIONPAGES", result: "number" },
  {
    name: "SEQ",
    category: "numbering",
    syntax: "SEQ Identifier [Bookmark] [\\c] [\\h] [\\n] [\\r n] [\\s level]",
    arg: { kind: "text", label: "identifier" },
    switches: ["\\c", "\\h", "\\n", "\\r", "\\s"],
    result: "number",
  },
  // User Information
  { name: "USERADDRESS", category: "user", syntax: 'USERADDRESS ["NewAddress"]', result: "text" },
  {
    name: "USERINITIALS",
    category: "user",
    syntax: 'USERINITIALS ["NewInitials"]',
    result: "text",
  },
  { name: "USERNAME", category: "user", syntax: 'USERNAME ["NewName"]', result: "text" },
  // Form Fields (legacy, §17.16.5.20–22)
  {
    name: "FORMCHECKBOX",
    category: "forms",
    syntax: "FORMCHECKBOX",
    arg: { kind: "form", form: "checkBox" },
    result: "none",
  },
  {
    name: "FORMDROPDOWN",
    category: "forms",
    syntax: "FORMDROPDOWN",
    arg: { kind: "form", form: "dropDown" },
    result: "none",
  },
  {
    name: "FORMTEXT",
    category: "forms",
    syntax: "FORMTEXT",
    arg: { kind: "form", form: "text" },
    result: "none",
  },
];

/** `\*` text formats (§17.16.4.3). */
export const TEXT_FORMATS = ["Upper", "Lower", "FirstCap", "Caps"] as const;
/** `\*` number formats (§17.16.4.3). */
export const NUMBER_FORMATS = [
  "Arabic",
  "ALPHABETIC",
  "alphabetic",
  "Roman",
  "roman",
  "Ordinal",
  "CardText",
  "OrdText",
  "Hex",
  "DollarText",
] as const;
/** `\#` numeric pictures Word's Field dialog offers (§17.16.4.2). */
export const NUMERIC_PICTURES = [
  "0",
  "#,##0",
  "#,##0.00",
  "$#,##0.00;($#,##0.00)",
  "0%",
  "0.00%",
] as const;
