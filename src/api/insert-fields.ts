/**
 * Fields (ECMA-376 Part 1 §17.16): inserting complex fields at a caret and
 * recomputing their results.
 */

import {
  appProperties,
  coreProperties,
  type Docx,
  footnotesPart,
  stylesPart,
  text as documentText,
} from "./docx.js";
import {
  type DocumentAppProperties,
  type DocumentCoreProperties,
  findStyle,
  styleName,
  type WmlInline,
  type WmlParagraph,
  type WmlRun,
  type WmlRunPiece,
} from "../internal/wordprocessingml/index.js";
import {
  applyGeneralSwitches,
  evaluateComparison,
  evaluateFormula,
  formatDatePicture,
  type ParsedFieldInstruction,
  parseFieldInstruction,
} from "../internal/wordprocessingml/field-code.js";
import type { XmlElement } from "../internal/xml/index.js";
import {
  assertOffset,
  bodyParagraphs,
  isW,
  normalizeRawRuns,
  RPR_ORDER,
  runVisibleText,
  setOrderedChild,
  spliceInlinesAt,
  textPieces,
  wAttrValue,
  wChild,
  wEl,
  xmlVisibleText,
} from "./insert-xml.js";

export { formatDatePicture, quoteFieldArgument } from "../internal/wordprocessingml/field-code.js";

/** Values a field result depends on that are not stored in the document. */
export interface FieldContext {
  /** Clock for DATE / TIME (defaults to the current time). */
  readonly now?: Date;
  /** Page holding the field (PAGE); fields are left as they are when unknown. */
  readonly page?: number;
  /** Total pages (NUMPAGES, SECTIONPAGES); defaults to the app.xml page count. */
  readonly pageCount?: number;
  /** Page a paragraph is laid out on, for PAGEREF. */
  readonly pageOf?: (paragraph: WmlParagraph) => number | undefined;
  /** FILENAME (with `\p`, the caller's full path if it passes one). */
  readonly fileName?: string;
  /** FILESIZE in bytes. */
  readonly fileSize?: number;
  /** USERNAME / USERINITIALS / USERADDRESS (Word's user information). */
  readonly userName?: string;
  readonly userInitials?: string;
  readonly userAddress?: string;
}

export interface UpdateFieldsOptions extends FieldContext {
  /** Only update fields of these types (upper case, e.g. `["DATE", "PAGE"]`). */
  readonly types?: readonly string[];
}

export interface InsertFieldOptions {
  /**
   * Cached result to show. When omitted the result is computed (where the
   * field is computable) from the document and `context`.
   */
  readonly result?: string;
  /** Formatting (`<w:rPr>`) for the field's runs. */
  readonly rPr?: XmlElement;
  /**
   * Language of the field's runs (BCP 47, written as `<w:lang>`). Date
   * results use its month and day names, now and on every update.
   */
  readonly lang?: string;
  readonly context?: FieldContext;
}

// Default pictures when a date field has no `\@` switch (Word's en-US).
const DEFAULT_DATE_PICTURE = "M/d/yyyy";
const DEFAULT_TIME_PICTURE = "h:mm am/pm";
const DEFAULT_DATETIME_PICTURE = "M/d/yyyy h:mm:ss am/pm";
// Types whose results another module generates (tables, mail merge) or that
// have no visible result: updating them here would clobber that content.
const NOT_UPDATED: ReadonlySet<string> = new Set([
  "TOC",
  "TOA",
  "INDEX",
  "BIBLIOGRAPHY",
  "CITATION",
  "MERGEFIELD",
  "ADDRESSBLOCK",
  "GREETINGLINE",
  "MERGEREC",
  "MERGESEQ",
  "NEXT",
  "NEXTIF",
  "SKIPIF",
  "DATABASE",
  "HYPERLINK",
  "INCLUDETEXT",
  "INCLUDEPICTURE",
  "LINK",
  "EMBED",
  "FORMTEXT",
  "FORMCHECKBOX",
  "FORMDROPDOWN",
  "EQ",
  "BARCODE",
  "ASK",
  "FILLIN",
]);

function fldChar(type: "begin" | "separate" | "end", children: XmlElement[] = []): WmlRunPiece {
  return {
    kind: "fieldChar",
    charType: type,
    raw: wEl("fldChar", { fldCharType: type }, children),
  };
}

function runOf(pieces: WmlRunPiece[], rPr: XmlElement | undefined): WmlRun {
  return { kind: "run", ...(rPr ? { rPr: structuredClone(rPr) } : {}), pieces, extras: [] };
}

/**
 * The runs of a complex field: begin, instruction, separate, result, end
 * (§17.16.18). `beginChildren` go inside the begin `fldChar` (form fields put
 * their `<w:ffData>` there).
 */
export function buildComplexField(
  instruction: string,
  result: string,
  rPr?: XmlElement,
  beginChildren: XmlElement[] = [],
): WmlRun[] {
  return [
    runOf([fldChar("begin", beginChildren)], rPr),
    runOf([{ kind: "instrText", value: ` ${instruction.trim()} `, preserveSpace: true }], rPr),
    runOf([fldChar("separate")], rPr),
    runOf(textPieces(result), rPr),
    runOf([fldChar("end")], rPr),
  ];
}

/**
 * Insert a complex field at a character offset of `paragraph` (offsets count
 * visible characters, as `runTextLength` does). The cached result is computed
 * for the fields this library can evaluate (DATE, AUTHOR, NUMWORDS, SEQ, REF,
 * `=` formulas …) unless `options.result` is given; the runs inserted are
 * returned (begin … end), followed in the paragraph by nothing new.
 */
export function insertField(
  doc: Docx,
  paragraph: WmlParagraph,
  offset: number,
  instruction: string,
  options: InsertFieldOptions = {},
): WmlRun[] {
  assertOffset(paragraph, offset);
  if (instruction.trim() === "") throw new Error("A field needs an instruction.");
  const rPr = options.lang
    ? setOrderedChild(
        options.rPr ? structuredClone(options.rPr) : wEl("rPr"),
        "lang",
        wEl("lang", { val: options.lang }),
        RPR_ORDER,
      )
    : options.rPr;
  const runs = buildComplexField(instruction, options.result ?? "", rPr);
  spliceInlinesAt(paragraph, offset, runs);
  if (options.result === undefined) {
    updateFieldsWhere(doc, options.context ?? {}, (f) => f.begin === runs[0]);
  }
  doc.dirty = true;
  return runs;
}

/** Legacy form field kinds (§17.16.17 `w:ffData`). */
export type FormFieldOptions =
  | {
      readonly kind: "text";
      readonly name?: string;
      readonly defaultText?: string;
      readonly maxLength?: number;
      readonly helpText?: string;
    }
  | {
      readonly kind: "checkBox";
      readonly name?: string;
      readonly checked?: boolean;
      /** Box size in half-points; Word sizes it to the text when omitted. */
      readonly sizeHalfPoints?: number;
      readonly helpText?: string;
    }
  | {
      readonly kind: "dropDown";
      readonly name?: string;
      readonly entries: readonly string[];
      readonly selected?: number;
      readonly helpText?: string;
    };

// Word shows five en spaces in an empty text form field (U+2002).
const EMPTY_FORMTEXT = " ".repeat(5);
const BALLOT_BOX = "☐";
const BALLOT_BOX_CHECKED = "☒";
// §17.16.30 ffData/name: at most 20 characters.
const FORM_FIELD_NAME_MAX = 20;

/**
 * Insert a legacy form field (FORMTEXT / FORMCHECKBOX / FORMDROPDOWN) with
 * its `<w:ffData>` at a character offset.
 */
export function insertFormField(
  doc: Docx,
  paragraph: WmlParagraph,
  offset: number,
  options: FormFieldOptions,
): WmlRun[] {
  assertOffset(paragraph, offset);
  if (options.name !== undefined && options.name.length > FORM_FIELD_NAME_MAX) {
    throw new RangeError(`Form field names are at most ${FORM_FIELD_NAME_MAX} characters.`);
  }
  const ffChildren: XmlElement[] = [wEl("name", { val: options.name ?? "" }), wEl("enabled")];
  ffChildren.push(wEl("calcOnExit", { val: "0" }));
  if (options.helpText) ffChildren.push(wEl("helpText", { type: "text", val: options.helpText }));
  let instruction: string;
  let result: string;
  if (options.kind === "text") {
    const textInput: XmlElement[] = [];
    if (options.defaultText) textInput.push(wEl("default", { val: options.defaultText }));
    if (options.maxLength !== undefined) {
      textInput.push(wEl("maxLength", { val: String(options.maxLength) }));
    }
    ffChildren.push(wEl("textInput", {}, textInput));
    instruction = "FORMTEXT";
    result = options.defaultText || EMPTY_FORMTEXT;
  } else if (options.kind === "checkBox") {
    const size =
      options.sizeHalfPoints === undefined
        ? wEl("sizeAuto")
        : wEl("size", { val: String(options.sizeHalfPoints) });
    ffChildren.push(
      wEl("checkBox", {}, [size, wEl("default", { val: options.checked ? "1" : "0" })]),
    );
    instruction = "FORMCHECKBOX";
    result = options.checked ? BALLOT_BOX_CHECKED : BALLOT_BOX;
  } else {
    if (options.entries.length === 0) throw new Error("A drop-down form field needs entries.");
    const selected = options.selected ?? 0;
    ffChildren.push(
      wEl("ddList", {}, [
        wEl("default", { val: String(selected) }),
        ...options.entries.map((e) => wEl("listEntry", { val: e })),
      ]),
    );
    instruction = "FORMDROPDOWN";
    result = options.entries[selected] ?? options.entries[0] ?? "";
  }
  // Check boxes have no separate/result in Word's output, but a cached result
  // keeps the box visible in consumers that only read text.
  const runs = buildComplexField(instruction, result, undefined, [wEl("ffData", {}, ffChildren)]);
  spliceInlinesAt(paragraph, offset, runs);
  doc.dirty = true;
  return runs;
}

/**
 * Recompute the cached result of every computable field in the body, in
 * document order (so SEQ and AUTONUM count correctly). Fields whose result
 * cannot be computed from the document and `options` are left untouched, as
 * are fields generated elsewhere (TOC, mail merge) and locked (`\!`) ones.
 * Returns how many results changed.
 */
export function updateFields(doc: Docx, options: UpdateFieldsOptions = {}): number {
  const types = options.types ? new Set(options.types.map((t) => t.toUpperCase())) : undefined;
  return updateFieldsWhere(doc, options, (f) => !types || types.has(f.parsed.type));
}

export interface FieldInfo {
  readonly paragraph: WmlParagraph;
  readonly instruction: string;
  readonly type: string;
  /** The run holding the begin `fldChar`. */
  readonly begin: WmlRun;
  /** Current cached result text. */
  readonly result: string;
}

/** Every complex field in the body, in document order (outer fields first). */
export function complexFields(doc: Docx): FieldInfo[] {
  const out: FieldInfo[] = [];
  walkFields(doc, (f) => {
    out.push({
      paragraph: f.paragraph,
      instruction: f.instruction,
      type: f.parsed.type,
      begin: f.begin,
      result: f.resultText,
    });
    return undefined;
  });
  // The walk reports inner fields first; list them by where they begin.
  const paraIndex = new Map(bodyParagraphs(doc).map((p, i) => [p, i] as const));
  const keyed = out.map((f) => ({
    f,
    para: paraIndex.get(f.paragraph) ?? 0,
    at: f.paragraph.children.indexOf(f.begin),
  }));
  keyed.sort((a, b) => a.para - b.para || a.at - b.at);
  return keyed.map((x) => x.f);
}

// --- field walking ---------------------------------------------------------------

interface WalkedField {
  readonly paragraph: WmlParagraph;
  readonly begin: WmlRun;
  readonly separate: WmlRun | undefined;
  readonly end: WmlRun;
  readonly instruction: string;
  readonly parsed: ParsedFieldInstruction;
  readonly resultText: string;
  /** All of begin/separate/end are in `paragraph`. */
  readonly local: boolean;
}

interface OpenField {
  readonly paragraph: WmlParagraph;
  readonly begin: WmlRun;
  separate: WmlRun | undefined;
  instruction: string;
  result: string;
  local: boolean;
}

/**
 * Walk the body's complex fields innermost-first in document order. `visit`
 * returns the new result for a field (or undefined to keep it); a nested
 * field's (new) result becomes part of its parent's instruction, as Word
 * evaluates them.
 */
function walkFields(doc: Docx, visit: (f: WalkedField) => string | undefined): number {
  let changed = 0;
  const stack: OpenField[] = [];
  for (const para of bodyParagraphs(doc)) {
    normalizeRawRuns(para);
    for (let i = 0; i < para.children.length; i++) {
      const child = para.children[i];
      if (!child) continue;
      if (child.kind !== "run") {
        const text = xmlVisibleText(child.node);
        for (const open of stack) if (open.separate) open.result += text;
        continue;
      }
      for (const piece of child.pieces) {
        const top = stack[stack.length - 1];
        if (piece.kind === "fieldChar") {
          if (piece.charType === "begin") {
            stack.push({
              paragraph: para,
              begin: child,
              separate: undefined,
              instruction: "",
              result: "",
              local: true,
            });
          } else if (piece.charType === "separate" && top) {
            top.separate = child;
            if (top.paragraph !== para) top.local = false;
          } else if (piece.charType === "end" && top) {
            stack.pop();
            const instruction = top.instruction;
            const field: WalkedField = {
              paragraph: top.paragraph,
              begin: top.begin,
              separate: top.separate,
              end: child,
              instruction,
              parsed: parseFieldInstruction(instruction),
              resultText: top.result,
              local: top.local && top.paragraph === para,
            };
            const next = visit(field);
            let shown = top.result;
            if (next !== undefined && field.local && next !== top.result) {
              replaceResult(field, next);
              changed++;
              shown = next;
              i = para.children.indexOf(child);
            }
            const parent = stack[stack.length - 1];
            if (parent) {
              if (parent.separate) parent.result += shown;
              else parent.instruction += shown;
            }
          }
          continue;
        }
        if (!top) continue;
        if (piece.kind === "instrText" && !top.separate) top.instruction += piece.value;
        else if (top.separate) {
          for (const open of stack) if (open.separate) open.result += pieceText(piece);
        }
      }
    }
  }
  return changed;
}

function pieceText(piece: WmlRunPiece): string {
  if (piece.kind === "text") return piece.value;
  if (piece.kind === "tab") return "\t";
  if (piece.kind === "break") return "\n";
  return "";
}

/** Swap the runs between `separate` and `end` for one run holding `result`. */
function replaceResult(field: WalkedField, result: string): void {
  const children = field.paragraph.children;
  const endIndex = children.indexOf(field.end);
  let startIndex: number;
  if (field.separate) {
    startIndex = children.indexOf(field.separate) + 1;
  } else {
    children.splice(endIndex, 0, runOf([fldChar("separate")], field.begin.rPr));
    startIndex = endIndex + 1;
  }
  const end = children.indexOf(field.end);
  const old = children.slice(startIndex, end);
  // \* MERGEFORMAT keeps the old result's look; otherwise the result takes the
  // field code's formatting (§17.16.4.3).
  const mergeFormat = field.parsed.switches.some(
    (s) => s.name === "*" && s.arg?.toUpperCase() === "MERGEFORMAT",
  );
  const firstOld = old.find((c): c is WmlRun => c.kind === "run");
  const rPr = mergeFormat && firstOld ? firstOld.rPr : field.begin.rPr;
  const replacement: WmlInline = runOf(textPieces(result), rPr);
  children.splice(startIndex, end - startIndex, replacement);
}

// --- evaluation --------------------------------------------------------------------

interface BookmarkInfo {
  text: string;
  readonly paragraph: WmlParagraph;
  readonly order: number;
  noteNumber: number | undefined;
}

interface EvalState {
  readonly ctx: FieldContext;
  readonly doc: Docx;
  readonly seq: Map<string, number>;
  readonly seqEpoch: Map<string, number>;
  autonum: number;
  /** Paragraph order of the field being evaluated (for REF \p). */
  readonly paragraphOrder: Map<WmlParagraph, number>;
  /** Per paragraph: how many headings of each level (or higher) precede it. */
  readonly paragraphHeadingEpochs: Map<WmlParagraph, number[]>;
  bookmarks?: Map<string, BookmarkInfo>;
}

const MAX_HEADING_LEVEL = 9;
const HEADING_STYLE = /^heading\s?(\d)$/i;

function headingLevel(doc: Docx, para: WmlParagraph): number | undefined {
  const id = para.pPr ? wAttrValue(wChild(para.pPr, "pStyle") ?? wEl("x"), "val") : undefined;
  if (!id) return undefined;
  const m = HEADING_STYLE.exec(id);
  if (m?.[1]) return Number(m[1]);
  const part = stylesPart(doc);
  const style = part ? findStyle(part, id) : undefined;
  const name = style ? styleName(style) : undefined;
  const byName = name ? HEADING_STYLE.exec(name) : null;
  return byName?.[1] ? Number(byName[1]) : undefined;
}

function updateFieldsWhere(
  doc: Docx,
  ctx: FieldContext,
  wanted: (f: WalkedField) => boolean,
): number {
  const paragraphOrder = new Map<WmlParagraph, number>();
  const paragraphHeadingEpochs = new Map<WmlParagraph, number[]>();
  const headingEpochs = Array.from({ length: MAX_HEADING_LEVEL + 1 }, () => 0);
  for (const [i, p] of bodyParagraphs(doc).entries()) {
    paragraphOrder.set(p, i);
    const level = headingLevel(doc, p);
    if (level !== undefined) {
      for (let l = level; l <= MAX_HEADING_LEVEL; l++)
        headingEpochs[l] = (headingEpochs[l] ?? 0) + 1;
    }
    paragraphHeadingEpochs.set(p, [...headingEpochs]);
  }
  const state: EvalState = {
    ctx,
    doc,
    seq: new Map(),
    seqEpoch: new Map(),
    autonum: 0,
    paragraphOrder,
    paragraphHeadingEpochs,
  };
  const changed = walkFields(doc, (f) => {
    // Counters advance for every field, so a single updated SEQ still counts
    // the ones before it.
    const value = evaluate(f, state);
    if (!wanted(f) || NOT_UPDATED.has(f.parsed.type)) return undefined;
    if (f.parsed.switches.some((s) => s.name === "!")) return undefined;
    return value;
  });
  if (changed) doc.dirty = true;
  return changed;
}

function dateFromIso(value: string | undefined): Date | undefined {
  if (!value) return undefined;
  const d = new Date(value);
  return Number.isNaN(d.getTime()) ? undefined : d;
}

/** The begin run's `w:lang` (Latin text language), if any. */
function fieldLanguage(f: WalkedField): string | undefined {
  const lang = wChild(f.begin.rPr, "lang");
  return lang ? wAttrValue(lang, "val") : undefined;
}

function datePicture(parsed: ParsedFieldInstruction, fallback: string): string {
  return parsed.switches.find((s) => s.name === "@")?.arg ?? fallback;
}

function hasSwitch(parsed: ParsedFieldInstruction, name: string): boolean {
  return parsed.switches.some((s) => s.name === name);
}

function switchArg(parsed: ParsedFieldInstruction, name: string): string | undefined {
  return parsed.switches.find((s) => s.name === name)?.arg;
}

// DOCPROPERTY names (Word's built-in property names) -> core.xml / app.xml keys.
const CORE_PROPERTY_NAMES: Readonly<Record<string, keyof DocumentCoreProperties>> = {
  TITLE: "title",
  SUBJECT: "subject",
  AUTHOR: "creator",
  KEYWORDS: "keywords",
  COMMENTS: "description",
  CATEGORY: "category",
  LASTSAVEDBY: "lastModifiedBy",
  REVISIONNUMBER: "revision",
  CREATETIME: "created",
  LASTSAVEDTIME: "modified",
  CONTENTSTATUS: "contentStatus",
};
const APP_PROPERTY_NAMES: Readonly<Record<string, keyof DocumentAppProperties>> = {
  COMPANY: "company",
  MANAGER: "manager",
  TEMPLATE: "template",
  PAGES: "pages",
  WORDS: "words",
  CHARACTERS: "characters",
  CHARACTERSWITHSPACES: "charactersWithSpaces",
  LINES: "lines",
  PARAGRAPHS: "paragraphs",
};
const DATE_PROPERTIES: ReadonlySet<string> = new Set(["created", "modified"]);

function documentProperty(doc: Docx, name: string): string | undefined {
  const key = name.toUpperCase().replace(/\s/g, "");
  const coreKey = CORE_PROPERTY_NAMES[key];
  if (coreKey) {
    const value = coreProperties(doc)[coreKey];
    if (!DATE_PROPERTIES.has(coreKey)) return value ?? "";
    const d = dateFromIso(value);
    return d ? formatDatePicture(d, DEFAULT_DATETIME_PICTURE) : undefined;
  }
  const appKey = APP_PROPERTY_NAMES[key];
  if (!appKey) return undefined;
  const value = appProperties(doc)[appKey];
  return value === undefined ? "" : String(value);
}

function wordCount(doc: Docx): number {
  return documentText(doc).trim().match(/\S+/g)?.length ?? 0;
}

function charCount(doc: Docx): number {
  return documentText(doc).replace(/\s/g, "").length;
}

function evaluate(f: WalkedField, state: EvalState): string | undefined {
  const raw = evaluateRaw(f, state);
  if (raw === undefined) return undefined;
  const picture = f.parsed.switches.some((s) => s.name === "@");
  // Date fields already applied their \@ picture; \* and \# apply to all.
  return picture
    ? applyGeneralSwitches(
        raw,
        f.parsed.switches.filter((s) => s.name !== "@"),
      )
    : applyGeneralSwitches(raw, f.parsed.switches);
}

function evaluateRaw(f: WalkedField, state: EvalState): string | undefined {
  const { parsed } = f;
  const { ctx, doc } = state;
  const arg0 = parsed.args[0];
  switch (parsed.type) {
    case "DATE":
      return formatDatePicture(
        ctx.now ?? new Date(),
        datePicture(parsed, DEFAULT_DATE_PICTURE),
        fieldLanguage(f),
      );
    case "TIME":
      return formatDatePicture(
        ctx.now ?? new Date(),
        datePicture(parsed, DEFAULT_TIME_PICTURE),
        fieldLanguage(f),
      );
    case "CREATEDATE":
    case "SAVEDATE": {
      const core = coreProperties(doc);
      const d = dateFromIso(parsed.type === "CREATEDATE" ? core.created : core.modified);
      return d
        ? formatDatePicture(d, datePicture(parsed, DEFAULT_DATETIME_PICTURE), fieldLanguage(f))
        : undefined;
    }
    case "AUTHOR":
      return arg0 ?? coreProperties(doc).creator ?? "";
    case "TITLE":
      return arg0 ?? coreProperties(doc).title ?? "";
    case "SUBJECT":
      return arg0 ?? coreProperties(doc).subject ?? "";
    case "KEYWORDS":
      return arg0 ?? coreProperties(doc).keywords ?? "";
    case "COMMENTS":
      return arg0 ?? coreProperties(doc).description ?? "";
    case "LASTSAVEDBY":
      return coreProperties(doc).lastModifiedBy ?? "";
    case "REVNUM":
      return coreProperties(doc).revision ?? "";
    case "TEMPLATE":
      return appProperties(doc).template ?? "";
    case "DOCPROPERTY":
      return arg0 ? documentProperty(doc, arg0) : undefined;
    case "INFO":
      return arg0 ? (parsed.args[1] ?? documentProperty(doc, arg0)) : undefined;
    case "NUMWORDS":
      return String(wordCount(doc));
    case "NUMCHARS":
      return String(charCount(doc));
    case "NUMPAGES":
    case "SECTIONPAGES": {
      const n = ctx.pageCount ?? appProperties(doc).pages;
      return n === undefined ? undefined : String(n);
    }
    case "PAGE":
      return ctx.page === undefined ? undefined : String(ctx.page);
    case "SECTION":
      return "1";
    case "FILENAME":
      if (ctx.fileName === undefined) return undefined;
      return hasSwitch(parsed, "p") ? ctx.fileName : (ctx.fileName.split(/[\\/]/).pop() ?? "");
    case "FILESIZE":
      return ctx.fileSize === undefined ? undefined : String(ctx.fileSize);
    case "USERNAME":
      return arg0 ?? ctx.userName;
    case "USERINITIALS":
      return arg0 ?? ctx.userInitials;
    case "USERADDRESS":
      return arg0 ?? ctx.userAddress;
    case "QUOTE":
      return parsed.args.join(" ");
    case "SYMBOL":
      return symbolFieldText(parsed);
    case "=":
      return formulaText(parsed, state);
    case "IF":
      return ifText(parsed);
    case "COMPARE":
      return parsed.args.length >= 3
        ? evaluateComparison(parsed.args[0] ?? "", parsed.args[1] ?? "", parsed.args[2] ?? "")
          ? "1"
          : "0"
        : "0";
    case "SEQ":
      return seqText(f, state);
    case "AUTONUM":
    case "AUTONUMLGL":
    case "AUTONUMOUT":
      state.autonum++;
      return `${state.autonum}.`;
    case "LISTNUM":
      state.autonum++;
      return `${state.autonum})`;
    case "REF":
      return refText(f, state, arg0);
    case "PAGEREF":
      return pageRefText(f, state, arg0);
    case "NOTEREF": {
      const bm = arg0 ? bookmarkIndex(state).get(arg0) : undefined;
      if (!bm) return undefined;
      if (hasSwitch(parsed, "p")) return aboveBelow(f, state, bm);
      return bm.noteNumber === undefined ? undefined : String(bm.noteNumber);
    }
    case "STYLEREF":
      return arg0 ? styleRefText(f, state, arg0) : undefined;
    case "MACROBUTTON":
      return parsed.args.slice(1).join(" ");
    case "GOTOBUTTON":
      return parsed.args.slice(1).join(" ");
    case "ADVANCE":
    case "PRINT":
    case "SET":
    case "PRIVATE":
    case "RD":
    case "TA":
    case "TC":
    case "XE":
      return "";
    default:
      return undefined;
  }
}

function symbolFieldText(parsed: ParsedFieldInstruction): string | undefined {
  const code = parsed.args[0];
  if (!code) return undefined;
  const n = /^0x/i.test(code) ? Number.parseInt(code.slice(2), 16) : Number(code);
  return Number.isInteger(n) && n > 0 ? String.fromCodePoint(n) : undefined;
}

function formulaText(parsed: ParsedFieldInstruction, state: EvalState): string {
  const expr = parsed.args[0] ?? "";
  try {
    const value = evaluateFormula(expr, (name) => {
      const bm = bookmarkIndex(state).get(name) ?? findBookmarkIgnoringCase(state, name);
      if (!bm) return undefined;
      const n = Number(bm.text.replace(/[^\d.-]/g, ""));
      return Number.isFinite(n) ? n : undefined;
    });
    if (parsed.switches.some((s) => s.name === "#")) return String(value);
    return Number.isInteger(value) ? String(value) : String(Math.round(value * 100) / 100);
  } catch (err) {
    // Word shows the error message as the field's result.
    if (err instanceof Error && err.message.startsWith("!")) return err.message;
    return "!Syntax Error";
  }
}

function findBookmarkIgnoringCase(state: EvalState, name: string): BookmarkInfo | undefined {
  for (const [k, v] of bookmarkIndex(state)) if (k.toUpperCase() === name) return v;
  return undefined;
}

const COMPARISON_OPERATORS = new Set(["=", "<>", "<", ">", "<=", ">="]);

function ifText(parsed: ParsedFieldInstruction): string | undefined {
  const [left, op, right, whenTrue = "", whenFalse = ""] = parsed.args;
  if (left === undefined || op === undefined || right === undefined) return undefined;
  if (!COMPARISON_OPERATORS.has(op)) return undefined;
  return evaluateComparison(left, op, right) ? whenTrue : whenFalse;
}

function seqText(f: WalkedField, state: EvalState): string | undefined {
  const id = f.parsed.args[0];
  if (!id) return undefined;
  const key = id.toUpperCase();
  const resetLevel = switchArg(f.parsed, "s");
  if (resetLevel !== undefined) {
    const level = Number(resetLevel);
    const epoch = state.paragraphHeadingEpochs.get(f.paragraph)?.[level] ?? 0;
    if (state.seqEpoch.get(key) !== epoch) state.seq.set(key, 0);
    state.seqEpoch.set(key, epoch);
  }
  const reset = switchArg(f.parsed, "r");
  let n = state.seq.get(key) ?? 0;
  if (reset !== undefined && Number.isFinite(Number(reset))) n = Number(reset);
  else if (!hasSwitch(f.parsed, "c")) n += 1;
  state.seq.set(key, n);
  if (hasSwitch(f.parsed, "h")) return "";
  return String(n);
}

function bookmarkIndex(state: EvalState): Map<string, BookmarkInfo> {
  if (state.bookmarks) return state.bookmarks;
  const map = new Map<string, BookmarkInfo>();
  const open = new Map<string, BookmarkInfo>();
  const noteNumbers = footnoteNumbers(state.doc);
  for (const [order, para] of bodyParagraphs(state.doc).entries()) {
    for (const child of para.children) {
      if (child.kind === "raw" && isW(child.node, "bookmarkStart")) {
        const name = wAttrValue(child.node, "name");
        const id = wAttrValue(child.node, "id");
        if (!name || id === undefined) continue;
        const info: BookmarkInfo = { text: "", paragraph: para, order, noteNumber: undefined };
        map.set(name, info);
        open.set(id, info);
        continue;
      }
      if (child.kind === "raw" && isW(child.node, "bookmarkEnd")) {
        const id = wAttrValue(child.node, "id");
        if (id !== undefined) open.delete(id);
        continue;
      }
      const text = child.kind === "run" ? runVisibleText(child) : xmlVisibleText(child.node);
      const note = child.kind === "run" ? noteNumberOf(child, noteNumbers) : undefined;
      for (const info of open.values()) {
        info.text += text;
        if (note !== undefined && info.noteNumber === undefined) info.noteNumber = note;
      }
    }
  }
  state.bookmarks = map;
  return map;
}

/** Footnote id → its displayed number (document order of references). */
function footnoteNumbers(doc: Docx): Map<string, number> {
  const out = new Map<string, number>();
  let n = 0;
  for (const para of bodyParagraphs(doc)) {
    for (const child of para.children) {
      if (child.kind !== "run") continue;
      for (const p of child.pieces) {
        if (p.kind === "raw" && isW(p.node, "footnoteReference")) {
          const id = wAttrValue(p.node, "id");
          if (id !== undefined && !out.has(id)) out.set(id, ++n);
        }
      }
    }
  }
  // Touch the part so a document without footnotes.xml still resolves to none.
  footnotesPart(doc);
  return out;
}

function noteNumberOf(run: WmlRun, numbers: Map<string, number>): number | undefined {
  for (const p of run.pieces) {
    if (p.kind === "raw" && isW(p.node, "footnoteReference")) {
      const id = wAttrValue(p.node, "id");
      if (id !== undefined) return numbers.get(id);
    }
  }
  return undefined;
}

function aboveBelow(f: WalkedField, state: EvalState, bm: BookmarkInfo): string {
  const here = state.paragraphOrder.get(f.paragraph) ?? 0;
  return bm.order < here ? "above" : "below";
}

function refText(f: WalkedField, state: EvalState, name: string | undefined): string | undefined {
  if (!name) return undefined;
  const bm = bookmarkIndex(state).get(name);
  if (!bm) return "Error! Reference source not found.";
  if (hasSwitch(f.parsed, "p")) return aboveBelow(f, state, bm);
  // \n \r \w (paragraph numbers) need list numbering, which is computed at
  // layout time; keep Word's cached number.
  if (hasSwitch(f.parsed, "n") || hasSwitch(f.parsed, "r") || hasSwitch(f.parsed, "w")) {
    return undefined;
  }
  return bm.text.replace(/\s+$/, "");
}

function pageRefText(
  f: WalkedField,
  state: EvalState,
  name: string | undefined,
): string | undefined {
  if (!name) return undefined;
  const bm = bookmarkIndex(state).get(name);
  if (!bm) return "Error! Bookmark not defined.";
  const page = state.ctx.pageOf?.(bm.paragraph) ?? (state.ctx.pageCount === 1 ? 1 : undefined);
  if (hasSwitch(f.parsed, "p")) {
    const here = state.ctx.pageOf?.(f.paragraph);
    if (page !== undefined && here !== undefined && page === here) return aboveBelow(f, state, bm);
    return page === undefined ? undefined : `on page ${page}`;
  }
  return page === undefined ? undefined : String(page);
}

function styleRefText(f: WalkedField, state: EvalState, styleRef: string): string | undefined {
  const paras = bodyParagraphs(state.doc);
  const here = state.paragraphOrder.get(f.paragraph) ?? 0;
  const part = stylesPart(state.doc);
  const wanted = styleRef.toLowerCase();
  const matches = (p: WmlParagraph): boolean => {
    const id = p.pPr ? wAttrValue(wChild(p.pPr, "pStyle") ?? wEl("x"), "val") : undefined;
    if (!id) return wanted === "normal";
    if (id.toLowerCase() === wanted) return true;
    const style = part ? findStyle(part, id) : undefined;
    return (style ? styleName(style) : undefined)?.toLowerCase() === wanted;
  };
  // Nearest paragraph with the style at or above the field, else below it.
  for (let i = here; i >= 0; i--) {
    const p = paras[i];
    if (p && matches(p)) return paragraphPlainText(p);
  }
  for (let i = here + 1; i < paras.length; i++) {
    const p = paras[i];
    if (p && matches(p)) return paragraphPlainText(p);
  }
  return "Error! No text of specified style in document.";
}

function paragraphPlainText(p: WmlParagraph): string {
  let out = "";
  for (const c of p.children) out += c.kind === "run" ? runVisibleText(c) : xmlVisibleText(c.node);
  return out;
}
