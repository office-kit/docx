/**
 * Document protection (Review ▸ Protect Document / Restrict Editing / Always
 * Open Read-Only): `<w:documentProtection>`, `<w:writeProtection>` and
 * editable-range exceptions (`<w:permStart>` / `<w:permEnd>`).
 */

import { getPart } from "../internal/opc/index.js";
import {
  parseXml,
  serializeXml,
  type XmlAttr,
  type XmlElement,
  type XmlNode,
} from "../internal/xml/index.js";
import {
  DEFAULT_SPIN_COUNT,
  fromBase64,
  protectionHash,
  SHA512_ALGORITHM_SID,
  toBase64,
} from "../internal/wordprocessingml/protection-hash.js";
import {
  revisionElement,
  sortSettingsChildren,
  WML_NS,
  wmlAttrValue,
  type WmlParagraph,
} from "../internal/wordprocessingml/index.js";
import {
  type Docx,
  isolateParagraphRunRange,
  runTextLength,
  setDocumentSettingVal,
} from "./docx.js";

const SETTINGS_PART_NAME = "/word/settings.xml";
const SALT_BYTES = 16;

/** ST_DocProtect: what the reader may still change while protection is enforced. */
export type ProtectionEdit = "none" | "readOnly" | "comments" | "trackedChanges" | "forms";

export interface DocumentProtection {
  readonly edit: ProtectionEdit;
  /** Formatting is limited to the unlocked styles (`w:formatting`). */
  readonly formatting: boolean;
  /** The restriction is in force (`w:enforcement`); off, it is only remembered. */
  readonly enforced: boolean;
  /** A password is needed to stop protection. */
  readonly hasPassword: boolean;
}

export interface ProtectOptions {
  readonly edit: ProtectionEdit;
  readonly formatting?: boolean;
  /** Optional password; without one, anyone can stop protection. */
  readonly password?: string;
  /** Salt for the password hash; random when omitted (tests pass a fixed one). */
  readonly salt?: Uint8Array;
  /** Hash iterations; Word's 100 000 when omitted. */
  readonly spinCount?: number;
}

function attr(local: string, value: string): XmlAttr {
  return { name: { uri: WML_NS, local, prefix: "w" }, value, isNamespaceDecl: false };
}

function randomSalt(): Uint8Array {
  return crypto.getRandomValues(new Uint8Array(SALT_BYTES));
}

/**
 * The hash attributes Word 2013+ writes for a password (§17.15.1.29): the
 * legacy key, SHA-512 (`cryptAlgorithmSid` 14), iterated `spinCount` times.
 */
function passwordAttrs(password: string, salt: Uint8Array, spinCount: number): XmlAttr[] {
  return [
    attr("cryptProviderType", "rsaAES"),
    attr("cryptAlgorithmClass", "hash"),
    attr("cryptAlgorithmType", "typeAny"),
    attr("cryptAlgorithmSid", String(SHA512_ALGORITHM_SID)),
    attr("cryptSpinCount", String(spinCount)),
    attr("hash", toBase64(protectionHash(password, salt, spinCount))),
    attr("salt", toBase64(salt)),
  ];
}

function settingsXml(doc: Docx): XmlElement | undefined {
  const part = getPart(doc.opc, SETTINGS_PART_NAME);
  return part ? parseXml(new TextDecoder("utf-8").decode(part.data)).root : undefined;
}

function readSetting(doc: Docx, local: string): XmlElement | undefined {
  return settingsXml(doc)?.children.find(
    (c): c is XmlElement => c.kind === "element" && c.name.uri === WML_NS && c.name.local === local,
  );
}

/** Replace (or, with `undefined`, remove) one `<w:settings>` child, in schema order. */
function replaceSetting(doc: Docx, local: string, element: XmlElement | undefined): void {
  // Creates the settings part when the document has none.
  setDocumentSettingVal(doc, local, undefined);
  const part = getPart(doc.opc, SETTINGS_PART_NAME);
  if (!part) throw new Error("The document has no settings part.");
  const xml = parseXml(new TextDecoder("utf-8").decode(part.data));
  if (element) (xml.root.children as XmlNode[]).push(element);
  sortSettingsChildren(xml.root);
  part.data = new TextEncoder().encode(serializeXml(xml));
  doc.dirty = true;
}

const ON = new Set(["1", "true", "on"]);
const PROTECTION_EDITS: readonly ProtectionEdit[] = [
  "none",
  "readOnly",
  "comments",
  "trackedChanges",
  "forms",
];

/** The document's editing restriction (Restrict Editing), if any. */
export function documentProtection(doc: Docx): DocumentProtection | undefined {
  const el = readSetting(doc, "documentProtection");
  if (!el) return undefined;
  const value = wmlAttrValue(el, "edit");
  const edit = PROTECTION_EDITS.find((e) => e === value) ?? "none";
  return {
    edit,
    formatting: ON.has(wmlAttrValue(el, "formatting") ?? ""),
    enforced: ON.has(wmlAttrValue(el, "enforcement") ?? ""),
    hasPassword: wmlAttrValue(el, "hash") !== undefined,
  };
}

/**
 * Start enforcing an editing restriction (Restrict Editing ▸ Yes, Start
 * Enforcing Protection): read only, comments only, tracked changes only or
 * filling in forms, optionally with formatting limited to unlocked styles and
 * a password.
 */
export function protectDocument(doc: Docx, options: ProtectOptions): void {
  const attrs: XmlAttr[] = [attr("edit", options.edit)];
  if (options.formatting) attrs.push(attr("formatting", "1"));
  attrs.push(attr("enforcement", "1"));
  if (options.password) {
    attrs.push(
      ...passwordAttrs(
        options.password,
        options.salt ?? randomSalt(),
        options.spinCount ?? DEFAULT_SPIN_COUNT,
      ),
    );
  }
  replaceSetting(doc, "documentProtection", revisionElement("documentProtection", attrs));
}

function passwordMatches(el: XmlElement, password: string): boolean {
  const hash = wmlAttrValue(el, "hash");
  const salt = wmlAttrValue(el, "salt");
  if (hash === undefined || salt === undefined) return true;
  const spinCount = Number(wmlAttrValue(el, "cryptSpinCount") ?? 0);
  return toBase64(protectionHash(password, fromBase64(salt), spinCount)) === hash;
}

/**
 * Stop enforcing protection (Stop Protection). Returns false, changing
 * nothing, when the document has a protection password and `password` does
 * not match it. The restriction itself is kept, unenforced, as Word does.
 */
export function unprotectDocument(doc: Docx, password = ""): boolean {
  const el = readSetting(doc, "documentProtection");
  if (!el) return true;
  if (!passwordMatches(el, password)) return false;
  const keep = el.attrs.filter((a) => a.name.local === "edit" || a.name.local === "formatting");
  replaceSetting(
    doc,
    "documentProtection",
    revisionElement("documentProtection", [...keep, attr("enforcement", "0")]),
  );
  return true;
}

/** Whether `password` opens the document's editing restriction. */
export function verifyProtectionPassword(doc: Docx, password: string): boolean {
  const el = readSetting(doc, "documentProtection");
  return !el || passwordMatches(el, password);
}

export interface WriteProtection {
  /** Word suggests opening the file read-only (Always Open Read-Only). */
  readonly recommended: boolean;
  /** A password is needed to open the file for editing. */
  readonly hasPassword: boolean;
}

/** `<w:writeProtection>` (§17.15.1.93), if any. */
export function writeProtection(doc: Docx): WriteProtection | undefined {
  const el = readSetting(doc, "writeProtection");
  if (!el) return undefined;
  return {
    recommended: ON.has(wmlAttrValue(el, "recommended") ?? ""),
    hasPassword: wmlAttrValue(el, "hash") !== undefined,
  };
}

export interface WriteProtectionOptions {
  readonly recommended?: boolean;
  /** Password to modify. */
  readonly password?: string;
  readonly salt?: Uint8Array;
  readonly spinCount?: number;
}

/**
 * Set or clear write protection: Always Open Read-Only (`recommended`) and the
 * password to modify. `undefined` removes `<w:writeProtection>`.
 */
export function setWriteProtection(doc: Docx, options: WriteProtectionOptions | undefined): void {
  if (!options || (!options.recommended && !options.password)) {
    replaceSetting(doc, "writeProtection", undefined);
    return;
  }
  const attrs: XmlAttr[] = [];
  if (options.recommended) attrs.push(attr("recommended", "1"));
  if (options.password) {
    attrs.push(
      ...passwordAttrs(
        options.password,
        options.salt ?? randomSalt(),
        options.spinCount ?? DEFAULT_SPIN_COUNT,
      ),
    );
  }
  replaceSetting(doc, "writeProtection", revisionElement("writeProtection", attrs));
}

// --- exceptions: ranges editable while protected --------------------------------

/** Who may edit an exception range: a group (`w:edGrp`) or one person (`w:ed`). */
export type EditorGroup =
  | "everyone"
  | "current"
  | "editors"
  | "owners"
  | "contributors"
  | "administrators"
  | "none";

const EDITOR_GROUPS: readonly EditorGroup[] = [
  "everyone",
  "current",
  "editors",
  "owners",
  "contributors",
  "administrators",
  "none",
];

export interface EditableRange {
  readonly id: string;
  readonly group?: EditorGroup;
  readonly editor?: string;
  /** Top-level block where the range starts. */
  readonly block: number;
}

function childIndexAt(para: WmlParagraph, offset: number): number {
  isolateParagraphRunRange(para, 0, offset);
  let cursor = 0;
  let index = 0;
  for (const [i, child] of para.children.entries()) {
    if (child.kind !== "run") continue;
    const len = runTextLength(child);
    if (cursor + len > offset || (len > 0 && cursor >= offset)) break;
    cursor += len;
    index = i + 1;
  }
  return index;
}

function permElements(doc: Docx): Array<{ el: XmlElement; block: number; para: WmlParagraph }> {
  const out: Array<{ el: XmlElement; block: number; para: WmlParagraph }> = [];
  const visit = (para: WmlParagraph, block: number): void => {
    for (const child of para.children) {
      if (
        child.kind === "raw" &&
        child.node.name.uri === WML_NS &&
        (child.node.name.local === "permStart" || child.node.name.local === "permEnd")
      ) {
        out.push({ el: child.node, block, para });
      }
    }
  };
  for (const [block, node] of doc.document.body.blocks.entries()) {
    if (node.kind === "paragraph") visit(node, block);
    else if (node.kind === "table") {
      for (const row of node.rows) {
        for (const cell of row.cells) for (const p of cell.paragraphs) visit(p, block);
      }
    }
  }
  return out;
}

/** The exception ranges (`<w:permStart>`) in the body. */
export function editableRanges(doc: Docx): EditableRange[] {
  return permElements(doc)
    .filter(({ el }) => el.name.local === "permStart")
    .map(({ el, block }) => {
      const groupValue = wmlAttrValue(el, "edGrp");
      const group = EDITOR_GROUPS.find((g) => g === groupValue);
      const editor = wmlAttrValue(el, "ed");
      return {
        id: wmlAttrValue(el, "id") ?? "",
        block,
        ...(group === undefined ? {} : { group }),
        ...(editor === undefined ? {} : { editor }),
      };
    });
}

export interface EditableRangeOptions {
  /** Paragraph where the range ends (default: the start paragraph). */
  readonly endParagraph?: WmlParagraph;
  readonly group?: EditorGroup;
  /** A single editor (user name / e-mail) instead of a group. */
  readonly editor?: string;
}

/**
 * Mark characters `[start, end)` as an exception that stays editable while
 * the document is protected (Restrict Editing ▸ Exceptions). Returns the
 * range's `w:id`.
 */
export function addEditableRange(
  doc: Docx,
  para: WmlParagraph,
  start: number,
  end: number,
  options: EditableRangeOptions = { group: "everyone" },
): string {
  const used = new Set(permElements(doc).map(({ el }) => wmlAttrValue(el, "id")));
  let next = 0;
  while (used.has(String(next))) next++;
  const id = String(next);
  const who: XmlAttr[] = options.editor
    ? [attr("ed", options.editor)]
    : [attr("edGrp", options.group ?? "everyone")];
  const endPara = options.endParagraph ?? para;
  endPara.children.splice(childIndexAt(endPara, end), 0, {
    kind: "raw",
    node: revisionElement("permEnd", [attr("id", id)]),
  });
  para.children.splice(childIndexAt(para, start), 0, {
    kind: "raw",
    node: revisionElement("permStart", [attr("id", id), ...who]),
  });
  doc.dirty = true;
  return id;
}

/** Remove an exception range by `w:id`. Returns false when there is none. */
export function removeEditableRange(doc: Docx, id: string): boolean {
  const found = permElements(doc).filter(({ el }) => wmlAttrValue(el, "id") === id);
  for (const { el, para } of found) {
    para.children = para.children.filter((c) => !(c.kind === "raw" && c.node === el));
  }
  if (found.length > 0) doc.dirty = true;
  return found.length > 0;
}
