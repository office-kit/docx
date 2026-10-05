/**
 * View: the view and zoom a document opens in (`<w:view>`, `<w:zoom>` in
 * `settings.xml`), and custom document properties (`docProps/custom.xml`,
 * File ▸ Properties ▸ Custom).
 */

import {
  addPart,
  addRelationship,
  getPart,
  packageRelationships,
  relationshipsByType,
} from "../internal/opc/index.js";
import {
  parseXml,
  serializeXml,
  type XmlAttr,
  type XmlElement,
  type XmlNode,
} from "../internal/xml/index.js";
import {
  revisionElement,
  sortSettingsChildren,
  WML_NS,
  wmlAttrValue,
} from "../internal/wordprocessingml/index.js";
import { type Docx, getDocumentSetting, setDocumentSettingVal } from "./docx.js";

// --- view and zoom --------------------------------------------------------------

/**
 * ST_View (§17.18.102) as Word names the views: Print Layout is the default
 * (no `<w:view>`), Draft is `normal`, Web Layout `web`, Outline `outline`.
 */
export type DocumentView = "print" | "web" | "outline" | "draft";

const VIEWS: readonly DocumentView[] = ["print", "web", "outline", "draft"];
const VIEW_TO_VAL: Readonly<Record<DocumentView, string | undefined>> = {
  print: undefined,
  web: "web",
  outline: "outline",
  draft: "normal",
};

/** The view the document asks to be opened in. */
export function documentView(doc: Docx): DocumentView {
  const val = getDocumentSetting(doc, "view").val;
  return VIEWS.find((view) => val !== undefined && VIEW_TO_VAL[view] === val) ?? "print";
}

/** Save the view the document opens in. */
export function setDocumentView(doc: Docx, view: DocumentView): void {
  setDocumentSettingVal(doc, "view", VIEW_TO_VAL[view]);
}

/** ST_Zoom presets (§17.18.105): Page Width, Text Width, Whole Page. */
export type ZoomPreset = "none" | "bestFit" | "fullPage" | "textFit";

export interface DocumentZoom {
  /** Zoom percentage (`w:percent`). */
  readonly percent: number;
  readonly preset: ZoomPreset;
}

const ZOOM_PRESETS: readonly ZoomPreset[] = ["none", "bestFit", "fullPage", "textFit"];
// Word's range for the zoom percentage.
const MIN_ZOOM_PERCENT = 10;
const MAX_ZOOM_PERCENT = 500;
const DEFAULT_ZOOM_PERCENT = 100;

const SETTINGS_PART_NAME = "/word/settings.xml";

function settingsRoot(doc: Docx): XmlElement | undefined {
  const part = getPart(doc.opc, SETTINGS_PART_NAME);
  return part ? parseXml(new TextDecoder("utf-8").decode(part.data)).root : undefined;
}

/** The zoom the document opens at (100 % when it records none). */
export function documentZoom(doc: Docx): DocumentZoom {
  const zoom = settingsRoot(doc)?.children.find(
    (c): c is XmlElement =>
      c.kind === "element" && c.name.uri === WML_NS && c.name.local === "zoom",
  );
  const percent = Number(zoom ? (wmlAttrValue(zoom, "percent") ?? "") : "");
  const val = zoom && wmlAttrValue(zoom, "val");
  return {
    percent: Number.isFinite(percent) && percent > 0 ? percent : DEFAULT_ZOOM_PERCENT,
    preset: ZOOM_PRESETS.find((p) => p === val) ?? "none",
  };
}

/**
 * Save the zoom the document opens at. Throws a `RangeError` for a percentage
 * outside Word's 10–500 range.
 */
export function setDocumentZoom(doc: Docx, zoom: DocumentZoom): void {
  if (
    !Number.isInteger(zoom.percent) ||
    zoom.percent < MIN_ZOOM_PERCENT ||
    zoom.percent > MAX_ZOOM_PERCENT
  ) {
    throw new RangeError(
      `Zoom must be a whole percentage from ${MIN_ZOOM_PERCENT} to ${MAX_ZOOM_PERCENT}, got ${zoom.percent}.`,
    );
  }
  // Creates the settings part (and drops any old <w:zoom>) first.
  setDocumentSettingVal(doc, "zoom", undefined);
  const part = getPart(doc.opc, SETTINGS_PART_NAME);
  if (!part) throw new Error("The document has no settings part.");
  const xml = parseXml(new TextDecoder("utf-8").decode(part.data));
  const attrs: XmlAttr[] = [];
  const wAttr = (local: string, value: string): XmlAttr => ({
    name: { uri: WML_NS, local, prefix: "w" },
    value,
    isNamespaceDecl: false,
  });
  if (zoom.preset !== "none") attrs.push(wAttr("val", zoom.preset));
  attrs.push(wAttr("percent", String(zoom.percent)));
  (xml.root.children as XmlNode[]).push(revisionElement("zoom", attrs));
  sortSettingsChildren(xml.root);
  part.data = new TextEncoder().encode(serializeXml(xml));
  doc.dirty = true;
}

// --- custom properties ----------------------------------------------------------

const CUSTOM_PART_NAME = "/docProps/custom.xml";
const CUSTOM_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.custom-properties+xml";
const CUSTOM_REL_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties";
const CUSTOM_NS = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";
const VT_NS = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";
// The format id every custom property carries (Part 1 §22.3.2.2).
const CUSTOM_FMTID = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";
// Property ids 0 and 1 are reserved; custom properties start at 2.
const FIRST_PID = 2;
const INT32_MIN = -2_147_483_648;
const INT32_MAX = 2_147_483_647;

/** A custom property value: Word's Text, Number, Date and Yes or No types. */
export type CustomPropertyValue = string | number | boolean | Date;

export interface CustomProperty {
  readonly name: string;
  readonly value: CustomPropertyValue;
}

function vtElement(local: string, text: string): XmlElement {
  return {
    kind: "element",
    name: { uri: VT_NS, local, prefix: "vt" },
    attrs: [],
    children: [{ kind: "text", value: text }],
    xmlSpace: "default",
    selfClosing: false,
  };
}

function valueElement(value: CustomPropertyValue): XmlElement {
  if (typeof value === "boolean") return vtElement("bool", value ? "true" : "false");
  if (value instanceof Date) {
    if (Number.isNaN(value.getTime())) throw new RangeError("Invalid date.");
    return vtElement("filetime", value.toISOString().replace(/\.\d{3}Z$/, "Z"));
  }
  if (typeof value === "number") {
    if (!Number.isFinite(value)) throw new RangeError(`Not a finite number: ${value}.`);
    return Number.isInteger(value) && value >= INT32_MIN && value <= INT32_MAX
      ? vtElement("i4", String(value))
      : vtElement("r8", String(value));
  }
  return vtElement("lpwstr", value);
}

function textOf(el: XmlElement): string {
  return el.children.map((c) => (c.kind === "text" || c.kind === "cdata" ? c.value : "")).join("");
}

const NUMERIC_TYPES: ReadonlySet<string> = new Set([
  "i1",
  "i2",
  "i4",
  "i8",
  "int",
  "ui1",
  "ui2",
  "ui4",
  "ui8",
  "uint",
  "r4",
  "r8",
  "decimal",
]);

function readValue(el: XmlElement): CustomPropertyValue {
  const text = textOf(el);
  if (el.name.local === "bool") return text === "true" || text === "1";
  if (el.name.local === "filetime" || el.name.local === "date") return new Date(text);
  if (NUMERIC_TYPES.has(el.name.local)) return Number(text);
  return text;
}

function customRoot(doc: Docx): XmlElement | undefined {
  const part = getPart(doc.opc, CUSTOM_PART_NAME);
  return part ? parseXml(new TextDecoder("utf-8").decode(part.data)).root : undefined;
}

function propertyElements(children: readonly XmlNode[]): XmlElement[] {
  return children.filter(
    (c): c is XmlElement =>
      c.kind === "element" && c.name.uri === CUSTOM_NS && c.name.local === "property",
  );
}

/** An unqualified attribute (the custom-properties schema uses no namespace on attributes). */
function plainAttr(local: string, value: string): XmlAttr {
  return { name: { uri: "", local, prefix: "" }, value, isNamespaceDecl: false };
}

function attrValue(el: XmlElement, local: string): string | undefined {
  return el.attrs.find((a) => a.name.uri === "" && a.name.local === local)?.value;
}

/** The document's custom properties (File ▸ Properties ▸ Custom). */
export function customProperties(doc: Docx): CustomProperty[] {
  return propertyElements(customRoot(doc)?.children ?? []).flatMap((prop) => {
    const name = attrValue(prop, "name");
    const value = prop.children.find((c): c is XmlElement => c.kind === "element");
    return name === undefined || !value ? [] : [{ name, value: readValue(value) }];
  });
}

const EMPTY_CUSTOM_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  `<Properties xmlns="${CUSTOM_NS}" xmlns:vt="${VT_NS}"/>`;

function writeCustom(doc: Docx, edit: (properties: XmlNode[]) => void): void {
  if (!getPart(doc.opc, CUSTOM_PART_NAME)) {
    addPart(doc.opc, {
      name: CUSTOM_PART_NAME,
      contentType: CUSTOM_CONTENT_TYPE,
      data: new TextEncoder().encode(EMPTY_CUSTOM_XML),
    });
    const rels = packageRelationships(doc.opc);
    if (relationshipsByType(rels, CUSTOM_REL_TYPE).length === 0) {
      addRelationship(rels, { type: CUSTOM_REL_TYPE, target: "docProps/custom.xml" });
    }
  }
  const part = getPart(doc.opc, CUSTOM_PART_NAME);
  if (!part) throw new Error("Could not create docProps/custom.xml.");
  const xml = parseXml(new TextDecoder("utf-8").decode(part.data));
  edit(xml.root.children as XmlNode[]);
  part.data = new TextEncoder().encode(serializeXml(xml));
}

/**
 * Add or change a custom property. A string is Word's Text type, a number
 * Number (`vt:i4`, or `vt:r8` when not a 32-bit integer), a boolean Yes or
 * No, a `Date` Date (`vt:filetime`). Throws a `RangeError` for an empty name,
 * a non-finite number or an invalid date.
 */
export function setCustomProperty(doc: Docx, name: string, value: CustomPropertyValue): void {
  if (name.trim() === "") throw new RangeError("A custom property needs a name.");
  const element = valueElement(value);
  writeCustom(doc, (list) => {
    const props = propertyElements(list);
    const existing = props.find((p) => attrValue(p, "name") === name);
    if (existing) {
      (existing.children as XmlNode[]).splice(0, existing.children.length, element);
      return;
    }
    const pids = props.map((p) => Number(attrValue(p, "pid") ?? 0));
    const pid = Math.max(FIRST_PID - 1, ...pids) + 1;
    list.push({
      kind: "element",
      name: { uri: CUSTOM_NS, local: "property", prefix: "" },
      attrs: [
        plainAttr("fmtid", CUSTOM_FMTID),
        plainAttr("pid", String(pid)),
        plainAttr("name", name),
      ],
      children: [element],
      xmlSpace: "default",
      selfClosing: false,
    });
  });
}

/** Delete a custom property. Returns false when there is none by that name. */
export function removeCustomProperty(doc: Docx, name: string): boolean {
  let removed = false;
  if (!getPart(doc.opc, CUSTOM_PART_NAME)) return false;
  writeCustom(doc, (list) => {
    const at = list.findIndex(
      (c) =>
        c.kind === "element" &&
        c.name.uri === CUSTOM_NS &&
        c.name.local === "property" &&
        attrValue(c, "name") === name,
    );
    if (at >= 0) {
      list.splice(at, 1);
      removed = true;
    }
  });
  return removed;
}
