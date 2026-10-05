/**
 * `word/settings.xml` access shared by the API modules. Not part of the public
 * API. Everything goes through the raw-part cache — the one tree
 * `setDocumentSettingVal`, the feature APIs and the raw-XML inspector all
 * edit — so no writer works on a stale copy of the part, and every child is
 * kept in `CT_Settings` schema order, which Word requires.
 */

import {
  addPart,
  addRelationship,
  getPart,
  type Part,
  partRelationships,
  relationshipsByType,
} from "../internal/opc/index.js";
import { WML_NS } from "../internal/wordprocessingml/namespaces.js";
import {
  insertOrderedChild,
  SETTINGS_ORDER,
  sortOrderedChildren,
} from "../internal/wordprocessingml/schema-order.js";
import { WML_CONTENT_TYPES, WML_RELATIONSHIPS } from "../internal/wordprocessingml/index.js";
import type { XmlElement } from "../internal/xml/index.js";
import { type Docx, getRawPartRoot, markRawPartDirty, setDocumentSettingVal } from "./docx.js";

const SETTINGS_PART_NAME = "/word/settings.xml";
// `r` is declared up front: mail merge settings reference their data source by r:id.
const EMPTY_SETTINGS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="${WML_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>`;

/** Ensure `word/settings.xml` exists (part + relationship) and return its part. */
function ensureSettingsPart(doc: Docx): Part {
  const existing = getPart(doc.opc, SETTINGS_PART_NAME);
  if (existing) return existing;
  const part = addPart(doc.opc, {
    name: SETTINGS_PART_NAME,
    contentType: WML_CONTENT_TYPES.settings,
    data: new TextEncoder().encode(EMPTY_SETTINGS_XML),
  });
  const docRels = partRelationships(doc.opc, doc.partName);
  if (relationshipsByType(docRels, WML_RELATIONSHIPS.settings).length === 0) {
    addRelationship(docRels, { type: WML_RELATIONSHIPS.settings, target: "settings.xml" });
  }
  return part;
}

/** The root of `word/settings.xml`, created if the document has none. */
export function settingsRoot(doc: Docx): XmlElement {
  ensureSettingsPart(doc);
  const root = getRawPartRoot(doc, SETTINGS_PART_NAME);
  if (!root) throw new Error("word/settings.xml could not be read");
  return root;
}

/** The `<w:settings>` root, or `undefined` when the document has none. */
export function readSettings(doc: Docx): XmlElement | undefined {
  return getRawPartRoot(doc, SETTINGS_PART_NAME);
}

/** One `<w:settings>` child, or `undefined` when the document has none. */
export function readSetting(doc: Docx, local: string): XmlElement | undefined {
  return readSettings(doc)?.children.find(
    (c): c is XmlElement => c.kind === "element" && c.name.uri === WML_NS && c.name.local === local,
  );
}

/** Replace (or, with `undefined`, remove) one `<w:settings>` child, at its schema position. */
export function replaceSetting(doc: Docx, local: string, element: XmlElement | undefined): void {
  // Creates the settings part when the document has none, and drops the old child.
  setDocumentSettingVal(doc, local, undefined);
  if (!element) return;
  insertOrderedChild(settingsRoot(doc), element, SETTINGS_ORDER);
  markRawPartDirty(doc, SETTINGS_PART_NAME);
  doc.dirty = true;
}

/**
 * Edit `<w:settings>` in place (creating the part if needed); the children
 * are put back in schema order afterwards.
 */
export function editSettings(doc: Docx, edit: (root: XmlElement) => void): void {
  const root = settingsRoot(doc);
  edit(root);
  sortOrderedChildren(root, SETTINGS_ORDER);
  markRawPartDirty(doc, SETTINGS_PART_NAME);
  doc.dirty = true;
}
