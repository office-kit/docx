/**
 * Whole-element access to `word/settings.xml` for API modules that write
 * structured settings (`w:documentProtection`, `w:zoom` …). Not part of the
 * public API. Everything goes through the raw-part cache, the same tree
 * `setDocumentSettingVal` and the raw-XML inspector edit, so no writer works
 * on a stale copy of the part.
 */

import { WML_NS } from "../internal/wordprocessingml/namespaces.js";
import { insertOrderedChild, SETTINGS_ORDER } from "../internal/wordprocessingml/schema-order.js";
import type { XmlElement } from "../internal/xml/index.js";
import { type Docx, getRawPartRoot, markRawPartDirty, setDocumentSettingVal } from "./docx.js";

const SETTINGS_PART_NAME = "/word/settings.xml";

/** One `<w:settings>` child, or `undefined` when the document has none. */
export function readSetting(doc: Docx, local: string): XmlElement | undefined {
  return getRawPartRoot(doc, SETTINGS_PART_NAME)?.children.find(
    (c): c is XmlElement => c.kind === "element" && c.name.uri === WML_NS && c.name.local === local,
  );
}

/** Replace (or, with `undefined`, remove) one `<w:settings>` child, at its schema position. */
export function replaceSetting(doc: Docx, local: string, element: XmlElement | undefined): void {
  // Creates the settings part when the document has none, and drops the old child.
  setDocumentSettingVal(doc, local, undefined);
  if (!element) return;
  const root = getRawPartRoot(doc, SETTINGS_PART_NAME);
  if (!root) throw new Error("word/settings.xml could not be read");
  insertOrderedChild(root, element, SETTINGS_ORDER);
  markRawPartDirty(doc, SETTINGS_PART_NAME);
  doc.dirty = true;
}
