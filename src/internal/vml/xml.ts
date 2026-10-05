/**
 * Namespace-aware helpers for building and editing VML element trees. VML
 * attributes come from several vocabularies (`o:spid`, `r:id`, `w10:…`), and
 * the prefixes a loaded document binds them to vary, so attributes are always
 * matched by namespace URI + local name.
 */

import type { XmlAttr, XmlElement, XmlNode } from "../xml/index.js";
import { REL_NS, WML_NS } from "../wordprocessingml/namespaces.js";
import { OFFICE_NS, VML_NS, WORD_VML_NS } from "./namespaces.js";

const URI_FOR_PREFIX: Readonly<Record<string, string>> = {
  v: VML_NS,
  o: OFFICE_NS,
  w10: WORD_VML_NS,
  w: WML_NS,
  r: REL_NS,
};

/** Split `o:spid` into its namespace and local name (unprefixed → no namespace). */
function qualify(name: string): { prefix: string; uri: string; local: string } {
  const colon = name.indexOf(":");
  if (colon < 0) return { prefix: "", uri: "", local: name };
  const prefix = name.slice(0, colon);
  const uri = URI_FOR_PREFIX[prefix];
  if (uri === undefined) throw new Error(`Unknown VML attribute prefix "${prefix}"`);
  return { prefix, uri, local: name.slice(colon + 1) };
}

export function element(
  name: string,
  attrs: Readonly<Record<string, string | undefined>> = {},
  children: readonly XmlNode[] = [],
): XmlElement {
  const q = qualify(name);
  const list: XmlAttr[] = [];
  for (const [key, value] of Object.entries(attrs)) {
    if (value === undefined) continue;
    const a = qualify(key);
    list.push({
      name: { uri: a.uri, local: a.local, prefix: a.prefix },
      value,
      isNamespaceDecl: false,
    });
  }
  return {
    kind: "element",
    name: { uri: q.uri, local: q.local, prefix: q.prefix },
    attrs: list,
    children: [...children],
    xmlSpace: "default",
    selfClosing: children.length === 0,
  };
}

/** An `xmlns:prefix` declaration, so a VML fragment is self-contained wherever it lands. */
export function nsDecl(prefix: string): XmlAttr {
  const uri = URI_FOR_PREFIX[prefix];
  if (uri === undefined) throw new Error(`Unknown namespace prefix "${prefix}"`);
  return {
    name: { uri: "http://www.w3.org/2000/xmlns/", local: prefix, prefix: "xmlns" },
    value: uri,
    isNamespaceDecl: true,
  };
}

export function getAttr(el: XmlElement, name: string): string | undefined {
  const q = qualify(name);
  return el.attrs.find(
    (a) => !a.isNamespaceDecl && a.name.local === q.local && a.name.uri === q.uri,
  )?.value;
}

/** Set (or with `undefined`, remove) an attribute, keeping the document's own prefix. */
export function setAttr(el: XmlElement, name: string, value: string | undefined): void {
  const q = qualify(name);
  const attrs = el.attrs as XmlAttr[];
  const index = attrs.findIndex(
    (a) => !a.isNamespaceDecl && a.name.local === q.local && a.name.uri === q.uri,
  );
  if (value === undefined) {
    if (index >= 0) attrs.splice(index, 1);
    return;
  }
  const existing = attrs[index];
  const attr: XmlAttr = {
    name: existing?.name ?? { uri: q.uri, local: q.local, prefix: q.prefix },
    value,
    isNamespaceDecl: false,
  };
  if (index >= 0) attrs[index] = attr;
  else attrs.push(attr);
}

export function isElement(el: XmlNode | undefined, name: string): el is XmlElement {
  if (el?.kind !== "element") return false;
  const q = qualify(name);
  return el.name.uri === q.uri && el.name.local === q.local;
}

export function children(el: XmlElement): XmlElement[] {
  return el.children.filter((c): c is XmlElement => c.kind === "element");
}

export function child(el: XmlElement, name: string): XmlElement | undefined {
  return el.children.find((c): c is XmlElement => isElement(c, name));
}

/** Replace (or with `undefined`, remove) the first child of that name, appending if absent. */
export function setChild(parent: XmlElement, name: string, next: XmlElement | undefined): void {
  const list = parent.children as XmlNode[];
  const index = list.findIndex((c) => isElement(c, name));
  if (next === undefined) {
    if (index >= 0) list.splice(index, 1);
  } else if (index >= 0) {
    list[index] = next;
  } else {
    list.push(next);
  }
  (parent as { selfClosing: boolean }).selfClosing = list.length === 0;
}

/** Rename an element in place (e.g. `v:rect` → `v:shape` when the preset changes). */
export function rename(el: XmlElement, name: string): void {
  const q = qualify(name);
  const prefix = el.name.uri === q.uri ? el.name.prefix : q.prefix;
  (el as { name: XmlElement["name"] }).name = { uri: q.uri, local: q.local, prefix };
}
