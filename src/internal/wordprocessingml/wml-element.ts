import type { XmlAttr, XmlElement, XmlNode } from "../xml/index.js";
import { WML_NS } from "./namespaces.js";

/**
 * Build a `w:`-namespaced element. Attributes whose value is `undefined` are
 * left out, so optional properties can be passed straight through.
 */
export function wmlElement(
  local: string,
  attrs: Readonly<Record<string, string | number | undefined>> = {},
  children: readonly XmlNode[] = [],
): XmlElement {
  const list: XmlAttr[] = [];
  for (const [name, value] of Object.entries(attrs)) {
    if (value === undefined) continue;
    list.push({
      name: { uri: WML_NS, local: name, prefix: "w" },
      value: String(value),
      isNamespaceDecl: false,
    });
  }
  return {
    kind: "element",
    name: { uri: WML_NS, local, prefix: "w" },
    attrs: list,
    children,
    xmlSpace: "default",
    selfClosing: children.length === 0,
  };
}

/** The first direct `w:` child named `local`. */
export function wmlChild(el: XmlElement | undefined, local: string): XmlElement | undefined {
  return el?.children.find(
    (c): c is XmlElement => c.kind === "element" && c.name.uri === WML_NS && c.name.local === local,
  );
}

/** Every direct `w:` child named `local`. */
export function wmlChildren(el: XmlElement | undefined, local: string): XmlElement[] {
  return (el?.children ?? []).filter(
    (c): c is XmlElement => c.kind === "element" && c.name.uri === WML_NS && c.name.local === local,
  );
}

/** An attribute's value by local name (WordprocessingML attributes are `w:`-qualified). */
export function attrOf(el: XmlElement | undefined, local: string): string | undefined {
  return el?.attrs.find((a) => a.name.local === local && !a.isNamespaceDecl)?.value;
}

/** An integer attribute, or `undefined` when absent or not an integer. */
export function intAttrOf(el: XmlElement | undefined, local: string): number | undefined {
  const raw = attrOf(el, local);
  if (raw === undefined) return undefined;
  const n = Number(raw);
  return Number.isInteger(n) ? n : undefined;
}

// ST_OnOff false spellings (§17.17.4); a bare element or any other value is on.
const OFF_VALUES: ReadonlySet<string> = new Set(["0", "false", "off"]);

/** Whether an on-off attribute or element value means "on". */
export function isOn(value: string | undefined): boolean {
  return value !== undefined && !OFF_VALUES.has(value);
}

/** Whether an on-off property element (`<w:titlePg/>`) is present and on. */
export function onOffChild(el: XmlElement | undefined, local: string): boolean {
  const child = wmlChild(el, local);
  if (!child) return false;
  const val = attrOf(child, "val");
  return val === undefined || isOn(val);
}
