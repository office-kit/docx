/**
 * Element-level helpers for editing DrawingML trees in place.
 *
 * DrawingML content models are strict sequences (ECMA-376 Part 1, §20/§21),
 * so a child must be inserted at its schema position, not appended. Every
 * setter here takes the parent's child order and slots the new element in.
 */

import { parseXml, type XmlAttr, type XmlElement, type XmlNode } from "../xml/index.js";

export const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
export const WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
export const PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture";
export const C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart";
export const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/** The prefixes this module writes, bound once on each `<w:drawing>` it edits. */
export const DRAWING_PREFIXES: Readonly<Record<string, string>> = {
  wp: WP_NS,
  a: A_NS,
  pic: PIC_NS,
  r: R_NS,
  c: C_NS,
};

const FRAGMENT_DECLS = Object.entries(DRAWING_PREFIXES)
  .map(([prefix, uri]) => `xmlns:${prefix}="${uri}"`)
  .join(" ");

/**
 * Parse markup written with the {@link DRAWING_PREFIXES} prefixes into a
 * detached element. The declarations live on a throw-away wrapper, so the
 * returned element carries none: callers insert it under a `<w:drawing>` that
 * {@link ensureDrawingNamespaces} has prepared.
 */
export function fragment(xml: string): XmlElement {
  const root = parseXml(`<wrap ${FRAGMENT_DECLS}>${xml}</wrap>`).root;
  const first = root.children.find((c): c is XmlElement => c.kind === "element");
  if (!first) throw new Error("fragment: no element in markup");
  return first;
}

/**
 * Bind every prefix this module writes on the `<w:drawing>` element itself.
 * Word declares `a:` on `<a:graphic>` and `pic:` on `<pic:pic>` rather than on
 * the document root, so an element inserted elsewhere (a new `wp:wrapSquare`,
 * an `a:effectLst`) needs a binding in scope; repeating one is harmless.
 */
export function ensureDrawingNamespaces(drawing: XmlElement): void {
  const attrs = drawing.attrs as XmlAttr[];
  for (const [prefix, uri] of Object.entries(DRAWING_PREFIXES)) {
    const declared = attrs.some((a) => a.isNamespaceDecl && a.name.local === prefix);
    if (!declared) {
      attrs.push({
        name: { uri: "", local: prefix, prefix: "xmlns" },
        value: uri,
        isNamespaceDecl: true,
      });
    }
  }
}

export function escapeXml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

export function elementChildren(el: XmlElement): XmlElement[] {
  return el.children.filter((c): c is XmlElement => c.kind === "element");
}

export function child(
  el: XmlElement | undefined,
  uri: string,
  local: string,
): XmlElement | undefined {
  if (!el) return undefined;
  for (const c of el.children) {
    if (c.kind === "element" && c.name.uri === uri && c.name.local === local) return c;
  }
  return undefined;
}

/** Follow a path of `[uri, local]` steps through direct children. */
export function path(
  el: XmlElement | undefined,
  ...steps: ReadonlyArray<readonly [string, string]>
): XmlElement | undefined {
  let cur = el;
  for (const [uri, local] of steps) cur = child(cur, uri, local);
  return cur;
}

export function descendant(el: XmlElement, uri: string, local: string): XmlElement | undefined {
  for (const c of el.children) {
    if (c.kind !== "element") continue;
    if (c.name.uri === uri && c.name.local === local) return c;
    const found = descendant(c, uri, local);
    if (found) return found;
  }
  return undefined;
}

export function attr(el: XmlElement | undefined, local: string): string | undefined {
  return el?.attrs.find((a) => a.name.local === local && !a.isNamespaceDecl)?.value;
}

/** An attribute in a namespace (e.g. `r:embed`), matched on URI. */
export function nsAttr(el: XmlElement | undefined, uri: string, local: string): string | undefined {
  return el?.attrs.find((a) => a.name.uri === uri && a.name.local === local)?.value;
}

export function numAttr(el: XmlElement | undefined, local: string): number | undefined {
  const raw = attr(el, local);
  if (raw === undefined) return undefined;
  const n = Number(raw);
  return Number.isFinite(n) ? n : undefined;
}

/** xsd:boolean: `1` / `true` are true. */
export function boolAttr(el: XmlElement | undefined, local: string): boolean | undefined {
  const raw = attr(el, local);
  return raw === undefined ? undefined : raw === "1" || raw === "true";
}

export function setAttr(el: XmlElement, local: string, value: string | undefined): void {
  const attrs = el.attrs as XmlAttr[];
  const index = attrs.findIndex((a) => a.name.local === local && a.name.prefix === "");
  if (value === undefined) {
    if (index >= 0) attrs.splice(index, 1);
    return;
  }
  const next: XmlAttr = { name: { uri: "", local, prefix: "" }, value, isNamespaceDecl: false };
  if (index >= 0) attrs[index] = next;
  else attrs.push(next);
}

/**
 * Put attributes in a fixed order (unknown ones keep their relative order at
 * the end). Attribute order is not significant in XML; this keeps output
 * readable and diff-stable next to what Word writes.
 */
export function orderAttrs(el: XmlElement, order: readonly string[]): void {
  const attrs = el.attrs as XmlAttr[];
  const rank = (a: XmlAttr): number => {
    if (a.isNamespaceDecl) return -1;
    const i = order.indexOf(a.name.local);
    return i < 0 ? order.length : i;
  };
  attrs.sort((x, y) => rank(x) - rank(y));
}

export function removeChildren(parent: XmlElement, uri: string, locals: ReadonlySet<string>): void {
  const kids = parent.children as XmlNode[];
  for (let i = kids.length - 1; i >= 0; i--) {
    const k = kids[i];
    if (k?.kind === "element" && k.name.uri === uri && locals.has(k.name.local)) kids.splice(i, 1);
  }
}

/**
 * Insert `el` into `parent` at its schema position. `order` lists the local
 * names of the parent's content model in sequence (names in one choice group
 * share a slot when listed in the same `|`-joined entry, e.g.
 * `"wrapNone|wrapSquare|wrapTight"`). Existing children in the same slot are
 * replaced.
 */
export function placeChild(parent: XmlElement, el: XmlElement, order: readonly string[]): void {
  const slotOf = (local: string): number =>
    order.findIndex((entry) => entry.split("|").includes(local));
  const slot = slotOf(el.name.local);
  if (slot < 0) throw new Error(`placeChild: ${el.name.local} is not in the content model`);
  const kids = parent.children as XmlNode[];
  for (let i = kids.length - 1; i >= 0; i--) {
    const k = kids[i];
    if (k?.kind === "element" && k.name.uri === el.name.uri && slotOf(k.name.local) === slot) {
      kids.splice(i, 1);
    }
  }
  let at = kids.length;
  for (let i = 0; i < kids.length; i++) {
    const k = kids[i];
    if (k?.kind !== "element") continue;
    const s = slotOf(k.name.local);
    if (s > slot) {
      at = i;
      break;
    }
  }
  kids.splice(at, 0, el);
}

/** Remove the child in `el`'s slot of `order` (whatever member of a choice is there). */
export function clearSlot(
  parent: XmlElement,
  uri: string,
  local: string,
  order: readonly string[],
): void {
  const entry = order.find((e) => e.split("|").includes(local));
  if (!entry) return;
  removeChildren(parent, uri, new Set(entry.split("|")));
}
