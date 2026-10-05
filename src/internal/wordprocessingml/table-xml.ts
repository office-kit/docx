/**
 * Small WordprocessingML element helpers for the table tools: build `w:`
 * elements, and insert property children in schema order. The property
 * containers (`<w:tblPr>`, `<w:tcPr>`, …) are `xsd:sequence`s, so a child
 * appended at the end can make the part schema-invalid; Word rejects some of
 * those files as corrupt.
 */

import type { XmlAttr, XmlElement } from "../xml/index.js";
import { WML_NS } from "./namespaces.js";

/** CT_TblPr (§17.4.60) child order. */
export const TBL_PR_ORDER = [
  "tblStyle",
  "tblpPr",
  "tblOverlap",
  "bidiVisual",
  "tblStyleRowBandSize",
  "tblStyleColBandSize",
  "tblW",
  "jc",
  "tblCellSpacing",
  "tblInd",
  "tblBorders",
  "shd",
  "tblLayout",
  "tblCellMar",
  "tblLook",
  "tblCaption",
  "tblDescription",
  "tblPrChange",
] as const;

// CT_TrPr is an unbounded xsd:choice, so any order is valid; Word's order is kept.
export const TR_PR_ORDER = [
  "cnfStyle",
  "divId",
  "gridBefore",
  "gridAfter",
  "wBefore",
  "wAfter",
  "cantSplit",
  "trHeight",
  "tblHeader",
  "tblCellSpacing",
  "jc",
  "hidden",
  "ins",
  "del",
  "trPrChange",
] as const;

/** CT_TcPr (§17.4.70) child order. */
export const TC_PR_ORDER = [
  "cnfStyle",
  "tcW",
  "gridSpan",
  "hMerge",
  "vMerge",
  "tcBorders",
  "shd",
  "noWrap",
  "tcMar",
  "textDirection",
  "tcFitText",
  "vAlign",
  "hideMark",
  "headers",
  "cellIns",
  "cellDel",
  "cellMerge",
  "tcPrChange",
] as const;

/** CT_TcBorders (§17.4.67) / CT_TblBorders (§17.4.38) child order. */
export const BORDER_ORDER = [
  "top",
  "left",
  "start",
  "bottom",
  "right",
  "end",
  "insideH",
  "insideV",
  "tl2br",
  "tr2bl",
] as const;

/** CT_TblCellMar / CT_TcMar child order. */
export const MARGIN_ORDER = ["top", "left", "start", "bottom", "right", "end"] as const;

/** CT_Style (§17.7.4.17) child order. */
export const STYLE_ORDER = [
  "name",
  "aliases",
  "basedOn",
  "next",
  "link",
  "autoRedefine",
  "hidden",
  "uiPriority",
  "semiHidden",
  "unhideWhenUsed",
  "qFormat",
  "locked",
  "personal",
  "personalCompose",
  "personalReply",
  "rsid",
  "pPr",
  "rPr",
  "tblPr",
  "trPr",
  "tcPr",
  "tblStylePr",
] as const;

/** CT_TblStylePr (§17.7.6.6) child order. */
export const TBL_STYLE_PR_ORDER = ["pPr", "rPr", "tblPr", "trPr", "tcPr"] as const;

/** CT_RPr child order (§17.3.2.28), the part the table tools write. */
export const RPR_ORDER = [
  "rStyle",
  "rFonts",
  "b",
  "bCs",
  "i",
  "iCs",
  "caps",
  "smallCaps",
  "strike",
  "dstrike",
  "outline",
  "shadow",
  "emboss",
  "imprint",
  "noProof",
  "snapToGrid",
  "vanish",
  "webHidden",
  "color",
  "spacing",
  "w",
  "kern",
  "position",
  "sz",
  "szCs",
  "highlight",
  "u",
  "effect",
  "bdr",
  "shd",
  "fitText",
  "vertAlign",
  "rtl",
  "cs",
  "em",
  "lang",
  "eastAsianLayout",
  "specVanish",
  "oMath",
] as const;

/** CT_PPr child order (§17.3.1.26), the part the table tools write. */
export const PPR_ORDER = [
  "pStyle",
  "keepNext",
  "keepLines",
  "pageBreakBefore",
  "framePr",
  "widowControl",
  "numPr",
  "suppressLineNumbers",
  "pBdr",
  "shd",
  "tabs",
  "suppressAutoHyphens",
  "kinsoku",
  "wordWrap",
  "overflowPunct",
  "topLinePunct",
  "autoSpaceDE",
  "autoSpaceDN",
  "bidi",
  "adjustRightInd",
  "snapToGrid",
  "spacing",
  "ind",
  "contextualSpacing",
  "mirrorIndents",
  "suppressOverlap",
  "jc",
  "textDirection",
  "textAlignment",
  "textboxTightWrap",
  "outlineLvl",
  "divId",
  "cnfStyle",
  "rPr",
  "sectPr",
  "pPrChange",
] as const;

export function wAttr(local: string, value: string): XmlAttr {
  return { name: { uri: WML_NS, local, prefix: "w" }, value, isNamespaceDecl: false };
}

/** Build `<w:local …attrs>children</w:local>`; undefined attribute values are skipped. */
export function wEl(
  local: string,
  attrs: Readonly<Record<string, string | undefined>> = {},
  children: XmlElement[] = [],
): XmlElement {
  const list: XmlAttr[] = [];
  for (const [name, value] of Object.entries(attrs)) {
    if (value !== undefined) list.push(wAttr(name, value));
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

export function isW(node: XmlElement, local: string): boolean {
  return node.name.uri === WML_NS && node.name.local === local;
}

/** The first `<w:local>` child. */
export function wChild(parent: XmlElement | undefined, local: string): XmlElement | undefined {
  if (!parent) return undefined;
  for (const c of parent.children) {
    if (c.kind === "element" && isW(c, local)) return c;
  }
  return undefined;
}

export function wChildren(parent: XmlElement | undefined, local: string): XmlElement[] {
  if (!parent) return [];
  return parent.children.filter((c): c is XmlElement => c.kind === "element" && isW(c, local));
}

/** A `w:`-namespaced (or unqualified) attribute value. */
export function wAttrOf(el: XmlElement | undefined, local: string): string | undefined {
  return el?.attrs.find((a) => a.name.local === local && (a.name.uri === WML_NS || !a.name.uri))
    ?.value;
}

export function removeWChild(parent: XmlElement, local: string): void {
  const children = parent.children as XmlElement[];
  for (let i = children.length - 1; i >= 0; i--) {
    const c = children[i];
    if (c && c.kind === "element" && isW(c, local)) children.splice(i, 1);
  }
}

/**
 * Put `child` into `parent`, replacing any existing child of the same name,
 * at the position `order` prescribes. Names missing from `order` go last.
 */
export function upsertWChild(
  parent: XmlElement,
  child: XmlElement,
  order: readonly string[],
): XmlElement {
  const local = child.name.local;
  const children = parent.children as XmlElement[];
  const existing = children.findIndex((c) => c.kind === "element" && isW(c, local));
  if (existing >= 0) {
    children.splice(existing, 1, child);
    removeDuplicates(children, local, existing);
    return child;
  }
  const rank = order.indexOf(local);
  const at =
    rank < 0
      ? -1
      : children.findIndex((c) => {
          if (c.kind !== "element" || c.name.uri !== WML_NS) return false;
          const other = order.indexOf(c.name.local);
          return other > rank;
        });
  if (at < 0) children.push(child);
  else children.splice(at, 0, child);
  return child;
}

function removeDuplicates(children: XmlElement[], local: string, keep: number): void {
  for (let i = children.length - 1; i > keep; i--) {
    const c = children[i];
    if (c && c.kind === "element" && isW(c, local)) children.splice(i, 1);
  }
}

/** The existing `<w:local>` child, or a new empty one inserted in schema order. */
export function ensureWChild(
  parent: XmlElement,
  local: string,
  order: readonly string[],
): XmlElement {
  return wChild(parent, local) ?? upsertWChild(parent, wEl(local), order);
}

/** Set (or remove, with undefined) a `w:` attribute. */
export function setWAttr(el: XmlElement, local: string, value: string | undefined): void {
  const attrs = el.attrs as XmlAttr[];
  const index = attrs.findIndex(
    (a) => a.name.local === local && (a.name.uri === WML_NS || !a.name.uri),
  );
  if (value === undefined) {
    if (index >= 0) attrs.splice(index, 1);
  } else if (index >= 0) {
    attrs[index] = wAttr(local, value);
  } else {
    attrs.push(wAttr(local, value));
  }
}

// ST_Border line styles Word offers for table borders (the art borders are
// page-border only).
const LINE_STYLES: ReadonlySet<string> = new Set([
  "single",
  "thick",
  "double",
  "dotted",
  "dashed",
  "dotDash",
  "dotDotDash",
  "triple",
  "thinThickSmallGap",
  "thickThinSmallGap",
  "thinThickThinSmallGap",
  "thinThickMediumGap",
  "thickThinMediumGap",
  "thinThickThinMediumGap",
  "thinThickLargeGap",
  "thickThinLargeGap",
  "thinThickThinLargeGap",
  "wave",
  "doubleWave",
  "dashSmallGap",
  "dashDotStroked",
  "threeDEmboss",
  "threeDEngrave",
  "outset",
  "inset",
]);
// ST_EighthPointMeasure limits for line borders (§17.3.4).
const MIN_BORDER_SIZE = 2;
const MAX_BORDER_SIZE = 96;
const COLOR = /^(auto|[0-9A-Fa-f]{6})$/;

/** `w:val` / `w:sz` / `w:space` / `w:color` of a border, validated; undefined is no border. */
export function borderAttrs(
  border: { readonly style: string; readonly size: number; readonly color: string } | undefined,
): Record<string, string> {
  if (!border) return { val: "nil" };
  if (!LINE_STYLES.has(border.style)) throw new RangeError(`Unknown border style "${border.style}".`);
  if (!Number.isInteger(border.size) || border.size < MIN_BORDER_SIZE || border.size > MAX_BORDER_SIZE) {
    throw new RangeError(`Border size must be ${MIN_BORDER_SIZE}–${MAX_BORDER_SIZE} eighths of a point.`);
  }
  if (!COLOR.test(border.color)) throw new RangeError(`Invalid border color "${border.color}".`);
  return { val: border.style, sz: String(border.size), space: "0", color: border.color };
}
