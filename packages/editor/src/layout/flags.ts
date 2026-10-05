/**
 * The pagination-relevant paragraph properties, resolved through the style
 * hierarchy (document defaults → paragraph style chain → direct formatting,
 * §17.7.2): Heading styles carry `w:keepNext`, and Normal usually carries
 * `w:widowControl`, so reading direct formatting alone would paginate wrong.
 */

import {
  childElementsOf,
  type Docx,
  getElementAttr,
  getParagraphStyle,
  stylesPart,
  type WmlParagraph,
  type XmlElement,
} from "@office-kit/docx";

export interface ParagraphFlags {
  readonly keepNext: boolean;
  readonly keepLines: boolean;
  readonly widowControl: boolean;
  readonly pageBreakBefore: boolean;
  /** `w:suppressLineNumbers` (§17.3.1.34). */
  readonly suppressLineNumbers: boolean;
}

const FLAG_ELEMENTS = {
  keepNext: "keepNext",
  keepLines: "keepLines",
  widowControl: "widowControl",
  pageBreakBefore: "pageBreakBefore",
  suppressLineNumbers: "suppressLineNumbers",
} as const satisfies Record<keyof ParagraphFlags, string>;

const OFF_VALUES: ReadonlySet<string> = new Set(["0", "false", "off"]);

function child(el: XmlElement | undefined, local: string): XmlElement | undefined {
  return el && childElementsOf(el).find((c) => c.name.local === local);
}

function apply(out: Record<keyof ParagraphFlags, boolean>, pPr: XmlElement | undefined): void {
  if (!pPr) return;
  for (const [key, local] of Object.entries(FLAG_ELEMENTS) as Array<
    [keyof ParagraphFlags, string]
  >) {
    const el = child(pPr, local);
    if (!el) continue;
    const val = getElementAttr(el, "val");
    out[key] = val === undefined || !OFF_VALUES.has(val);
  }
}

/** Build a flag resolver over the document's styles (rebuild when styles change). */
export function createFlagResolver(doc: Docx): (para: WmlParagraph) => ParagraphFlags {
  const part = stylesPart(doc);
  const byId = new Map<string, XmlElement>();
  let defaultStyle: string | undefined;
  for (const style of part?.styles ?? []) {
    const id = getElementAttr(style, "styleId");
    if (id === undefined) continue;
    byId.set(id, style);
    if (getElementAttr(style, "type") === "paragraph" && getElementAttr(style, "default") === "1") {
      defaultStyle = id;
    }
  }
  const defaultPPr = child(child(part?.docDefaults, "pPrDefault"), "pPr");
  // Resolved per style once: a document has few styles and many paragraphs.
  const styleCache = new Map<string, Record<keyof ParagraphFlags, boolean>>();
  const resolveStyle = (id: string | undefined): Record<keyof ParagraphFlags, boolean> => {
    const key = id ?? "";
    const cached = styleCache.get(key);
    if (cached) return cached;
    const chain: XmlElement[] = [];
    const seen = new Set<string>();
    let current = id;
    // A basedOn cycle is invalid but occurs in the wild; stop at the repeat.
    while (current !== undefined && !seen.has(current)) {
      seen.add(current);
      const style = byId.get(current);
      if (!style) break;
      chain.unshift(style);
      const basedOn = child(style, "basedOn");
      current = basedOn && getElementAttr(basedOn, "val");
    }
    const out = {
      keepNext: false,
      keepLines: false,
      widowControl: false,
      pageBreakBefore: false,
      suppressLineNumbers: false,
    };
    apply(out, defaultPPr);
    for (const style of chain) apply(out, child(style, "pPr"));
    styleCache.set(key, out);
    return out;
  };
  return (para) => {
    const id = getParagraphStyle(para);
    const base = resolveStyle(id !== undefined && byId.has(id) ? id : defaultStyle);
    if (!para.pPr) return base;
    const out = { ...base };
    apply(out, para.pPr);
    return out;
  };
}
