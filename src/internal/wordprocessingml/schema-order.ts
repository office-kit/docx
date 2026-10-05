import type { XmlElement, XmlNode } from "../xml/index.js";
import { WML_NS } from "./namespaces.js";

/**
 * Child sequences of the WordprocessingML property containers, from the
 * ECMA-376 Part 1 schema (wml.xsd). These are `xsd:sequence`s, so their
 * children must appear in this order; Word refuses to open ("unreadable
 * content") a file whose `<w:pPr>` lists `<w:jc>` before `<w:spacing>`.
 *
 * The property setters append a new child at the end, which is the simple
 * thing to do on an in-memory tree; the writer puts the children back in
 * schema order with {@link inSchemaOrder}.
 */
const SEQUENCES: Readonly<Record<string, readonly string[]>> = {
  // CT_ParaRPr's revision marks (ins … moveTo) come first; CT_RPr has none.
  rPr: [
    "ins",
    "del",
    "moveFrom",
    "moveTo",
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
    "rPrChange",
  ],
  pPr: [
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
  ],
  pBdr: ["top", "left", "bottom", "right", "between", "bar"],
  numPr: ["ilvl", "numId", "numberingChange", "ins"],
  style: [
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
  ],
  docDefaults: ["rPrDefault", "pPrDefault"],
  abstractNum: ["nsid", "multiLevelType", "tmpl", "name", "styleLink", "numStyleLink", "lvl"],
  lvl: [
    "start",
    "numFmt",
    "lvlRestart",
    "pStyle",
    "isLgl",
    "suff",
    "lvlText",
    "lvlPicBulletId",
    "legacy",
    "lvlJc",
    "pPr",
    "rPr",
  ],
  num: ["abstractNumId", "lvlOverride"],
  lvlOverride: ["startOverride", "lvl"],
};

const RANKS: ReadonlyMap<string, ReadonlyMap<string, number>> = new Map(
  Object.entries(SEQUENCES).map(([container, order]) => [
    container,
    new Map(order.map((local, rank) => [local, rank])),
  ]),
);

// Containers whose descendants hold one of the sequences above.
const DESCEND: ReadonlySet<string> = new Set([
  ...Object.keys(SEQUENCES),
  "rPrDefault",
  "pPrDefault",
  "pPrChange",
  "rPrChange",
]);

function isWml(node: XmlNode): node is XmlElement {
  return node.kind === "element" && node.name.uri === WML_NS;
}

/**
 * `el` with the children of every property container inside it (itself
 * included) in schema order. Children the schema sequence does not name —
 * extension elements, whitespace, comments — keep their place right after the
 * known child they followed. Returns `el` itself when nothing moves, so an
 * untouched document serializes exactly as it was read.
 */
export function inSchemaOrder(el: XmlElement): XmlElement {
  let changed = false;
  const children = el.children.map((child) => {
    if (!isWml(child) || !DESCEND.has(child.name.local)) return child;
    const ordered = inSchemaOrder(child);
    if (ordered !== child) changed = true;
    return ordered;
  });
  const ranks = el.name.uri === WML_NS ? RANKS.get(el.name.local) : undefined;
  if (ranks) {
    let previous = -1;
    const keyed = children.map((node, index) => {
      const rank = isWml(node) ? ranks.get(node.name.local) : undefined;
      if (rank !== undefined) previous = rank;
      return { node, rank: rank ?? previous, index };
    });
    const sorted = keyed.toSorted((a, b) => a.rank - b.rank || a.index - b.index);
    if (sorted.some((entry, i) => entry.index !== i)) {
      return { ...el, children: sorted.map((entry) => entry.node) };
    }
  }
  return changed ? { ...el, children } : el;
}
