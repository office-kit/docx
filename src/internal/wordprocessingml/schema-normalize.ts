import type { XmlElement, XmlNode } from "../xml/index.js";
import { WML_NS } from "./namespaces.js";

/**
 * Container → its children's local names in order, space-separated (strings
 * keep the writer, which every bundle includes, small).
 */
export type SchemaSequences = Readonly<Record<string, string>>;

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
const SEQUENCES: SchemaSequences = {
  // CT_ParaRPr's revision marks (ins … moveTo) come first; CT_RPr has none.
  rPr: "ins del moveFrom moveTo rStyle rFonts b bCs i iCs caps smallCaps strike dstrike outline shadow emboss imprint noProof snapToGrid vanish webHidden color spacing w kern position sz szCs highlight u effect bdr shd fitText vertAlign rtl cs em lang eastAsianLayout specVanish oMath rPrChange",
  pPr: "pStyle keepNext keepLines pageBreakBefore framePr widowControl numPr suppressLineNumbers pBdr shd tabs suppressAutoHyphens kinsoku wordWrap overflowPunct topLinePunct autoSpaceDE autoSpaceDN bidi adjustRightInd snapToGrid spacing ind contextualSpacing mirrorIndents suppressOverlap jc textDirection textAlignment textboxTightWrap outlineLvl divId cnfStyle rPr sectPr pPrChange",
  pBdr: "top left bottom right between bar",
  numPr: "ilvl numId numberingChange ins",
  style:
    "name aliases basedOn next link autoRedefine hidden uiPriority semiHidden unhideWhenUsed qFormat locked personal personalCompose personalReply rsid pPr rPr tblPr trPr tcPr tblStylePr",
  docDefaults: "rPrDefault pPrDefault",
};

type Ranks = ReadonlyMap<string, ReadonlyMap<string, number>>;

// Rank lookups, built once per sequence table on first use.
const rankCache = new WeakMap<SchemaSequences, Ranks>();

function ranksOf(sequences: SchemaSequences): Ranks {
  let ranks = rankCache.get(sequences);
  if (!ranks) {
    ranks = new Map(
      Object.entries(sequences).map(([container, order]) => [
        container,
        new Map(order.split(" ").map((local, rank) => [local, rank])),
      ]),
    );
    rankCache.set(sequences, ranks);
  }
  return ranks;
}

// Containers that hold a property container without being one.
const WRAPPERS: ReadonlySet<string> = new Set([
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
 * untouched document serializes exactly as it was read. `extra` adds the
 * sequences of another part (numbering definitions).
 */
export function inSchemaOrder(el: XmlElement, extra?: SchemaSequences): XmlElement {
  const own = ranksOf(SEQUENCES);
  const more = extra && ranksOf(extra);
  const ranksFor = (local: string): ReadonlyMap<string, number> | undefined =>
    more?.get(local) ?? own.get(local);
  let changed = false;
  const children = el.children.map((child) => {
    if (!isWml(child)) return child;
    const local = child.name.local;
    if (!WRAPPERS.has(local) && !ranksFor(local)) return child;
    const ordered = inSchemaOrder(child, extra);
    if (ordered !== child) changed = true;
    return ordered;
  });
  const ranks = el.name.uri === WML_NS ? ranksFor(el.name.local) : undefined;
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
