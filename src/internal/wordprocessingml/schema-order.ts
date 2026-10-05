import type { XmlElement, XmlNode } from "../xml/index.js";
import { WML_NS } from "./namespaces.js";

// `CT_SectPr` and `CT_Settings` are `xsd:sequence`s: Word rejects a file
// ("unreadable content") when their children are out of schema order, so
// writers that add a child must insert it at its schema position rather than
// append it. The lists below are the sequences of ECMA-376 Part 1 §17.6.17
// (`sectPr`) and §17.15.1.78 (`settings`), by local name.

/** Children of `<w:sectPr>` in schema order (header/footer references first). */
export const SECT_PR_ORDER: readonly string[] = [
  "headerReference",
  "footerReference",
  "footnotePr",
  "endnotePr",
  "type",
  "pgSz",
  "pgMar",
  "paperSrc",
  "pgBorders",
  "lnNumType",
  "pgNumType",
  "cols",
  "formProt",
  "vAlign",
  "noEndnote",
  "titlePg",
  "textDirection",
  "bidi",
  "rtlGutter",
  "docGrid",
  "printerSettings",
  "sectPrChange",
];

/** Children of `<w:settings>` in schema order. */
export const SETTINGS_ORDER: readonly string[] = [
  "writeProtection",
  "view",
  "zoom",
  "removePersonalInformation",
  "removeDateAndTime",
  "doNotDisplayPageBoundaries",
  "displayBackgroundShape",
  "printPostScriptOverText",
  "printFractionalCharacterWidth",
  "printFormsData",
  "embedTrueTypeFonts",
  "embedSystemFonts",
  "saveSubsetFonts",
  "saveFormsData",
  "mirrorMargins",
  "alignBordersAndEdges",
  "bordersDoNotSurroundHeader",
  "bordersDoNotSurroundFooter",
  "gutterAtTop",
  "hideSpellingErrors",
  "hideGrammaticalErrors",
  "activeWritingStyle",
  "proofState",
  "formsDesign",
  "attachedTemplate",
  "linkStyles",
  "stylePaneFormatFilter",
  "stylePaneSortMethod",
  "documentType",
  "mailMerge",
  "revisionView",
  "trackRevisions",
  "doNotTrackMoves",
  "doNotTrackFormatting",
  "documentProtection",
  "autoFormatOverride",
  "styleLockTheme",
  "styleLockQFSet",
  "defaultTabStop",
  "autoHyphenation",
  "consecutiveHyphenLimit",
  "hyphenationZone",
  "doNotHyphenateCaps",
  "showEnvelope",
  "summaryLength",
  "clickAndTypeStyle",
  "defaultTableStyle",
  "evenAndOddHeaders",
  "bookFoldRevPrinting",
  "bookFoldPrinting",
  "bookFoldPrintingSheets",
  "drawingGridHorizontalSpacing",
  "drawingGridVerticalSpacing",
  "displayHorizontalDrawingGridEvery",
  "displayVerticalDrawingGridEvery",
  "doNotUseMarginsForDrawingGridOrigin",
  "drawingGridHorizontalOrigin",
  "drawingGridVerticalOrigin",
  "doNotShadeFormData",
  "noPunctuationKerning",
  "characterSpacingControl",
  "printTwoOnOne",
  "strictFirstAndLastChars",
  "noLineBreaksAfter",
  "noLineBreaksBefore",
  "savePreviewPicture",
  "doNotValidateAgainstSchema",
  "saveInvalidXml",
  "ignoreMixedContent",
  "alwaysShowPlaceholderText",
  "doNotDemarcateInvalidXml",
  "saveXmlDataOnly",
  "useXSLTWhenSaving",
  "saveThroughXslt",
  "showXMLTags",
  "alwaysMergeEmptyNamespace",
  "updateFields",
  "hdrShapeDefaults",
  "footnotePr",
  "endnotePr",
  "compat",
  "docVars",
  "rsids",
  "mathPr",
  "attachedSchema",
  "themeFontLang",
  "clrSchemeMapping",
  "doNotIncludeSubdocsInStats",
  "doNotAutoCompressPictures",
  "forceUpgrade",
  "captions",
  "readModeInkLockDown",
  "smartTagType",
  "schemaLibrary",
  "shapeDefaults",
  "doNotEmbedSmartTags",
  "decimalSymbol",
  "listSeparator",
];

function isWml(node: XmlNode, local: string): node is XmlElement {
  return node.kind === "element" && node.name.uri === WML_NS && node.name.local === local;
}

/** Remove every direct WML child named `local`. */
export function removeOrderedChild(parent: XmlElement, local: string): void {
  const children = parent.children as XmlNode[];
  for (let i = children.length - 1; i >= 0; i--) {
    const c = children[i];
    if (c && isWml(c, local)) children.splice(i, 1);
  }
}

/**
 * Replace the WML children named `el`'s local name with `el`, at its schema
 * position: after the last child that the schema orders at or before it.
 * Children the order list does not know (`w14:` / `w15:` extensions) never
 * move the insertion point.
 */
export function upsertOrderedChild(
  parent: XmlElement,
  el: XmlElement,
  order: readonly string[],
): void {
  const local = el.name.local;
  removeOrderedChild(parent, local);
  insertOrderedChild(parent, el, order);
}

/** Insert `el` at its schema position without removing same-named siblings. */
export function insertOrderedChild(
  parent: XmlElement,
  el: XmlElement,
  order: readonly string[],
): void {
  const ranks = new Map(order.map((local, i) => [local, i]));
  const rank = ranks.get(el.name.local) ?? -1;
  const children = parent.children as XmlNode[];
  let at = 0;
  for (let i = 0; i < children.length; i++) {
    const c = children[i];
    // Ranked by local name alone: `m:mathPr` and `sl:schemaLibrary` sit in the
    // settings sequence under their own namespaces.
    if (!c || c.kind !== "element") continue;
    const r = ranks.get(c.name.local);
    if (r !== undefined && r <= rank) at = i + 1;
  }
  children.splice(at, 0, el);
}

function wmlElement(local: string, val: string | undefined): XmlElement {
  return {
    kind: "element",
    name: { uri: WML_NS, local, prefix: "w" },
    attrs:
      val === undefined
        ? []
        : [{ name: { uri: WML_NS, local: "val", prefix: "w" }, value: val, isNamespaceDecl: false }],
    children: [],
    xmlSpace: "default",
    selfClosing: true,
  };
}

/** Add (at its schema position) or remove an on-off child such as `<w:titlePg/>`. */
export function setOrderedOnOff(
  parent: XmlElement,
  local: string,
  on: boolean,
  order: readonly string[],
): void {
  if (on) upsertOrderedChild(parent, wmlElement(local, undefined), order);
  else removeOrderedChild(parent, local);
}

/** Set (at its schema position) or, with `undefined`, remove a `<w:x w:val="…"/>` child. */
export function setOrderedVal(
  parent: XmlElement,
  local: string,
  val: string | undefined,
  order: readonly string[],
): void {
  if (val === undefined) removeOrderedChild(parent, local);
  else upsertOrderedChild(parent, wmlElement(local, val), order);
}
