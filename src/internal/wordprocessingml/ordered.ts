/**
 * Schema order for the containers the References and Mailings features write
 * into. WordprocessingML containers are `xsd:sequence`s, so a child appended
 * at the end can make the part invalid; these insert it at its schema slot.
 */

import type { XmlElement, XmlNode } from "../xml/index.js";
import { WML_NS } from "./namespaces.js";

/** CT_Settings (§17.15.1.78) child order. */
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

/** CT_SectPr (§17.6.17) child order; header/footer references come first. */
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

/** CT_PPr (§17.3.1.26) child order. */
export const P_PR_ORDER: readonly string[] = [
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
];

/** CT_MailMerge (§17.14.20) child order. */
export const MAIL_MERGE_ORDER: readonly string[] = [
  "mainDocumentType",
  "linkToQuery",
  "dataType",
  "connectString",
  "query",
  "dataSource",
  "headerSource",
  "doNotSuppressBlankLines",
  "destination",
  "addressFieldName",
  "mailSubject",
  "mailAsAttachment",
  "viewMergedData",
  "activeRecord",
  "checkErrors",
  "odso",
];

/** CT_Odso (§17.14.25) child order. */
export const ODSO_ORDER: readonly string[] = [
  "udl",
  "table",
  "src",
  "colDelim",
  "type",
  "fHdr",
  "fieldMapData",
  "recipientData",
];

/** CT_FtnProps / CT_EdnProps (§17.11.11): position, format, start, restart. */
export const NOTE_PR_ORDER: readonly string[] = ["pos", "numFmt", "numStart", "numRestart"];

function isWml(node: XmlNode, local: string): node is XmlElement {
  return node.kind === "element" && node.name.uri === WML_NS && node.name.local === local;
}

/** The first `w:local` child, if any. */
export function findChild(parent: XmlElement | undefined, local: string): XmlElement | undefined {
  return parent?.children.find((c): c is XmlElement => isWml(c, local));
}

/**
 * Put `el` into `parent` at its schema position, replacing every existing
 * child of the same name. Children not in `order` keep their place.
 */
export function setOrderedChild(
  parent: XmlElement,
  el: XmlElement,
  order: readonly string[],
): void {
  const children = parent.children as XmlNode[];
  for (let i = children.length - 1; i >= 0; i--) {
    if (isWml(children[i] as XmlNode, el.name.local)) children.splice(i, 1);
  }
  const rank = order.indexOf(el.name.local);
  let at = children.length;
  if (rank >= 0) {
    const later = children.findIndex(
      (c) => c.kind === "element" && c.name.uri === WML_NS && order.indexOf(c.name.local) > rank,
    );
    if (later >= 0) at = later;
  }
  children.splice(at, 0, el);
}

/** Remove every `w:local` child. */
export function removeChild(parent: XmlElement, local: string): void {
  const children = parent.children as XmlNode[];
  for (let i = children.length - 1; i >= 0; i--) {
    if (isWml(children[i] as XmlNode, local)) children.splice(i, 1);
  }
}
