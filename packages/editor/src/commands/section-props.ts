/**
 * Section property commands for the long tail of `<w:sectPr>`: title page,
 * vertical alignment, text direction, RTL, form protection, endnote
 * suppression, on the section holding the caret. The typed ones go through
 * `setSectionProperties` (which keeps `CT_SectPr` order); form protection and
 * endnote suppression use the generic on-off setter.
 */

import {
  getElementProp,
  getSectionProperties,
  type SectionPropertiesPatch,
  type SectionTextDirection,
  sectionCount,
  setElementOnOff,
  setSectionProperties,
  type VerticalAlignment,
  type XmlElement,
} from "@office-kit/docx";
import type { EditorModel } from "../model.js";
import { caretSection } from "./section.js";
import type { Command } from "./types.js";

/** The caret section's `<w:sectPr>`, if it exists yet (the body's may not). */
function findCaretSectPr(model: EditorModel): XmlElement | undefined {
  const doc = model.doc;
  const index = caretSection(model) ?? sectionCount(doc) - 1;
  let seen = -1;
  for (const block of doc.document.body.blocks) {
    if (block.kind !== "paragraph" || !block.pPr) continue;
    const sectPr = block.pPr.children.find(
      (c): c is XmlElement => c.kind === "element" && c.name.local === "sectPr",
    );
    if (sectPr && ++seen === index) return sectPr;
  }
  return doc.document.body.sectPr;
}

/** The caret section's `<w:sectPr>`, creating the body's when missing. */
function caretSectPr(model: EditorModel): XmlElement {
  // An empty patch creates a missing body sectPr without changing anything else.
  setSectionProperties(model.doc, {}, caretSection(model));
  const sectPr = findCaretSectPr(model);
  if (!sectPr) throw new Error("The document has no section properties.");
  return sectPr;
}

type BoolKey = "titlePage" | "bidi" | "rtlGutter";
const TYPED_ONOFF: ReadonlyArray<{ local: string; key: BoolKey; label: string }> = [
  { local: "titlePg", key: "titlePage", label: "Different first page" },
  { local: "bidi", key: "bidi", label: "Right-to-left section" },
  { local: "rtlGutter", key: "rtlGutter", label: "RTL gutter" },
];
const GENERIC_ONOFF = [
  { local: "noEndnote", label: "Suppress endnotes" },
  { local: "formProt", label: "Protect form" },
];

export const sectionTextDirectionCommand: Command<{ val: SectionTextDirection | undefined }> = {
  id: "section.textDirection",
  group: "section",
  label: "Section text direction",
  run(model, { val }) {
    setSectionProperties(model.doc, { textDirection: val ?? "lrTb" }, caretSection(model));
  },
};

export const sectionVerticalAlignCommand: Command<{ val: VerticalAlignment | undefined }> = {
  id: "section.vAlign",
  group: "section",
  label: "Vertical page alignment",
  run(model, { val }) {
    setSectionProperties(model.doc, { verticalAlignment: val ?? "top" }, caretSection(model));
  },
  isActive: (model) =>
    getSectionProperties(model.doc, caretSection(model)).verticalAlignment !== "top",
};

export const sectionPropertyCommands: Command<never>[] = [
  ...TYPED_ONOFF.map(
    ({ local, key, label }): Command<void> => ({
      id: `section.${local}`,
      group: "section",
      label,
      run(model) {
        const section = caretSection(model);
        const patch: SectionPropertiesPatch = {
          [key]: !getSectionProperties(model.doc, section)[key],
        };
        setSectionProperties(model.doc, patch, section);
      },
      isActive: (model) => getSectionProperties(model.doc, caretSection(model))[key],
    }),
  ),
  ...GENERIC_ONOFF.map(
    ({ local, label }): Command<void> => ({
      id: `section.${local}`,
      group: "section",
      label,
      run(model) {
        const sectPr = caretSectPr(model);
        setElementOnOff(sectPr, local, !getElementProp(sectPr, local).present);
      },
      isActive: (model) => getElementProp(findCaretSectPr(model), local).present,
    }),
  ),
  sectionTextDirectionCommand,
  sectionVerticalAlignCommand,
] as Command<never>[];

/** `w:<local>` element names these commands make editable. */
export const SECTION_PROPERTY_ELEMENTS = [
  ...TYPED_ONOFF.map((e) => e.local),
  ...GENERIC_ONOFF.map((e) => e.local),
  "textDirection",
  "vAlign",
];
