/**
 * Header & Footer commands (Word's contextual Header & Footer tab): entering
 * a header or footer, the section's first-page and odd/even settings, the
 * header/footer distances, and Link to Previous.
 *
 * Sections are addressed by index (see `sectionProperties`), because the
 * header being edited belongs to the section of the page it was opened on,
 * not to wherever the caret is in the body. They read the real document
 * (`model.document`): while the caret is in a header, `model.doc` is a view
 * whose body is that header.
 */

import {
  type Docx,
  ensureHeaderFooter,
  getDocumentSetting,
  type HeaderFooterKind,
  type HeaderFooterType,
  isHeaderFooterLinked,
  makePropsElement,
  resolveHeaderFooter,
  sectionProperties,
  setDocumentSettingOnOff,
  setElementAttr,
  setElementOnOff,
  setHeaderFooterLinked,
  storyBody,
  type XmlElement,
} from "@office-kit/docx";
import type { EditorModel } from "../model.js";
import { caretAt } from "../selection.js";
import type { Command } from "./types.js";

/** Which header or footer of which section. */
export interface HeaderFooterTarget {
  readonly section: number;
  readonly kind: HeaderFooterKind;
  readonly type: HeaderFooterType;
}

function sectPrFor(doc: Docx, section: number): XmlElement {
  const all = sectionProperties(doc);
  const existing = all[section];
  if (existing) return existing;
  if (section !== all.length - 1) throw new Error(`Section ${section} does not exist.`);
  const sectPr = makePropsElement("sectPr");
  doc.document.body.sectPr = sectPr;
  return sectPr;
}

function child(el: XmlElement, local: string): XmlElement | undefined {
  return el.children.find((c): c is XmlElement => c.kind === "element" && c.name.local === local);
}

function ensureChild(el: XmlElement, local: string): XmlElement {
  const existing = child(el, local);
  if (existing) return existing;
  const created = makePropsElement(local);
  (el.children as XmlElement[]).push(created);
  return created;
}

const OFF: ReadonlySet<string> = new Set(["0", "false", "off"]);

/** The documented range of the pgMar header/footer distances (ST_TwipsMeasure, 0 – 22"). */
const MAX_DISTANCE_TWIPS = 31680;

/**
 * Open a header or footer for editing: create it when the section shows none
 * (as double-clicking an empty header area does in Word), then put the caret
 * at its start.
 */
export const editHeaderFooterCommand: Command<HeaderFooterTarget> = {
  id: "stories.editHeaderFooter",
  group: "headerFooter",
  label: "Edit Header or Footer",
  run(model, { section, kind, type }) {
    const partName = ensureHeaderFooter(model.document, section, kind, type);
    const story = { kind, partName } as const;
    if (!storyBody(model.document, story)?.blocks.length) {
      throw new Error(`The ${kind} part ${partName} has no paragraphs.`);
    }
    model.setSelection(caretAt({ story, block: 0, inline: 0, offset: 0 }));
  },
};

function titlePageOn(model: EditorModel, section: number): boolean {
  const titlePg = sectionProperties(model.document)[section];
  const el = titlePg && child(titlePg, "titlePg");
  const val = el?.attrs.find((a) => a.name.local === "val")?.value;
  return !!el && (val === undefined || !OFF.has(val));
}

export const differentFirstPageCommand: Command<{ section: number; on: boolean }> = {
  id: "stories.differentFirstPage",
  group: "headerFooter",
  label: "Different First Page",
  run(model, { section, on }) {
    setElementOnOff(sectPrFor(model.document, section), "titlePg", on);
  },
};

/** Whether a section has a distinct first-page header/footer (for the checkbox). */
export function hasDifferentFirstPage(model: EditorModel, section: number): boolean {
  return titlePageOn(model, section);
}

export const differentOddEvenCommand: Command<{ on: boolean }> = {
  id: "stories.differentOddEven",
  group: "headerFooter",
  label: "Different Odd & Even Pages",
  run(model, { on }) {
    setDocumentSettingOnOff(model.document, "evenAndOddHeaders", on);
  },
  isActive(model) {
    const setting = getDocumentSetting(model.document, "evenAndOddHeaders");
    return setting.present && !OFF.has(setting.val ?? "1");
  },
};

function distanceCommand(
  kind: HeaderFooterKind,
  label: string,
): Command<{ section: number; twips: number }> {
  return {
    id: kind === "header" ? "stories.headerDistance" : "stories.footerDistance",
    group: "headerFooter",
    label,
    run(model, { section, twips }) {
      if (!Number.isInteger(twips) || twips < 0 || twips > MAX_DISTANCE_TWIPS) {
        throw new Error(`Distance must be 0 – ${MAX_DISTANCE_TWIPS} twips, got ${twips}.`);
      }
      setElementAttr(ensureChild(sectPrFor(model.document, section), "pgMar"), kind, String(twips));
    },
  };
}

export const headerDistanceCommand = distanceCommand("header", "Header from Top");
export const footerDistanceCommand = distanceCommand("footer", "Footer from Bottom");

export const linkToPreviousCommand: Command<HeaderFooterTarget & { linked: boolean }> = {
  id: "stories.linkToPrevious",
  group: "headerFooter",
  label: "Link to Previous",
  run(model, { section, kind, type, linked }) {
    setHeaderFooterLinked(model.document, section, kind, type, linked);
    // The caret was in the header the section showed; after (un)linking it
    // shows another part, so follow it there — or, when linking to a
    // section without one, back to the body.
    const partName = resolveHeaderFooter(model.document, section, kind, type)?.partName;
    model.setSelection(
      caretAt(
        partName
          ? { story: { kind, partName }, block: 0, inline: 0, offset: 0 }
          : { block: 0, inline: 0, offset: 0 },
      ),
    );
  },
};

/** Whether the section shows the previous section's header/footer (Link to Previous pressed). */
export function headerFooterLinked(model: EditorModel, target: HeaderFooterTarget): boolean {
  return (
    target.section > 0 &&
    isHeaderFooterLinked(model.document, target.section, target.kind, target.type)
  );
}

export const storyCommands: ReadonlyArray<Command<never, unknown>> = [
  editHeaderFooterCommand,
  differentFirstPageCommand,
  differentOddEvenCommand,
  headerDistanceCommand,
  footerDistanceCommand,
  linkToPreviousCommand,
] as ReadonlyArray<Command<never, unknown>>;
