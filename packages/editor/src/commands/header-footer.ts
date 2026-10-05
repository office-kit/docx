/**
 * Header / footer commands: the Insert ▸ Header / Footer galleries and their
 * Remove entries, the Page Number menu and Format Page Numbers. They act on
 * the last section (the body's `w:sectPr`). `addHeader` / `addFooter` /
 * `addPageNumberFooter` remain the plain "add a part" commands.
 */

import {
  addFooter,
  addHeader,
  addPageNumberFooter,
  type HeaderFooterType,
  insertPageNumbers,
  type PageNumberFormatOptions,
  type PageNumberOptions,
  removeHeaderFooter,
  removePageNumbers,
  setHeaderFooterParagraphs,
  setPageNumberFormat,
} from "@office-kit/docx";
import { type HeaderFooterPresetId, headerFooterPreset } from "./insert-presets.js";
import type { Command } from "./types.js";

export { HEADER_FOOTER_PRESETS, type HeaderFooterPresetId } from "./insert-presets.js";
export { COVER_PAGE_PRESETS, type CoverPagePresetId } from "./insert-presets.js";

export const addHeaderCommand: Command<{ text: string; type?: HeaderFooterType }> = {
  id: "headerFooter.addHeader",
  group: "headerFooter",
  label: "Header",
  run(model, { text, type = "default" }) {
    addHeader(model.doc, text, type);
  },
};

export const addFooterCommand: Command<{ text: string; type?: HeaderFooterType }> = {
  id: "headerFooter.addFooter",
  group: "headerFooter",
  label: "Footer",
  run(model, { text, type = "default" }) {
    addFooter(model.doc, text, type);
  },
};

export const addPageNumberFooterCommand: Command<{ type?: HeaderFooterType }> = {
  id: "headerFooter.pageNumberFooter",
  group: "headerFooter",
  label: "Page number footer",
  run(model, { type = "default" }) {
    addPageNumberFooter(model.doc, "", "", type);
  },
};

type PresetParams = { preset: HeaderFooterPresetId; type?: HeaderFooterType };

export const insertHeaderPresetCommand: Command<PresetParams> = {
  id: "headerFooter.headerPreset",
  group: "headerFooter",
  label: "Header",
  run(model, { preset, type = "default" }) {
    setHeaderFooterParagraphs(
      model.doc,
      "header",
      type,
      headerFooterPreset(model.doc, "header", preset),
    );
  },
};

export const insertFooterPresetCommand: Command<PresetParams> = {
  id: "headerFooter.footerPreset",
  group: "headerFooter",
  label: "Footer",
  run(model, { preset, type = "default" }) {
    setHeaderFooterParagraphs(
      model.doc,
      "footer",
      type,
      headerFooterPreset(model.doc, "footer", preset),
    );
  },
};

export const removeHeaderCommand: Command<{ type?: HeaderFooterType }> = {
  id: "headerFooter.removeHeader",
  group: "headerFooter",
  label: "Remove Header",
  run(model, { type = "default" }) {
    if (!removeHeaderFooter(model.doc, "header", type))
      throw new Error("There is no header to remove.");
  },
};

export const removeFooterCommand: Command<{ type?: HeaderFooterType }> = {
  id: "headerFooter.removeFooter",
  group: "headerFooter",
  label: "Remove Footer",
  run(model, { type = "default" }) {
    if (!removeHeaderFooter(model.doc, "footer", type))
      throw new Error("There is no footer to remove.");
  },
};

export const insertPageNumberCommand: Command<PageNumberOptions> = {
  id: "headerFooter.pageNumber",
  group: "headerFooter",
  label: "Page Number",
  run(model, options) {
    insertPageNumbers(model.doc, options);
  },
};

export const formatPageNumbersCommand: Command<PageNumberFormatOptions> = {
  id: "headerFooter.pageNumberFormat",
  group: "headerFooter",
  label: "Format Page Numbers",
  run(model, options) {
    setPageNumberFormat(model.doc, options);
  },
};

export const removePageNumbersCommand: Command<void, number> = {
  id: "headerFooter.removePageNumbers",
  group: "headerFooter",
  label: "Remove Page Numbers",
  run(model) {
    return removePageNumbers(model.doc);
  },
};

export const headerFooterCommands = [
  addHeaderCommand,
  addFooterCommand,
  addPageNumberFooterCommand,
  insertHeaderPresetCommand,
  insertFooterPresetCommand,
  removeHeaderCommand,
  removeFooterCommand,
  insertPageNumberCommand,
  formatPageNumbersCommand,
  removePageNumbersCommand,
];
