/**
 * Layout tab commands beyond the plain page size / margins / orientation
 * ones in `section.ts`: Text Direction, Columns, Breaks, Line Numbers,
 * Hyphenation, the Indent / Spacing boxes, and the Page Setup dialog (which
 * applies all its tabs at once, as Word's OK button does).
 */

import {
  childElementsOf,
  getDocumentSetting,
  getElementAttr,
  getParagraphProp,
  getSectionProperties,
  type LineNumbering,
  type PageMargins,
  type PageSize,
  type SectionColumns,
  type SectionPropertiesPatch,
  type SectionTextDirection,
  setDocumentSettingOnOff,
  setDocumentSettingVal,
  setPageMargins,
  setPageSize,
  setParagraphIndent,
  setParagraphOnOff,
  setParagraphSpacing,
  setSectionProperties,
  type WmlParagraph,
} from "@office-kit/docx";
import { paragraphsInRange } from "../doc-access.js";
import type { EditorModel } from "../model.js";
import { orderSelection } from "../selection.js";
import { caretSection, type SectionTarget, sectionScopeFor } from "./section.js";
import { insertLineBreakCommand } from "./structure.js";
import type { Command } from "./types.js";

function selectedParagraphs(model: EditorModel): WmlParagraph[] {
  const sel = model.selection;
  return sel ? paragraphsInRange(model.doc, orderSelection(sel)) : [];
}

/** The caret section's properties (the last section without a caret). */
export function currentSectionProperties(
  model: EditorModel,
): ReturnType<typeof getSectionProperties> {
  return getSectionProperties(model.doc, caretSection(model));
}

export const textDirectionCommand: Command<{
  direction: SectionTextDirection;
  target?: SectionTarget;
}> = {
  id: "layout.textDirection",
  group: "layout",
  label: "Text Direction",
  run(model, { direction, target = "section" }) {
    setSectionProperties(
      model.doc,
      { textDirection: direction },
      sectionScopeFor(model, target, "nextPage"),
    );
  },
};

/** Columns presets and the Columns dialog. "This point forward" starts a continuous section. */
export const columnsCommand: Command<{ columns: SectionColumns; target?: SectionTarget }> = {
  id: "layout.columns",
  group: "layout",
  label: "Columns",
  run(model, { columns, target = "section" }) {
    setSectionProperties(model.doc, { columns }, sectionScopeFor(model, target, "continuous"));
  },
};

/** Breaks ▸ Page / Column / Text Wrapping: a `<w:br>` at the caret. */
export const insertBreakCommand: Command<{ kind: "page" | "column" | "textWrapping" }> = {
  id: "layout.break",
  group: "layout",
  label: "Breaks",
  run(model, { kind }) {
    insertLineBreakCommand.run(model, { kind });
  },
  isEnabled: (model) => insertLineBreakCommand.isEnabled?.(model) ?? true,
};

/** Line Numbers menu and the Line Numbers dialog; `null` is None. */
export const lineNumbersCommand: Command<{
  lineNumbering: LineNumbering | null;
  target?: SectionTarget;
}> = {
  id: "layout.lineNumbers",
  group: "layout",
  label: "Line Numbers",
  run(model, { lineNumbering, target = "section" }) {
    setSectionProperties(model.doc, { lineNumbering }, sectionScopeFor(model, target, "nextPage"));
  },
};

/** Line Numbers ▸ Suppress for Current Paragraph (`w:suppressLineNumbers`). */
export const suppressLineNumbersCommand: Command<void> = {
  id: "layout.suppressLineNumbers",
  group: "layout",
  label: "Suppress for Current Paragraph",
  run(model) {
    const paras = selectedParagraphs(model);
    const on = !paras.every((p) => getParagraphProp(p, "suppressLineNumbers").present);
    for (const p of paras) setParagraphOnOff(p, "suppressLineNumbers", on);
  },
  isEnabled: (model) => selectedParagraphs(model).length > 0,
  isActive: (model) => {
    const paras = selectedParagraphs(model);
    return (
      paras.length > 0 && paras.every((p) => getParagraphProp(p, "suppressLineNumbers").present)
    );
  },
};

/** Hyphenation options (`w:settings`); omitted fields are left alone. */
export interface HyphenationOptions {
  /** Hyphenate the document automatically (`w:autoHyphenation`). */
  readonly automatic?: boolean;
  /** Hyphenate words in CAPS (the inverse of `w:doNotHyphenateCaps`). */
  readonly hyphenateCaps?: boolean;
  /** Hyphenation zone in twips; `undefined` leaves it, `null` removes it. */
  readonly zoneTwips?: number | null;
  /** Limit consecutive hyphens to; 0 is "No limit". */
  readonly consecutiveLimit?: number;
}

export const hyphenationCommand: Command<HyphenationOptions> = {
  id: "layout.hyphenation",
  group: "layout",
  label: "Hyphenation",
  run(model, options) {
    const doc = model.doc;
    if (options.automatic !== undefined)
      setDocumentSettingOnOff(doc, "autoHyphenation", options.automatic);
    if (options.hyphenateCaps !== undefined) {
      setDocumentSettingOnOff(doc, "doNotHyphenateCaps", !options.hyphenateCaps);
    }
    if (options.zoneTwips !== undefined) {
      setDocumentSettingVal(
        doc,
        "hyphenationZone",
        options.zoneTwips === null ? undefined : String(options.zoneTwips),
      );
    }
    if (options.consecutiveLimit !== undefined) {
      const limit = options.consecutiveLimit;
      if (!Number.isInteger(limit) || limit < 0 || limit > 32767) {
        throw new RangeError(`The consecutive hyphen limit must be 0–32767, got ${limit}.`);
      }
      setDocumentSettingVal(doc, "consecutiveHyphenLimit", limit === 0 ? undefined : String(limit));
    }
  },
  isActive: (model) => getDocumentSetting(model.doc, "autoHyphenation").present,
};

/** The paragraph's own `<w:ind>` / `<w:spacing>` attributes, so one box can change without the others. */
function directAttrs(p: WmlParagraph, local: string): Record<string, number> {
  const el = p.pPr && childElementsOf(p.pPr).find((c) => c.name.local === local);
  const out: Record<string, number> = {};
  for (const name of ["left", "right", "firstLine", "hanging", "before", "after", "line"]) {
    const raw = el && getElementAttr(el, name);
    if (raw !== undefined && Number.isInteger(Number(raw))) out[name] = Number(raw);
  }
  return out;
}

/** Layout ▸ Indent Left / Right boxes (twips, applied as direct formatting). */
export const layoutIndentCommand: Command<{ left?: number; right?: number }> = {
  id: "layout.indent",
  group: "layout",
  label: "Indent",
  run(model, change) {
    for (const p of selectedParagraphs(model))
      setParagraphIndent(p, { ...directAttrs(p, "ind"), ...change });
  },
  isEnabled: (model) => selectedParagraphs(model).length > 0,
};

/** Layout ▸ Spacing Before / After boxes (twips). The line spacing rule is kept. */
export const layoutSpacingCommand: Command<{ before?: number; after?: number }> = {
  id: "layout.spacing",
  group: "layout",
  label: "Spacing",
  run(model, change) {
    for (const p of selectedParagraphs(model)) {
      const el = p.pPr && childElementsOf(p.pPr).find((c) => c.name.local === "spacing");
      const rule = el && getElementAttr(el, "lineRule");
      const lineRule = rule === "auto" || rule === "exact" || rule === "atLeast" ? rule : undefined;
      setParagraphSpacing(p, {
        ...directAttrs(p, "spacing"),
        ...(lineRule ? { lineRule } : {}),
        ...change,
      });
    }
  },
  isEnabled: (model) => selectedParagraphs(model).length > 0,
};

/** Document-wide settings the Page Setup dialog edits. */
export interface PageSetupSettings {
  readonly mirrorMargins?: boolean;
  readonly printTwoOnOne?: boolean;
  readonly bookFoldPrinting?: boolean;
  /** Sheets per booklet; 0 is "All". */
  readonly bookFoldPrintingSheets?: number;
  /** Gutter position Top (`w:gutterAtTop`). */
  readonly gutterAtTop?: boolean;
  /** Different odd and even pages (`w:evenAndOddHeaders`). */
  readonly evenAndOddHeaders?: boolean;
}

/** Everything the Page Setup dialog's OK applies. */
export interface PageSetup {
  readonly target: SectionTarget;
  readonly pageSize?: PageSize;
  readonly margins?: PageMargins;
  readonly section?: SectionPropertiesPatch;
  readonly settings?: PageSetupSettings;
}

const SETTING_FLAGS = [
  "mirrorMargins",
  "printTwoOnOne",
  "bookFoldPrinting",
  "gutterAtTop",
  "evenAndOddHeaders",
] as const;

export const pageSetupCommand: Command<PageSetup> = {
  id: "layout.pageSetup",
  group: "layout",
  label: "Page Setup",
  run(model, setup) {
    const scope = sectionScopeFor(model, setup.target, "nextPage");
    if (setup.pageSize) setPageSize(model.doc, setup.pageSize, scope);
    if (setup.margins) setPageMargins(model.doc, setup.margins, scope);
    if (setup.section) setSectionProperties(model.doc, setup.section, scope);
    const settings = setup.settings ?? {};
    for (const flag of SETTING_FLAGS) {
      const value = settings[flag];
      if (value !== undefined) setDocumentSettingOnOff(model.doc, flag, value);
    }
    if (settings.bookFoldPrintingSheets !== undefined) {
      const sheets = settings.bookFoldPrintingSheets;
      setDocumentSettingVal(
        model.doc,
        "bookFoldPrintingSheets",
        sheets === 0 ? undefined : String(sheets),
      );
    }
  },
};

export const layoutCommands = [
  textDirectionCommand,
  columnsCommand,
  insertBreakCommand,
  lineNumbersCommand,
  suppressLineNumbersCommand,
  hyphenationCommand,
  layoutIndentCommand,
  layoutSpacingCommand,
  pageSetupCommand,
];
