/**
 * Paragraph-level formatting commands: alignment, indent, spacing, borders,
 * shading, and paragraph style. Each applies to every paragraph the selection
 * spans, via the `@office-kit/docx` paragraph functions.
 */

import {
  childElementsOf,
  getElementAttr,
  type ParagraphAlignment,
  type ParagraphBordersOptions,
  type ParagraphIndent,
  type ParagraphShadingOptions,
  type ParagraphSpacing,
  setParagraphAlignment,
  setParagraphBorders,
  setParagraphIndent,
  setParagraphShading,
  setParagraphSpacing,
  setParagraphStyle,
  type WmlParagraph,
} from "@office-kit/docx";
import { paragraphsInRange } from "../doc-access.js";
import type { EditorModel } from "../model.js";
import { createStyleResolver } from "../resolve.js";
import { orderSelection } from "../selection.js";
import type { Command } from "./types.js";

function selectedParagraphs(model: EditorModel): WmlParagraph[] {
  const sel = model.selection;
  if (!sel) return [];
  return paragraphsInRange(model.doc, orderSelection(sel));
}

export const setAlignmentCommand: Command<{ alignment: ParagraphAlignment }> = {
  id: "paragraph.align",
  group: "paragraph",
  label: "Alignment",
  run(model, { alignment }) {
    for (const p of selectedParagraphs(model)) setParagraphAlignment(p, alignment);
  },
  isEnabled: (model) => selectedParagraphs(model).length > 0,
};

// `start` / `end` are the bidi-aware spellings of left / right (§17.18.44).
const EQUIVALENT_JC: Readonly<Record<string, ParagraphAlignment>> = {
  left: "left",
  start: "left",
  center: "center",
  right: "right",
  end: "right",
  both: "both",
  distribute: "distribute",
};

/** Build a fixed-alignment command (used for the four ribbon buttons). */
function alignTo(id: string, alignment: ParagraphAlignment, label: string): Command<void> {
  return {
    id,
    group: "paragraph",
    label,
    run(model) {
      for (const p of selectedParagraphs(model)) setParagraphAlignment(p, alignment);
    },
    isEnabled: (model) => selectedParagraphs(model).length > 0,
    // Word presses the button for the effective alignment, style included; a
    // paragraph with no `w:jc` anywhere is left-aligned.
    isActive(model) {
      const ps = selectedParagraphs(model);
      const styles = createStyleResolver(model.doc);
      return (
        ps.length > 0 &&
        ps.every(
          (p) => (EQUIVALENT_JC[styles.paragraph(p).alignment ?? "left"] ?? "left") === alignment,
        )
      );
    },
  };
}

export const alignLeftCommand = alignTo("paragraph.alignLeft", "left", "Align left");
export const alignCenterCommand = alignTo("paragraph.alignCenter", "center", "Center");
export const alignRightCommand = alignTo("paragraph.alignRight", "right", "Align right");
export const alignJustifyCommand = alignTo("paragraph.alignJustify", "both", "Justify");

export const setIndentCommand: Command<ParagraphIndent> = {
  id: "paragraph.indent",
  group: "paragraph",
  label: "Indent",
  run(model, indent) {
    for (const p of selectedParagraphs(model)) setParagraphIndent(p, indent);
  },
  isEnabled: (model) => selectedParagraphs(model).length > 0,
};

// Word's Increase/Decrease Indent moves the left indent to the next/previous
// default tab stop, which is 0.5 in (720 twips) unless settings say otherwise.
const INDENT_STEP_TWIPS = 720;
// `w:line` with lineRule="auto" is in 240ths of a line (240 = single spacing).
const AUTO_LINE_UNIT = 240;

/** Only the defined entries, for the library's exact-optional option objects. */
function defined<T extends object>(obj: T): { [K in keyof T]?: Exclude<T[K], undefined> } {
  return Object.fromEntries(Object.entries(obj).filter(([, v]) => v !== undefined)) as {
    [K in keyof T]?: Exclude<T[K], undefined>;
  };
}

/**
 * Increase / Decrease Indent: step the effective left indent (style included,
 * as Word measures it) to the next / previous multiple of 0.5 in, never below
 * 0, and write it as direct formatting. The paragraph's direct right,
 * first-line and hanging indents are kept (`setParagraphIndent` rewrites the
 * whole `<w:ind>`, so they are passed again).
 */
export const indentStepCommand: Command<{ direction: "increase" | "decrease" }> = {
  id: "paragraph.indentStep",
  group: "paragraph",
  label: "Change indent",
  run(model, { direction }) {
    const styles = createStyleResolver(model.doc);
    for (const p of selectedParagraphs(model)) {
      const steps = (styles.paragraph(p).left ?? 0) / INDENT_STEP_TWIPS;
      const left =
        direction === "increase"
          ? (Math.floor(steps) + 1) * INDENT_STEP_TWIPS
          : Math.max(0, (Math.ceil(steps) - 1) * INDENT_STEP_TWIPS);
      setParagraphIndent(p, { ...defined(directIndent(p)), left });
    }
  },
  isEnabled: (model) => selectedParagraphs(model).length > 0,
};

/** The paragraph's own `<w:ind>` attributes (no style), as `ParagraphIndent` fields. */
function directIndent(p: WmlParagraph): ParagraphIndent {
  const ind = p.pPr && childElementsOf(p.pPr).find((c) => c.name.local === "ind");
  if (!ind) return {};
  const int = (local: string): number | undefined => {
    const n = Number(getElementAttr(ind, local));
    return Number.isInteger(n) ? n : undefined;
  };
  return defined({
    right: int("right") ?? int("end"),
    firstLine: int("firstLine"),
    hanging: int("hanging"),
  });
}

/** The paragraph's own `before` / `after` spacing (no style). */
function directSpacing(p: WmlParagraph): ParagraphSpacing {
  const spacing = p.pPr && childElementsOf(p.pPr).find((c) => c.name.local === "spacing");
  if (!spacing) return {};
  const int = (local: string): number | undefined => {
    const n = Number(getElementAttr(spacing, local));
    return Number.isInteger(n) ? n : undefined;
  };
  return defined({ before: int("before"), after: int("after") });
}

/**
 * Line spacing as a multiple of single spacing (Word's 1.0 / 1.15 / 1.5 …
 * menu). The paragraph's direct space before/after is kept.
 */
export const setLineSpacingCommand: Command<{ multiple: number }> = {
  id: "paragraph.lineSpacing",
  group: "paragraph",
  label: "Line spacing",
  run(model, { multiple }) {
    if (typeof multiple !== "number" || !Number.isFinite(multiple) || multiple <= 0) {
      throw new RangeError(`Line spacing must be a positive number, got ${String(multiple)}.`);
    }
    for (const p of selectedParagraphs(model)) {
      setParagraphSpacing(p, {
        ...directSpacing(p),
        line: Math.round(multiple * AUTO_LINE_UNIT),
        lineRule: "auto",
      });
    }
  },
  isEnabled: (model) => selectedParagraphs(model).length > 0,
};

/**
 * The selection's effective line spacing as a multiple, when every paragraph
 * agrees and uses auto (proportional) spacing; otherwise `undefined`.
 */
export function lineSpacingOf(model: EditorModel): number | undefined {
  const styles = createStyleResolver(model.doc);
  const values = selectedParagraphs(model).map((p) => {
    const fmt = styles.paragraph(p);
    return fmt.lineRule === "auto" && fmt.line !== undefined
      ? fmt.line / AUTO_LINE_UNIT
      : undefined;
  });
  const first = values[0];
  return first !== undefined && values.every((v) => v === first) ? first : undefined;
}

export const setSpacingCommand: Command<ParagraphSpacing> = {
  id: "paragraph.spacing",
  group: "paragraph",
  label: "Line & paragraph spacing",
  run(model, spacing) {
    for (const p of selectedParagraphs(model)) setParagraphSpacing(p, spacing);
  },
  isEnabled: (model) => selectedParagraphs(model).length > 0,
};

export const setParagraphBordersCommand: Command<ParagraphBordersOptions> = {
  id: "paragraph.borders",
  group: "paragraph",
  label: "Borders",
  run(model, borders) {
    for (const p of selectedParagraphs(model)) setParagraphBorders(p, borders);
  },
  isEnabled: (model) => selectedParagraphs(model).length > 0,
};

export const setParagraphShadingCommand: Command<ParagraphShadingOptions> = {
  id: "paragraph.shading",
  group: "paragraph",
  label: "Shading",
  run(model, shading) {
    for (const p of selectedParagraphs(model)) setParagraphShading(p, shading);
  },
  isEnabled: (model) => selectedParagraphs(model).length > 0,
};

export const setParagraphStyleCommand: Command<{ styleId: string | undefined }> = {
  id: "paragraph.style",
  group: "paragraph",
  label: "Paragraph style",
  run(model, { styleId }) {
    for (const p of selectedParagraphs(model)) setParagraphStyle(p, styleId);
  },
  isEnabled: (model) => selectedParagraphs(model).length > 0,
};

export const paragraphCommands = [
  setAlignmentCommand,
  alignLeftCommand,
  alignCenterCommand,
  alignRightCommand,
  alignJustifyCommand,
  setIndentCommand,
  indentStepCommand,
  setSpacingCommand,
  setLineSpacingCommand,
  setParagraphBordersCommand,
  setParagraphShadingCommand,
  setParagraphStyleCommand,
];
