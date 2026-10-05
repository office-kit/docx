/**
 * Run-level (character) formatting commands. Each routes through
 * `setRunFormat` / `getRunFormat` from `@office-kit/docx`, applied to exactly
 * the selected characters — {@link applyToSelectionRuns} splits runs at the
 * selection boundaries first, so bolding part of a line bolds only that part.
 */

import {
  childElementsOf,
  clearRunFormat,
  getRunProp,
  type HighlightColor,
  replaceText,
  type RunFormatting,
  setRunFormat,
  setRunOnOff,
  setParagraphValProp,
  setRunValProp,
  type WmlParagraph,
  type WmlRun,
} from "@office-kit/docx";
import { paragraphAt, paragraphsInRange } from "../doc-access.js";
import type { EditorModel } from "../model.js";
import { createStyleResolver, type ResolvedRunFormat } from "../resolve.js";
import { orderSelection } from "../selection.js";
import {
  applyToSelectionRuns,
  overlappingSelectionRunRefs,
  overlappingSelectionRuns,
} from "../selection-runs.js";
import type { Command } from "./types.js";

type ToggleKey = "bold" | "italic" | "strike";
const TOGGLE_ELEMENT: Record<ToggleKey, string> = { bold: "b", italic: "i", strike: "strike" };

function isUnderlined(fmt: ResolvedRunFormat): boolean {
  return fmt.underline !== undefined && fmt.underline !== "none";
}

/**
 * True when every run overlapping the selection *shows* the property, style
 * included — Word presses Bold for a Heading 1 caret even though the run
 * itself carries no `<w:b>`.
 */
function allShow(model: EditorModel, shows: (fmt: ResolvedRunFormat) => boolean): boolean {
  const refs = overlappingSelectionRunRefs(model);
  if (refs.length === 0) return false;
  const styles = createStyleResolver(model.doc);
  return refs.every(({ run, para }) => shows(styles.run(para, run)));
}

function applyToRuns(model: EditorModel, patch: RunFormatting): void {
  // Validate first, with the library's own check, on a throwaway run: the
  // selection's runs are split before the patch is applied, and a bad value
  // must be rejected before that.
  setRunFormat({ kind: "run", pieces: [], extras: [] }, patch);
  applyToSelectionRuns(model, (run) => setRunFormat(run, patch));
}

/**
 * Build a boolean toggle command (bold / italic / strike). Turning a property
 * off drops the direct element, and writes an explicit `w:val="0"` only when
 * the style would otherwise still apply it — what Word writes.
 */
function toggleBool(id: string, key: ToggleKey, label: string): Command<void> {
  const local = TOGGLE_ELEMENT[key];
  return {
    id,
    group: "text",
    label,
    run(model) {
      const next = !allShow(model, (fmt) => fmt[key]);
      const styles = createStyleResolver(model.doc);
      applyToSelectionRuns(model, (run, para) => {
        setRunOnOff(run, local, false);
        if (styles.run(para, run)[key] === next) return;
        if (next) setRunOnOff(run, local, true);
        else setRunValProp(run, local, "0");
      });
    },
    isEnabled: (model) => overlappingSelectionRuns(model).length > 0,
    isActive: (model) => allShow(model, (fmt) => fmt[key]),
  };
}

export const toggleBoldCommand = toggleBool("text.bold", "bold", "Bold");
export const toggleItalicCommand = toggleBool("text.italic", "italic", "Italic");
export const toggleStrikeCommand = toggleBool("text.strike", "strike", "Strikethrough");

export const toggleUnderlineCommand: Command<void> = {
  id: "text.underline",
  group: "text",
  label: "Underline",
  run(model) {
    const next = !allShow(model, isUnderlined);
    const styles = createStyleResolver(model.doc);
    applyToSelectionRuns(model, (run, para) => {
      setRunValProp(run, "u", undefined);
      if (isUnderlined(styles.run(para, run)) === next) return;
      setRunValProp(run, "u", next ? "single" : "none");
    });
  },
  isEnabled: (model) => overlappingSelectionRuns(model).length > 0,
  isActive: (model) => allShow(model, isUnderlined),
};

/** Set a specific underline style (single/double/wave/dotted/thick). */
export const setUnderlineStyleCommand: Command<{
  style: NonNullable<RunFormatting["underline"]>;
}> = {
  id: "text.underlineStyle",
  group: "text",
  label: "Underline style",
  run(model, { style }) {
    applyToRuns(model, { underline: style });
  },
  isEnabled: (model) => overlappingSelectionRuns(model).length > 0,
};

/** Set font family (ASCII/hAnsi). */
export const setFontCommand: Command<{ font: string }> = {
  id: "text.font",
  group: "text",
  label: "Font",
  run(model, { font }) {
    applyToRuns(model, { font });
  },
  isEnabled: (model) => overlappingSelectionRuns(model).length > 0,
};

/** Set East Asian (CJK) font family. */
export const setFontEastAsiaCommand: Command<{ font: string }> = {
  id: "text.fontEastAsia",
  group: "text",
  label: "East Asian font",
  run(model, { font }) {
    applyToRuns(model, { fontEastAsia: font });
  },
  isEnabled: (model) => overlappingSelectionRuns(model).length > 0,
};

/**
 * Set font size in points, rounded to the half-points OOXML stores. A size
 * that is negative or not a finite number throws a `RangeError` before the
 * document changes; 0 is valid OOXML and is written as is.
 */
export const setFontSizeCommand: Command<{ points: number }> = {
  id: "text.fontSize",
  group: "text",
  label: "Font size",
  run(model, { points }) {
    // Checked before `points * 2`, which would coerce "12" to 24 and round
    // -0.1 to 0; the library then checks the resulting half-point integer.
    if (typeof points !== "number" || !Number.isFinite(points) || points < 0) {
      throw new RangeError(
        `Font size must be a finite, non-negative number of points, got ${typeof points === "number" ? points : typeof points}.`,
      );
    }
    applyToRuns(model, { fontSizeHalfPoints: Math.round(points * 2) });
  },
  isEnabled: (model) => overlappingSelectionRuns(model).length > 0,
};

/**
 * Set text color: six hex digits (a leading `#` is dropped) or `"auto"`.
 * Anything else throws a `RangeError` before the document changes.
 */
export const setColorCommand: Command<{ color: string }> = {
  id: "text.color",
  group: "text",
  label: "Font color",
  run(model, { color }) {
    applyToRuns(model, { color: normalizeHex(color) });
  },
  isEnabled: (model) => overlappingSelectionRuns(model).length > 0,
};

/**
 * Set the text highlight to a named {@link HighlightColor}; `"none"` writes an
 * explicit no-highlight. `<w:highlight>` has no RGB form, so a hex or any other
 * value throws a `RangeError` before the document changes (JS callers and UI
 * values are not type-checked).
 */
export const setHighlightCommand: Command<{ color: HighlightColor }> = {
  id: "text.highlight",
  group: "text",
  label: "Highlight",
  run(model, { color }) {
    applyToRuns(model, { highlight: color });
  },
  isEnabled: (model) => overlappingSelectionRuns(model).length > 0,
};

// Word's Clear All Formatting leaves highlighting alone (observed in Word for
// Mac: the highlight survives while every other run property goes).
const KEPT_RUN_PROPERTY = "highlight";
// A section break lives in the last paragraph's pPr; it is layout, not formatting.
const KEPT_PARAGRAPH_PROPERTY = "sectPr";

function clearRunFormatting(run: WmlRun): void {
  const highlight = getRunProp(run, KEPT_RUN_PROPERTY).val;
  clearRunFormat(run);
  if (highlight !== undefined) setRunValProp(run, KEPT_RUN_PROPERTY, highlight);
}

/** Back to the Normal style with no direct paragraph formatting. */
function clearParagraphFormatting(para: WmlParagraph): void {
  if (!para.pPr) return;
  for (const el of childElementsOf(para.pPr)) {
    if (el.name.local !== KEPT_PARAGRAPH_PROPERTY)
      setParagraphValProp(para, el.name.local, undefined);
  }
}

/**
 * Clear All Formatting, as Word does it (checked against Word for Mac):
 * - a caret resets its paragraph to Normal with no direct paragraph
 *   formatting, and leaves the runs alone;
 * - a range drops every character property of the selected text except the
 *   highlight, and resets the paragraphs whose paragraph mark it includes —
 *   every paragraph but the last one, where the range ends before the mark.
 */
export const clearFormattingCommand: Command<void> = {
  id: "text.clearFormat",
  group: "text",
  label: "Clear All Formatting",
  run(model) {
    const sel = model.selection;
    if (!sel) return;
    const ordered = orderSelection(sel);
    if (ordered.collapsed) {
      const para = paragraphAt(model.doc, sel.focus);
      if (para) clearParagraphFormatting(para);
      return;
    }
    const paras = paragraphsInRange(model.doc, ordered);
    applyToSelectionRuns(model, clearRunFormatting);
    for (const para of paras.slice(0, -1)) clearParagraphFormatting(para);
  },
  isEnabled: (model) => model.selection !== null,
};

/**
 * Replace every plain-substring occurrence of `query` in the body (Find &
 * Replace → Replace All). Returns the replacement count. Routed through the
 * command layer so the whole replacement is one undoable step.
 */
export const replaceAllCommand: Command<{ query: string; replacement: string }, number> = {
  id: "text.replaceAll",
  group: "text",
  label: "Replace all",
  run(model, { query, replacement }) {
    if (query === "") throw new Error("Replace All needs a non-empty search string.");
    return replaceText(model.doc, query, replacement);
  },
};

function normalizeHex(color: string): string {
  // Untyped JS callers may pass a non-string; leave it for the library to reject.
  return typeof color === "string" && color.startsWith("#") ? color.slice(1) : color;
}

export const textCommands = [
  toggleBoldCommand,
  toggleItalicCommand,
  toggleStrikeCommand,
  toggleUnderlineCommand,
  setUnderlineStyleCommand,
  setFontCommand,
  setFontEastAsiaCommand,
  setFontSizeCommand,
  setColorCommand,
  setHighlightCommand,
  clearFormattingCommand,
  replaceAllCommand,
];
