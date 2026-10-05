/**
 * Run-level (character) formatting commands — the Home tab's Font group, the
 * Font dialog, Format Painter, Change Case, Phonetic Guide, Enclose
 * Characters, and Find and Replace. Formatting is applied to exactly the
 * selected characters: {@link applyToSelectionRuns} splits runs at the
 * selection boundaries first, so bolding part of a line bolds only that part.
 */

import {
  type BorderLine,
  buildEnclosedCharacterRuns,
  buildRubyRun,
  childElementsOf,
  clearRunFormat,
  type ColorValue,
  type EncloseOptions,
  type FontChoice,
  getElementAttr,
  getRunProp,
  type HighlightColor,
  isolateParagraphRunRange,
  makePropsElement,
  readRuby,
  replaceText,
  type RubyOptions,
  type RunFormatting,
  runTextLength,
  setElementAttr,
  setParagraphValProp,
  setRunBorder,
  setRunColor,
  setRunFont,
  setRunFormat,
  setRunOnOff,
  setRunShading,
  setRunUnderline,
  setRunValProp,
  type ShadingOptions,
  type UnderlineStyle,
  type WmlParagraph,
  type WmlRun,
  type XmlElement,
} from "@office-kit/docx";
import { absoluteOffset } from "../char-offset.js";
import { paragraphAt, paragraphsInRange } from "../doc-access.js";
import {
  bodyParagraphs,
  expandReplacement,
  type FindMatch,
  findMatches,
  type FindOptions,
  replaceRange,
  selectionOf,
} from "../find.js";
import type { EditorModel } from "../model.js";
import { createStyleResolver, type ResolvedRunFormat } from "../resolve.js";
import { caretAt, orderSelection, type Selection } from "../selection.js";
import {
  applyToSelectionRuns,
  overlappingSelectionRunRefs,
  overlappingSelectionRuns,
} from "../selection-runs.js";
import { type CaseMode, changeCase, endsSentence, LENGTH_PRESERVING_MODES } from "../text-case.js";
import {
  applyFontPatch,
  type FontPatch,
  underlineColorOf,
  underlineStyleOf,
} from "./format-patch.js";
import type { Command } from "./types.js";

type ToggleKey = "bold" | "italic" | "strike";
const TOGGLE_ELEMENT: Record<ToggleKey, string> = { bold: "b", italic: "i", strike: "strike" };

function isUnderlined(fmt: ResolvedRunFormat): boolean {
  return fmt.underline !== undefined && fmt.underline !== "none";
}

const hasRuns = (model: EditorModel): boolean => overlappingSelectionRuns(model).length > 0;

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

/**
 * Apply `edit` to the selected runs, after running it once on a throwaway run:
 * the selection's runs are split before the edit, and a value the library
 * rejects must be rejected before that.
 */
function editRuns(model: EditorModel, edit: (run: WmlRun, para: WmlParagraph) => void): void {
  edit({ kind: "run", pieces: [], extras: [] }, { kind: "paragraph", children: [], extras: [] });
  applyToSelectionRuns(model, edit);
}

function applyToRuns(model: EditorModel, patch: RunFormatting): void {
  editRuns(model, (run) => setRunFormat(run, patch));
}

/** What a run shows without its own formatting (its styles only). */
function inheritedFormat(
  styles: ReturnType<typeof createStyleResolver>,
  para: WmlParagraph,
  run: WmlRun,
): ResolvedRunFormat {
  const rStyle = getRunProp(run, "rStyle").val;
  const bare: WmlRun = { kind: "run", pieces: [], extras: [] };
  if (rStyle !== undefined) setRunValProp(bare, "rStyle", rStyle);
  return styles.run(para, bare);
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
    isEnabled: hasRuns,
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
      const color = underlineColorOf(run);
      setRunUnderline(run, undefined);
      if (isUnderlined(styles.run(para, run)) === next) return;
      setRunUnderline(
        run,
        next ? { style: "single", ...(color ? { color } : {}) } : { style: "none" },
      );
    });
  },
  isEnabled: hasRuns,
  isActive: (model) => allShow(model, isUnderlined),
};

/** Set an underline style (any ST_Underline value), keeping the underline colour. */
export const setUnderlineStyleCommand: Command<{ style: UnderlineStyle }> = {
  id: "text.underlineStyle",
  group: "text",
  label: "Underline style",
  run(model, { style }) {
    editRuns(model, (run) => {
      const color = underlineColorOf(run);
      setRunUnderline(run, { style, ...(color ? { color } : {}) });
    });
  },
  isEnabled: hasRuns,
};

/**
 * Underline Color: set the colour of the underline, underlining (single) text
 * that is not underlined yet, as Word does. `undefined` is Automatic.
 */
export const setUnderlineColorCommand: Command<{ color: ColorValue | undefined }> = {
  id: "text.underlineColor",
  group: "text",
  label: "Underline color",
  run(model, { color }) {
    editRuns(model, (run) => {
      const current = underlineStyleOf(run);
      const style = current && current !== "none" ? current : "single";
      setRunUnderline(run, { style, ...(color ? { color } : {}) });
    });
  },
  isEnabled: hasRuns,
};

/**
 * Set the Latin font by name, or the theme's body (`minor`) / headings
 * (`major`) font — the "+Body" / "+Headings" entries of the font list.
 */
export const setFontCommand: Command<{ font: string } | { theme: "major" | "minor" }> = {
  id: "text.font",
  group: "text",
  label: "Font",
  run(model, params) {
    const choice: FontChoice = "theme" in params ? { theme: params.theme } : { name: params.font };
    editRuns(model, (run) => setRunFont(run, "latin", choice));
  },
  isEnabled: hasRuns,
};

/** Set the East Asian (CJK) font. */
export const setFontEastAsiaCommand: Command<{ font: string }> = {
  id: "text.fontEastAsia",
  group: "text",
  label: "East Asian font",
  run(model, { font }) {
    editRuns(model, (run) => setRunFont(run, "eastAsia", { name: font }));
  },
  isEnabled: hasRuns,
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
  isEnabled: hasRuns,
};

/** Word's font size list, in points. */
export const FONT_SIZES: readonly number[] = [
  8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72,
];
// Word's font size box accepts 1–1638 pt.
const MIN_POINTS = 1;
const MAX_POINTS = 1638;
// Past the ends of the list Word steps 10 pt up / 1 pt down.
const LARGE_STEP = 10;
// Word's default size when nothing in the document sets one.
const DEFAULT_HALF_POINTS = 20;

/** The next size in Word's list from `points` (Increase / Decrease Font Size). */
export function nextFontSize(points: number, direction: 1 | -1): number {
  const listed =
    direction > 0 ? FONT_SIZES.find((s) => s > points) : FONT_SIZES.findLast((s) => s < points);
  const fallback =
    direction > 0
      ? Math.floor(points / LARGE_STEP) * LARGE_STEP + LARGE_STEP
      : Math.max(MIN_POINTS, Math.ceil(points) - 1);
  return Math.min(MAX_POINTS, listed ?? fallback);
}

function stepSize(model: EditorModel, next: (points: number) => number): void {
  const styles = createStyleResolver(model.doc);
  applyToSelectionRuns(model, (run, para) => {
    const points = (styles.run(para, run).sizeHalfPoints ?? DEFAULT_HALF_POINTS) / 2;
    const target = Math.min(MAX_POINTS, Math.max(MIN_POINTS, next(points)));
    setRunFormat(run, { fontSizeHalfPoints: Math.round(target * 2) });
  });
}

/** Increase / Decrease Font Size (⌘⇧> / ⌘⇧<): each run moves to the next listed size. */
export const growFontCommand: Command<{ direction: 1 | -1 }> = {
  id: "text.growFont",
  group: "text",
  label: "Change font size",
  run(model, { direction }) {
    stepSize(model, (points) => nextFontSize(points, direction));
  },
  isEnabled: hasRuns,
};

/** Grow / shrink by one point (⌘] / ⌘[). */
export const nudgeFontCommand: Command<{ direction: 1 | -1 }> = {
  id: "text.nudgeFont",
  group: "text",
  label: "Change font size by 1 point",
  run(model, { direction }) {
    stepSize(model, (points) => Math.round(points) + direction);
  },
  isEnabled: hasRuns,
};

/**
 * Set text color: six hex digits (a leading `#` is dropped) or `"auto"`, plus
 * the theme slot / tint / shade when picked from the Theme Colors grid.
 * Anything invalid throws a `RangeError` before the document changes.
 */
export const setColorCommand: Command<
  { color: string } & Partial<Pick<ColorValue, "themeColor" | "themeTint" | "themeShade">>
> = {
  id: "text.color",
  group: "text",
  label: "Font color",
  run(model, { color, themeColor, themeTint, themeShade }) {
    const value: ColorValue = {
      rgb: normalizeHex(color),
      ...(themeColor ? { themeColor } : {}),
      ...(themeTint !== undefined ? { themeTint } : {}),
      ...(themeShade !== undefined ? { themeShade } : {}),
    };
    editRuns(model, (run) => setRunColor(run, value));
  },
  isEnabled: hasRuns,
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
  isEnabled: hasRuns,
};

// Word's Character Shading button: 15 % grey over white (observed in Word).
const CHARACTER_SHADING: ShadingOptions = { pattern: "pct15", color: "auto", fill: "FFFFFF" };
// Word's Character Border button: a ½ pt box with no gap.
const CHARACTER_BORDER: BorderLine = { style: "single", sizeEighths: 4, spacePt: 0 };

function hasShading(fmt: ResolvedRunFormat): boolean {
  const shading = fmt.shading;
  if (!shading || shading.pattern === "nil") return false;
  return shading.pattern !== "clear" || shading.fill !== "auto";
}

function hasBorder(fmt: ResolvedRunFormat): boolean {
  return !!fmt.border && fmt.border.style !== "none" && fmt.border.style !== "nil";
}

/** Character Shading: grey behind the selected text, on / off. */
export const toggleCharacterShadingCommand: Command<void> = {
  id: "text.characterShading",
  group: "text",
  label: "Character Shading",
  run(model) {
    const on = !allShow(model, hasShading);
    applyToSelectionRuns(model, (run) => setRunShading(run, on ? CHARACTER_SHADING : undefined));
  },
  isEnabled: hasRuns,
  isActive: (model) => allShow(model, hasShading),
};

/** Character Border: a box around the selected text, on / off. */
export const toggleCharacterBorderCommand: Command<void> = {
  id: "text.characterBorder",
  group: "text",
  label: "Character Border",
  run(model) {
    const on = !allShow(model, hasBorder);
    applyToSelectionRuns(model, (run) => setRunBorder(run, on ? CHARACTER_BORDER : undefined));
  },
  isEnabled: hasRuns,
  isActive: (model) => allShow(model, hasBorder),
};

/** Shading of the selected text (Borders and Shading ▸ Shading, Apply to: Text). */
export const setRunShadingCommand: Command<{ shading: ShadingOptions | undefined }> = {
  id: "text.shd",
  group: "text",
  label: "Text shading",
  run(model, { shading }) {
    editRuns(model, (run) => setRunShading(run, shading));
  },
  isEnabled: hasRuns,
};

/** Border around the selected text (Borders and Shading ▸ Borders, Apply to: Text). */
export const setRunBorderCommand: Command<{ border: BorderLine | undefined }> = {
  id: "text.bdr",
  group: "text",
  label: "Text border",
  run(model, { border }) {
    editRuns(model, (run) => setRunBorder(run, border));
  },
  isEnabled: hasRuns,
};

/** The Font dialog's OK: apply every field it changed. */
export const fontFormatCommand: Command<FontPatch> = {
  id: "text.fontFormat",
  group: "text",
  label: "Font",
  run(model, patch) {
    // Validate on a throwaway run before splitting the selection's runs.
    applyFontPatch({ kind: "run", pieces: [], extras: [] }, patch);
    const styles = createStyleResolver(model.doc);
    applyToSelectionRuns(model, (run, para) =>
      applyFontPatch(run, patch, inheritedFormat(styles, para, run)),
    );
  },
  isEnabled: hasRuns,
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
 * Reset Character Formatting (⌃Space): drop the selected text's direct
 * character formatting but keep its character style.
 */
export const clearCharacterFormattingCommand: Command<void> = {
  id: "text.clearCharacterFormat",
  group: "text",
  label: "Reset Character Formatting",
  run(model) {
    applyToSelectionRuns(model, (run) => {
      const rStyle = getRunProp(run, "rStyle").val;
      clearRunFormat(run);
      if (rStyle !== undefined) setRunValProp(run, "rStyle", rStyle);
    });
  },
  isEnabled: hasRuns,
};

/**
 * Change Case. Word changes the selection; with a caret it changes the word
 * the caret is in — here, the caret's run.
 */
export const changeCaseCommand: Command<{ mode: CaseMode }> = {
  id: "text.changeCase",
  group: "text",
  label: "Change Case",
  run(model, { mode }) {
    const touched: Array<{ run: WmlRun; para: WmlParagraph }> = [];
    applyToSelectionRuns(model, (run, para) => touched.push({ run, para }));
    let previous: WmlParagraph | undefined;
    let sentenceStart = true;
    for (const { run, para } of touched) {
      if (para !== previous) {
        // The selection's first paragraph starts a sentence only if the text
        // before the selection ends one; later paragraphs always do.
        const before = previous === undefined ? textBefore(para, run) : "";
        sentenceStart = before.trim() === "" || endsSentence(before);
        previous = para;
      }
      for (const piece of run.pieces) {
        if (piece.kind !== "text") continue;
        const next = changeCase(piece.value, mode, sentenceStart);
        if (piece.value.trim() !== "") sentenceStart = endsSentence(piece.value);
        piece.value = next;
      }
    }
    // Width and kana conversions can change the text's length; keep the whole
    // changed text selected.
    if (!LENGTH_PRESERVING_MODES.has(mode)) reselect(model, touched);
  },
  isEnabled: hasRuns,
};

function textBefore(para: WmlParagraph, run: WmlRun): string {
  let out = "";
  for (const child of para.children) {
    if (child === run) break;
    if (child.kind === "run") for (const p of child.pieces) if (p.kind === "text") out += p.value;
  }
  return out;
}

function reselect(model: EditorModel, touched: Array<{ run: WmlRun; para: WmlParagraph }>): void {
  const sel = model.selection;
  const first = touched[0];
  const last = touched.at(-1);
  if (!sel || !first || !last || first.para !== last.para) return;
  const runs = first.para.children.filter((c): c is WmlRun => c.kind === "run");
  const base = orderSelection(sel).start;
  model.setSelection({
    anchor: { ...base, inline: runs.indexOf(first.run), offset: 0 },
    focus: { ...base, inline: runs.indexOf(last.run), offset: runTextLength(last.run) },
  });
}

// ---------------------------------------------------------------------------
// Format Painter

/** Formatting copied by Format Painter. */
export interface CopiedFormat {
  /** The run properties of the first selected character (or of the caret's run). */
  readonly run?: XmlElement;
  /** The paragraph properties, when the source includes paragraph formatting. */
  readonly paragraph?: XmlElement;
}

function cloneElement(el: XmlElement): XmlElement {
  return structuredClone(el);
}

/**
 * Format Painter's copy (⌘⇧C): the character formatting at the start of the
 * selection, and the paragraph formatting too when the selection is a caret
 * or spans paragraphs — as Word copies the paragraph mark's formatting then.
 */
export function copyFormat(model: EditorModel): CopiedFormat | undefined {
  const sel = model.selection;
  if (!sel) return undefined;
  const ordered = orderSelection(sel);
  const ref = overlappingSelectionRunRefs(model)[0];
  const para = paragraphAt(model.doc, ordered.start);
  const wholeParagraph = ordered.collapsed || ordered.start.block !== ordered.end.block;
  // Section breaks are layout, not formatting.
  const pPr = para?.pPr && {
    ...cloneElement(para.pPr),
    children: childElementsOf(cloneElement(para.pPr)).filter((c) => c.name.local !== "sectPr"),
  };
  return {
    ...(ref?.run.rPr ? { run: cloneElement(ref.run.rPr) } : {}),
    ...(wholeParagraph && pPr ? { paragraph: pPr } : {}),
  };
}

/**
 * Format Painter's paste (⌘⇧V): the copied character formatting replaces the
 * selected text's, and copied paragraph formatting replaces the selected
 * paragraphs' (each keeps its own section break).
 */
export const pasteFormatCommand: Command<CopiedFormat> = {
  id: "text.pasteFormat",
  group: "text",
  label: "Format Painter",
  run(model, format) {
    const sel = model.selection;
    if (!sel) return;
    const ordered = orderSelection(sel);
    if (!ordered.collapsed) {
      applyToSelectionRuns(model, (run) => {
        if (format.run) run.rPr = cloneElement(format.run);
        else clearRunFormat(run);
      });
    }
    if (!format.paragraph) return;
    for (const para of paragraphsInRange(model.doc, ordered)) {
      const sectPr = para.pPr && childElementsOf(para.pPr).find((c) => c.name.local === "sectPr");
      const pPr = cloneElement(format.paragraph);
      para.pPr = sectPr ? { ...pPr, children: [...pPr.children, sectPr] } : pPr;
    }
  },
  isEnabled: (model) => model.selection !== null,
};

// ---------------------------------------------------------------------------
// Phonetic Guide, Enclose Characters, Asian layout

/** The selection as one paragraph and a character range in it, if it is that. */
function selectedRange(
  model: EditorModel,
): { para: WmlParagraph; start: number; end: number; base: Selection["anchor"] } | undefined {
  const sel = model.selection;
  if (!sel) return undefined;
  const { start, end, collapsed } = orderSelection(sel);
  const para = paragraphAt(model.doc, start);
  if (collapsed || !para || paragraphAt(model.doc, end) !== para) return undefined;
  return { para, start: absoluteOffset(para, start), end: absoluteOffset(para, end), base: start };
}

/** The selected text and the formatting of its first character. */
function selectedText(
  para: WmlParagraph,
  start: number,
  end: number,
): { text: string; rPr?: XmlElement } {
  const runs = isolateParagraphRunRange(para, start, end);
  const text = runs
    .flatMap((r) => r.pieces)
    .map((p) => (p.kind === "text" ? p.value : ""))
    .join("");
  const rPr = runs[0]?.rPr;
  return rPr ? { text, rPr } : { text };
}

/** Replace the selected runs of `para` with `runs`, and select them. */
function replaceSelection(
  model: EditorModel,
  range: NonNullable<ReturnType<typeof selectedRange>>,
  runs: WmlRun[],
): void {
  const old = new Set<WmlRun>(isolateParagraphRunRange(range.para, range.start, range.end));
  const at = range.para.children.findIndex((c) => c.kind === "run" && old.has(c));
  range.para.children = range.para.children.filter((c) => c.kind !== "run" || !old.has(c));
  range.para.children.splice(at < 0 ? range.para.children.length : at, 0, ...runs);
  const allRuns = range.para.children.filter((c): c is WmlRun => c.kind === "run");
  const first = allRuns.indexOf(runs[0]!);
  const last = allRuns.indexOf(runs.at(-1)!);
  model.setSelection({
    anchor: { ...range.base, inline: first, offset: 0 },
    focus: { ...range.base, inline: last, offset: runTextLength(runs.at(-1)!) },
  });
}

/** The ruby run the caret / selection is on, for editing an existing guide. */
export function selectedRuby(model: EditorModel): ReturnType<typeof readRuby> {
  const ref = overlappingSelectionRunRefs(model)[0];
  return ref ? readRuby(ref.run) : undefined;
}

/**
 * Phonetic Guide: show `ruby` over the selected text (Word's Group mode: one
 * guide for the whole selection). `ruby: ""` removes the guide from a ruby
 * run the caret is on, leaving its base text.
 */
export const phoneticGuideCommand: Command<{ ruby: string; options: RubyOptions }> = {
  id: "text.ruby",
  group: "text",
  label: "Phonetic Guide",
  run(model, { ruby, options }) {
    const ref = overlappingSelectionRunRefs(model)[0];
    const existing = ref && readRuby(ref.run);
    if (ref && existing) {
      const replacement =
        ruby === ""
          ? {
              kind: "run" as const,
              ...(ref.run.rPr ? { rPr: ref.run.rPr } : {}),
              pieces: [{ kind: "text" as const, value: existing.base, preserveSpace: false }],
              extras: [],
            }
          : buildRubyRun(existing.base, ruby, options, ref.run.rPr);
      const index = ref.para.children.indexOf(ref.run);
      ref.para.children[index] = replacement;
      return;
    }
    const range = selectedRange(model);
    if (!range) throw new Error("Select the text to add a phonetic guide to.");
    const { text, rPr } = selectedText(range.para, range.start, range.end);
    replaceSelection(model, range, [buildRubyRun(text, ruby, options, rPr)]);
  },
  isEnabled: (model) => !!selectedRange(model) || !!selectedRuby(model),
};

function sizeOf(model: EditorModel, para: WmlParagraph, rPr: XmlElement | undefined): number {
  const run: WmlRun = { kind: "run", ...(rPr ? { rPr } : {}), pieces: [], extras: [] };
  return createStyleResolver(model.doc).run(para, run).sizeHalfPoints ?? DEFAULT_HALF_POINTS;
}

/** Enclose Characters: draw a circle / square / triangle / diamond around the selected text. */
export const encloseCharactersCommand: Command<Omit<EncloseOptions, "sizeHalfPoints">> = {
  id: "text.enclose",
  group: "text",
  label: "Enclose Characters",
  run(model, options) {
    const range = selectedRange(model);
    if (!range) throw new Error("Select the character to enclose.");
    const { text, rPr } = selectedText(range.para, range.start, range.end);
    const runs = buildEnclosedCharacterRuns(
      text,
      { ...options, sizeHalfPoints: sizeOf(model, range.para, rPr) },
      rPr,
    );
    replaceSelection(model, range, runs);
  },
  isEnabled: (model) => !!selectedRange(model),
};

/** The East Asian layouts Word's Asian Layout menu writes as `<w:eastAsianLayout>`. */
export type EastAsianLayout =
  | {
      readonly kind: "twoLinesInOne";
      readonly brackets?: "none" | "round" | "square" | "angle" | "curly";
    }
  | { readonly kind: "horizontalInVertical"; readonly fitLine: boolean };

/** A fresh `w:id` for an East Asian layout / fit text run group. */
function nextLayoutId(model: EditorModel): string {
  let max = 0;
  for (const { para } of bodyParagraphs(model.doc)) {
    for (const child of para.children) {
      if (child.kind !== "run" || !child.rPr) continue;
      for (const el of childElementsOf(child.rPr)) {
        const id = Number(getElementAttr(el, "id"));
        if ((el.name.local === "eastAsianLayout" || el.name.local === "fitText") && id > max)
          max = id;
      }
    }
  }
  return String(max + 1);
}

/**
 * Asian Layout ▸ Two Lines in One / Horizontal in Vertical
 * (`<w:eastAsianLayout>`, §17.3.2.10); `undefined` removes the layout.
 */
export const eastAsianLayoutCommand: Command<{ layout: EastAsianLayout | undefined }> = {
  id: "text.eastAsianLayout",
  group: "text",
  label: "Asian Layout",
  run(model, { layout }) {
    const id = nextLayoutId(model);
    applyToSelectionRuns(model, (run) => {
      setRunValProp(run, "eastAsianLayout", undefined);
      if (!layout) return;
      const rPr = run.rPr ?? makePropsElement("rPr");
      const el = { ...makePropsElement("eastAsianLayout"), selfClosing: true };
      setWAttr(el, "id", id);
      if (layout.kind === "twoLinesInOne") {
        setWAttr(el, "combine", "1");
        if (layout.brackets && layout.brackets !== "none")
          setWAttr(el, "combineBrackets", layout.brackets);
      } else {
        setWAttr(el, "vert", "1");
        if (layout.fitLine) setWAttr(el, "vertCompress", "1");
      }
      run.rPr = { ...rPr, children: [...rPr.children, el] };
    });
  },
  isEnabled: hasRuns,
};

/**
 * Asian Layout ▸ Fit Text: squeeze or stretch the selected text to `twips`
 * wide (`<w:fitText>`, §17.3.2.14); `undefined` removes it.
 */
export const fitTextCommand: Command<{ twips: number | undefined }> = {
  id: "text.fitText",
  group: "text",
  label: "Fit Text",
  run(model, { twips }) {
    if (twips !== undefined && (!Number.isInteger(twips) || twips <= 0)) {
      throw new RangeError("The text width must be a positive number of twips.");
    }
    const id = nextLayoutId(model);
    applyToSelectionRuns(model, (run) => {
      setRunValProp(run, "fitText", twips === undefined ? undefined : String(twips));
      const el =
        twips !== undefined &&
        run.rPr &&
        childElementsOf(run.rPr).find((c) => c.name.local === "fitText");
      if (el) setWAttr(el, "id", id);
    });
  },
  isEnabled: hasRuns,
};

/** Set a `w:`-namespaced attribute (setElementAttr writes unqualified ones). */
function setWAttr(el: XmlElement, local: string, value: string): void {
  setElementAttr(el, local, undefined);
  (el.attrs as Array<XmlElement["attrs"][number]>).push({
    name: {
      uri: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
      local,
      prefix: "w",
    },
    value,
    isNamespaceDecl: false,
  });
}

// ---------------------------------------------------------------------------
// Find and Replace

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

export interface ReplaceParams {
  readonly find: FindOptions;
  /** The Replace with text (`^&` is the found text, `^t` a tab, `\1` a wildcard group). */
  readonly replacement: string;
  /** Formatting to give the replaced text (the dialog's Format for Replace with). */
  readonly format?: RunFormatting;
}

function sameParagraph(a: FindMatch["at"], b: FindMatch["at"]): boolean {
  return (
    a.block === b.block &&
    a.cell?.row === b.cell?.row &&
    a.cell?.col === b.cell?.col &&
    (a.para ?? 0) === (b.para ?? 0)
  );
}

/** Replace one match; returns the length of the text put in its place. */
function replaceOne(para: WmlParagraph, match: FindMatch, params: ReplaceParams): number {
  const text = expandReplacement(params.replacement, match, !!params.find.wildcards);
  replaceRange(para, match.start, match.end, text, params.format);
  return text.length;
}

/**
 * Advanced Replace All: every match of `find` (case, whole words, wildcards,
 * format filter) is replaced. Returns the number of replacements. Matches are
 * replaced from the end of each paragraph so earlier offsets stay valid.
 */
export const findReplaceAllCommand: Command<ReplaceParams, number> = {
  id: "text.findReplaceAll",
  group: "text",
  label: "Replace All",
  run(model, params) {
    const matches = findMatches(model.doc, params.find);
    const paragraphs = new Map(bodyParagraphs(model.doc).map((p) => [addressKey(p.at), p.para]));
    for (const match of matches.toReversed()) {
      const para = paragraphs.get(addressKey(match.at));
      if (para) replaceOne(para, match, params);
    }
    if (matches.length > 0) model.setSelection(caretAt({ block: 0, inline: 0, offset: 0 }));
    return matches.length;
  },
};

/** The match the selection covers exactly, if any. */
function selectedMatch(model: EditorModel, matches: readonly FindMatch[]): FindMatch | undefined {
  const range = selectedRange(model);
  if (!range) return undefined;
  return matches.find(
    (m) => sameParagraph(m.at, range.base) && m.start === range.start && m.end === range.end,
  );
}

/** The first match after the selection (wrapping around), for Find Next. */
export function nextMatch(
  model: EditorModel,
  matches: readonly FindMatch[],
  direction: 1 | -1,
): FindMatch | undefined {
  if (matches.length === 0) return undefined;
  const sel = model.selection;
  if (!sel) return matches[0];
  const ordered = orderSelection(sel);
  const para = paragraphAt(model.doc, ordered.start);
  const order = bodyParagraphs(model.doc);
  const here = para ? order.findIndex((p) => p.para === para) : 0;
  const offset = para ? absoluteOffset(para, direction > 0 ? ordered.end : ordered.start) : 0;
  const indexOf = new Map(order.map((p, i) => [addressKey(p.at), i]));
  const positions = matches.map((m) => ({ m, p: indexOf.get(addressKey(m.at)) ?? 0 }));
  if (direction > 0) {
    return (
      positions.find(({ m, p }) => p > here || (p === here && m.start >= offset))?.m ?? matches[0]
    );
  }
  return (
    positions.findLast(({ m, p }) => p < here || (p === here && m.end <= offset))?.m ??
    matches.at(-1)
  );
}

/**
 * Replace (one): when the selection is a match, replace it; then select the
 * next match. Returns whether a replacement was made.
 */
export const replaceNextCommand: Command<ReplaceParams, boolean> = {
  id: "text.replaceNext",
  group: "text",
  label: "Replace",
  run(model, params) {
    const current = selectedMatch(model, findMatches(model.doc, params.find));
    const para = current && paragraphAt(model.doc, current.at);
    if (current && para) {
      const length = replaceOne(para, current, params);
      // Continue searching after the replacement, so it is not found again.
      const end = current.start + length;
      const caret = selectionOf(model.doc, { ...current, start: end, end });
      if (caret) model.setSelection(caret);
    }
    const next = nextMatch(model, findMatches(model.doc, params.find), 1);
    const sel = next && selectionOf(model.doc, next);
    if (sel) model.setSelection(sel);
    return current !== undefined;
  },
  isEnabled: (model) => model.selection !== null,
};

function addressKey(at: FindMatch["at"]): string {
  return `${at.block}:${at.cell?.row ?? ""}:${at.cell?.col ?? ""}:${at.para ?? 0}`;
}

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
  setUnderlineColorCommand,
  setFontCommand,
  setFontEastAsiaCommand,
  setFontSizeCommand,
  growFontCommand,
  nudgeFontCommand,
  setColorCommand,
  setHighlightCommand,
  toggleCharacterShadingCommand,
  toggleCharacterBorderCommand,
  setRunShadingCommand,
  setRunBorderCommand,
  fontFormatCommand,
  clearFormattingCommand,
  clearCharacterFormattingCommand,
  changeCaseCommand,
  pasteFormatCommand,
  phoneticGuideCommand,
  encloseCharactersCommand,
  eastAsianLayoutCommand,
  fitTextCommand,
  replaceAllCommand,
  findReplaceAllCommand,
  replaceNextCommand,
];
