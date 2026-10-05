/**
 * Style commands: apply a style from the gallery or the Styles pane (creating
 * Word's built-in definition the first time a latent style is used), and the
 * New Style / Modify Style / Delete of the Styles pane.
 */

import {
  addStyle,
  type BuildStyleOptions,
  builtinStyles,
  ensureBuiltinStyle,
  ensureHeadingStyles,
  getElementAttr,
  getParagraphStyle,
  getRunProp,
  removeStyle,
  setParagraphStyle,
  setRunValProp,
  setStyleOnOff,
  setStyleValProp,
  stylesPart,
  updateStyleFormatting,
  type WmlParagraph,
  type XmlElement,
} from "@office-kit/docx";
import { absoluteOffset } from "../char-offset.js";
import { paragraphAt, paragraphsInRange } from "../doc-access.js";
import { bodyParagraphs } from "../find.js";
import type { EditorModel } from "../model.js";
import { orderSelection } from "../selection.js";
import { applyToSelectionRuns } from "../selection-runs.js";
import {
  applyFontPatch,
  applyParagraphPatch,
  type FontPatch,
  type ParagraphPatch,
} from "./format-patch.js";
import type { Command } from "./types.js";

/** Options for a new style. */
export type AddStyleParams = BuildStyleOptions;

export const addStyleCommand: Command<AddStyleParams> = {
  id: "style.add",
  group: "style",
  label: "New style",
  run(model, options) {
    addStyle(model.doc, options);
  },
};

export const ensureHeadingStylesCommand: Command<void> = {
  id: "style.ensureHeadings",
  group: "style",
  label: "Ensure heading styles",
  run(model) {
    ensureHeadingStyles(model.doc);
  },
};

function styleElement(model: EditorModel, styleId: string): XmlElement | undefined {
  return stylesPart(model.doc)?.styles.find((s) => getElementAttr(s, "styleId") === styleId);
}

function child(el: XmlElement | undefined, local: string): XmlElement | undefined {
  return el?.children.find((c): c is XmlElement => c.kind === "element" && c.name.local === local);
}

/** The paragraph style a paragraph without `w:pStyle` has (`Normal`). */
export function defaultParagraphStyleId(model: EditorModel): string | undefined {
  const style = stylesPart(model.doc)?.styles.find(
    (s) => getElementAttr(s, "type") === "paragraph" && getElementAttr(s, "default") === "1",
  );
  return style && getElementAttr(style, "styleId");
}

/**
 * Whether the selection is part of one paragraph, not all of it: Word then
 * applies a linked style's character half to the selected text only.
 */
function partialParagraph(model: EditorModel): boolean {
  const sel = model.selection;
  if (!sel) return false;
  const { start, end, collapsed } = orderSelection(sel);
  const para = paragraphAt(model.doc, start);
  if (collapsed || !para || paragraphAt(model.doc, end) !== para) return false;
  const length = absoluteOffset(para, { ...end, inline: Number.MAX_SAFE_INTEGER });
  return absoluteOffset(para, start) > 0 || absoluteOffset(para, end) < length;
}

/**
 * Apply a style by id, as clicking it in the gallery or the Styles pane does.
 * A built-in style the document does not define yet is added with Word's
 * definition first. A paragraph style applies to the selected paragraphs —
 * or, for a linked style and a selection inside one paragraph, its character
 * style to the selected text; a character style applies to the selected text.
 */
export const applyStyleCommand: Command<{ styleId: string }> = {
  id: "style.apply",
  group: "style",
  label: "Apply style",
  run(model, { styleId }) {
    ensureBuiltinStyle(model.doc, styleId);
    const style = styleElement(model, styleId);
    if (!style) throw new Error(`The document has no style "${styleId}".`);
    const type = getElementAttr(style, "type");
    const link = child(style, "link");
    const linked = link && getElementAttr(link, "val");
    if (type === "character" || (type === "paragraph" && linked && partialParagraph(model))) {
      const charStyle = type === "character" ? styleId : (linked ?? styleId);
      applyToSelectionRuns(model, (run) => setRunValProp(run, "rStyle", charStyle));
      return;
    }
    if (type !== "paragraph")
      throw new Error(`"${styleId}" is not a paragraph or character style.`);
    const sel = model.selection;
    if (!sel) return;
    const normal = defaultParagraphStyleId(model);
    for (const p of paragraphsInRange(model.doc, orderSelection(sel))) {
      // Word writes no pStyle for the default style.
      setParagraphStyle(p, styleId === normal ? undefined : styleId);
    }
  },
  isEnabled: (model) => model.selection !== null,
};

/** The Modify Style / New Style dialog's fields. */
export interface StyleDefinition {
  readonly name: string;
  readonly basedOn?: string | undefined;
  /** Style for following paragraph (paragraph styles only). */
  readonly next?: string | undefined;
  /** Add to the Styles gallery (`w:qFormat`). */
  readonly quickStyle?: boolean;
  /** Automatically update (`w:autoRedefine`). */
  readonly autoUpdate?: boolean;
  readonly font?: FontPatch;
  readonly paragraph?: ParagraphPatch;
}

const NON_ID_CHARS = /[^A-Za-z0-9]/g;

/** A style id for a new style name: its letters and digits, made unique. */
function newStyleId(model: EditorModel, name: string): string {
  const base = name.replace(NON_ID_CHARS, "") || "Style";
  const taken = new Set(
    (stylesPart(model.doc)?.styles ?? []).map((s) => getElementAttr(s, "styleId")),
  );
  let id = base;
  for (let n = 1; taken.has(id); n++) id = `${base}${n}`;
  return id;
}

function writeDefinition(model: EditorModel, styleId: string, def: StyleDefinition): void {
  setStyleValProp(model.doc, styleId, "name", def.name);
  setStyleValProp(model.doc, styleId, "basedOn", def.basedOn);
  if (def.next !== undefined) setStyleValProp(model.doc, styleId, "next", def.next || undefined);
  if (def.quickStyle !== undefined) setStyleOnOff(model.doc, styleId, "qFormat", def.quickStyle);
  if (def.autoUpdate !== undefined)
    setStyleOnOff(model.doc, styleId, "autoRedefine", def.autoUpdate);
  updateStyleFormatting(model.doc, styleId, {
    ...(def.font ? { run: (run) => applyFontPatch(run, def.font ?? {}) } : {}),
    ...(def.paragraph ? { paragraph: (p) => applyParagraphPatch(p, def.paragraph ?? {}) } : {}),
  });
}

function nameTaken(model: EditorModel, name: string, except?: string): boolean {
  return (stylesPart(model.doc)?.styles ?? []).some((s) => {
    const n = child(s, "name");
    return (
      getElementAttr(s, "styleId") !== except &&
      n !== undefined &&
      getElementAttr(n, "val") === name
    );
  });
}

/**
 * New Style: define a paragraph or character style and apply it to the
 * selection, as Word's New Style dialog does. Returns the new style id.
 */
export const newStyleCommand: Command<
  StyleDefinition & { type: "paragraph" | "character" },
  string
> = {
  id: "style.new",
  group: "style",
  label: "New Style",
  run(model, def) {
    if (def.name.trim() === "") throw new Error("A style needs a name.");
    if (nameTaken(model, def.name)) throw new Error(`A style named "${def.name}" already exists.`);
    const styleId = newStyleId(model, def.name);
    addStyle(model.doc, { type: def.type, styleId, customStyle: true });
    writeDefinition(model, styleId, def);
    applyStyleCommand.run(model, { styleId });
    return styleId;
  },
};

/** Modify Style: change a style's definition (and so every paragraph using it). */
export const modifyStyleCommand: Command<StyleDefinition & { styleId: string }> = {
  id: "style.modify",
  group: "style",
  label: "Modify Style",
  run(model, { styleId, ...def }) {
    ensureBuiltinStyle(model.doc, styleId);
    if (!styleElement(model, styleId)) throw new Error(`The document has no style "${styleId}".`);
    if (nameTaken(model, def.name, styleId))
      throw new Error(`A style named "${def.name}" already exists.`);
    writeDefinition(model, styleId, def);
  },
};

/** Styles Word does not let you delete. */
export function isDeletableStyle(model: EditorModel, styleId: string): boolean {
  const style = styleElement(model, styleId);
  if (!style || getElementAttr(style, "default") === "1") return false;
  return !builtinStyles().some((s) => s.styleId === styleId);
}

/**
 * Delete a style: the paragraphs and text that used it go back to Normal /
 * Default Paragraph Font (as in Word). A linked character style goes too.
 */
export const deleteStyleCommand: Command<{ styleId: string }> = {
  id: "style.delete",
  group: "style",
  label: "Delete Style",
  run(model, { styleId }) {
    if (!isDeletableStyle(model, styleId)) throw new Error(`"${styleId}" cannot be deleted.`);
    const link = child(styleElement(model, styleId), "link");
    const linked = link && getElementAttr(link, "val");
    const gone = new Set([styleId, ...(linked ? [linked] : [])]);
    for (const id of gone) removeStyle(model.doc, id);
    for (const { para } of bodyParagraphs(model.doc)) clearStyleRefs(para, gone);
  },
};

function clearStyleRefs(para: WmlParagraph, gone: ReadonlySet<string>): void {
  const pStyle = getParagraphStyle(para);
  if (pStyle !== undefined && gone.has(pStyle)) setParagraphStyle(para, undefined);
  for (const run of para.children) {
    if (run.kind !== "run") continue;
    const rStyle = getRunProp(run, "rStyle").val;
    if (rStyle !== undefined && gone.has(rStyle)) setRunValProp(run, "rStyle", undefined);
  }
}

export const styleCommands = [
  addStyleCommand,
  ensureHeadingStylesCommand,
  applyStyleCommand,
  newStyleCommand,
  modifyStyleCommand,
  deleteStyleCommand,
];
