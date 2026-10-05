/**
 * Paragraph-level formatting commands: alignment, indent, spacing, borders,
 * shading, and paragraph style. Each applies to every paragraph the selection
 * spans, via the `@office-kit/docx` paragraph functions.
 */

import {
  applyListToParagraph,
  type BorderLine,
  childElementsOf,
  type ColorValue,
  getElementAttr,
  getParagraphNumbering,
  getParagraphStyle,
  makePropsElement,
  type ParagraphAlignment,
  type ParagraphBordersOptions,
  type ParagraphBorderSide,
  type ParagraphIndent,
  type ParagraphShadingOptions,
  type ParagraphSpacing,
  paragraphText,
  setParagraphAlignment,
  setParagraphBorder,
  setParagraphBorders,
  setParagraphIndent,
  setParagraphShading,
  setParagraphSpacing,
  setParagraphStyle,
  setParagraphTabs,
  setParagraphValProp,
  type TabStop,
  type WmlParagraph,
  type XmlAttr,
  type XmlElement,
} from "@office-kit/docx";
import { paragraphsInRange } from "../doc-access.js";
import type { EditorModel } from "../model.js";
import { createStyleResolver, type StyleResolver } from "../resolve.js";
import { orderSelection } from "../selection.js";
import { applyParagraphPatch, DEFAULT_BORDER, type ParagraphPatch } from "./format-patch.js";
import type { Command } from "./types.js";

// ST_DecimalNumber for w:ilvl: list levels 0–8.
const MAX_LIST_LEVEL = 8;

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
      // On a list item Word changes the list level instead.
      const list = getParagraphNumbering(p);
      if (list) {
        const ilvl = Math.min(
          MAX_LIST_LEVEL,
          Math.max(0, list.ilvl + (direction === "increase" ? 1 : -1)),
        );
        applyListToParagraph(model.doc, p, list.numId, ilvl);
        continue;
      }
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
  const lineRule = getElementAttr(spacing, "lineRule");
  return defined({
    before: int("before"),
    after: int("after"),
    line: int("line"),
    lineRule:
      lineRule === "auto" || lineRule === "exact" || lineRule === "atLeast" ? lineRule : undefined,
  });
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
      const { before, after } = directSpacing(p);
      setParagraphSpacing(p, {
        ...defined({ before, after }),
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

export const alignDistributedCommand = alignTo(
  "paragraph.alignDistributed",
  "distribute",
  "Distributed",
);

/** The Paragraph dialog's OK: apply every field it changed. */
export const paragraphFormatCommand: Command<ParagraphPatch> = {
  id: "paragraph.format",
  group: "paragraph",
  label: "Paragraph",
  run(model, patch) {
    const styles = createStyleResolver(model.doc);
    for (const p of selectedParagraphs(model)) {
      const styleOnly: WmlParagraph = { kind: "paragraph", children: [], extras: [] };
      const styleId = getParagraphStyle(p);
      if (styleId !== undefined) setParagraphStyle(styleOnly, styleId);
      applyParagraphPatch(p, patch, styles.paragraph(styleOnly));
    }
  },
  isEnabled: (model) => selectedParagraphs(model).length > 0,
};

// Word's Add / Remove Space Before / After Paragraph use 12 pt.
const ADDED_SPACE_TWIPS = 240;

/**
 * Line and Paragraph Spacing ▸ Add / Remove Space Before (After) Paragraph:
 * the menu offers "Add" when the first selected paragraph has no space there.
 */
export const toggleParagraphSpaceCommand: Command<{ side: "before" | "after" }> = {
  id: "paragraph.toggleSpace",
  group: "paragraph",
  label: "Space before / after paragraph",
  run(model, { side }) {
    const paras = selectedParagraphs(model);
    const styles = createStyleResolver(model.doc);
    const first = paras[0];
    const add = !!first && (styles.paragraph(first)[side] ?? 0) === 0;
    for (const p of paras) {
      setParagraphSpacing(p, { ...directSpacing(p), [side]: add ? ADDED_SPACE_TWIPS : 0 });
    }
  },
  isEnabled: (model) => selectedParagraphs(model).length > 0,
};

/** Whether the first selected paragraph has space on `side` (labels the menu items). */
export function hasParagraphSpace(model: EditorModel, side: "before" | "after"): boolean {
  const first = selectedParagraphs(model)[0];
  return !!first && (createStyleResolver(model.doc).paragraph(first)[side] ?? 0) > 0;
}

/** Shading ▸ a colour from the palette (paragraph shading); `undefined` is No Color. */
export const paragraphShadingColorCommand: Command<{ fill: ColorValue | undefined }> = {
  id: "paragraph.shadingColor",
  group: "paragraph",
  label: "Shading",
  run(model, { fill }) {
    for (const p of selectedParagraphs(model)) {
      if (!fill) {
        setParagraphValProp(p, "shd", undefined);
        continue;
      }
      setParagraphShading(p, {
        pattern: "clear",
        color: "auto",
        fill: fill.rgb,
        ...(fill.themeColor ? { themeFill: fill.themeColor } : {}),
        ...(fill.themeTint !== undefined ? { themeFillTint: fill.themeTint } : {}),
        ...(fill.themeShade !== undefined ? { themeFillShade: fill.themeShade } : {}),
      });
    }
  },
  isEnabled: (model) => selectedParagraphs(model).length > 0,
};

/** The entries of Word's Borders menu that apply to paragraphs. */
export type BorderPreset =
  | "bottom"
  | "top"
  | "left"
  | "right"
  | "none"
  | "all"
  | "outside"
  | "inside"
  | "insideHorizontal";

const PRESET_SIDES: Readonly<Record<BorderPreset, readonly ParagraphBorderSide[]>> = {
  bottom: ["bottom"],
  top: ["top"],
  left: ["left"],
  right: ["right"],
  none: [],
  all: ["top", "left", "bottom", "right", "between"],
  outside: ["top", "left", "bottom", "right"],
  // For paragraphs, Inside Borders and Inside Horizontal are the line drawn
  // between paragraphs that share the box.
  inside: ["between"],
  insideHorizontal: ["between"],
};
const ALL_SIDES: readonly ParagraphBorderSide[] = [
  "top",
  "left",
  "bottom",
  "right",
  "between",
  "bar",
];

function hasSide(styles: StyleResolver, p: WmlParagraph, side: ParagraphBorderSide): boolean {
  const border = styles.paragraph(p).borders[side];
  return !!border && border.style !== "none" && border.style !== "nil";
}

/** Whether every selected paragraph shows all of a preset's sides (the menu's check mark). */
export function bordersPresetActive(model: EditorModel, preset: BorderPreset): boolean {
  const paras = selectedParagraphs(model);
  const styles = createStyleResolver(model.doc);
  const sides = PRESET_SIDES[preset];
  if (paras.length === 0) return false;
  if (sides.length === 0) return paras.every((p) => ALL_SIDES.every((s) => !hasSide(styles, p, s)));
  return paras.every((p) => sides.every((s) => hasSide(styles, p, s)));
}

/**
 * The Borders button / menu: toggle the preset's sides on the selected
 * paragraphs (with `line`, Word's last-used border, ½ pt single by default).
 * No Border removes every side. Paragraphs with the same borders are drawn
 * as one box, so a Bottom Border on three paragraphs underlines the group.
 */
export const bordersPresetCommand: Command<{ preset: BorderPreset; line?: BorderLine }> = {
  id: "paragraph.bordersPreset",
  group: "paragraph",
  label: "Borders",
  run(model, { preset, line = DEFAULT_BORDER }) {
    const paras = selectedParagraphs(model);
    if (preset === "none") {
      for (const p of paras) setParagraphValProp(p, "pBdr", undefined);
      return;
    }
    const off = bordersPresetActive(model, preset);
    for (const p of paras) {
      for (const side of PRESET_SIDES[preset]) setParagraphBorder(p, side, off ? undefined : line);
    }
  },
  isEnabled: (model) => selectedParagraphs(model).length > 0,
};

/** Borders and Shading ▸ Borders (Apply to: Paragraph): set each side. */
export const paragraphBordersCommand: Command<{
  sides: Partial<Record<ParagraphBorderSide, BorderLine | undefined>>;
}> = {
  id: "paragraph.bdr",
  group: "paragraph",
  label: "Paragraph borders",
  run(model, { sides }) {
    for (const p of selectedParagraphs(model)) {
      for (const [side, line] of Object.entries(sides) as Array<
        [ParagraphBorderSide, BorderLine | undefined]
      >) {
        setParagraphBorder(p, side, line);
      }
    }
  },
  isEnabled: (model) => selectedParagraphs(model).length > 0,
};

/** The Tabs dialog's OK: the selected paragraphs' custom tab stops. */
export const setTabsCommand: Command<{ tabs: readonly TabStop[] }> = {
  id: "paragraph.tabs",
  group: "paragraph",
  label: "Tabs",
  run(model, { tabs }) {
    for (const p of selectedParagraphs(model)) setParagraphTabs(p, tabs);
  },
  isEnabled: (model) => selectedParagraphs(model).length > 0,
};

/** Sort Text: what to compare and in which order. */
export interface SortOptions {
  readonly by: "text" | "number" | "date";
  readonly order: "ascending" | "descending";
  /** Leave the first selected paragraph (a header row) in place. */
  readonly header?: boolean;
  readonly matchCase?: boolean;
}

const NUMBER_IN_TEXT = /-?\d+(?:[.,]\d+)?/;

function sortKey(text: string, by: SortOptions["by"]): number | string {
  if (by === "number") {
    const m = NUMBER_IN_TEXT.exec(text);
    return m ? Number(m[0].replace(",", ".")) : Number.NEGATIVE_INFINITY;
  }
  if (by === "date") {
    const time = Date.parse(text.trim());
    return Number.isNaN(time) ? Number.NEGATIVE_INFINITY : time;
  }
  return text;
}

/**
 * Sort Text: reorder the selected top-level paragraphs by their text, by the
 * first number in them, or by the date they hold. Paragraphs Word cannot
 * read a number or date from sort first, as in Word. The sort is stable.
 */
export const sortParagraphsCommand: Command<SortOptions> = {
  id: "paragraph.sort",
  group: "paragraph",
  label: "Sort",
  run(model, options) {
    const sel = model.selection;
    if (!sel) return;
    const { start, end } = orderSelection(sel);
    if (start.cell || end.cell) throw new Error("Sort Text sorts paragraphs outside tables.");
    const blocks = model.doc.document.body.blocks;
    const first = start.block + (options.header ? 1 : 0);
    const range = blocks.slice(first, end.block + 1);
    if (range.some((b) => b.kind !== "paragraph"))
      throw new Error("Select paragraphs only to sort them.");
    const collator = new Intl.Collator(undefined, {
      numeric: false,
      sensitivity: options.matchCase ? "variant" : "base",
    });
    const keyed = range.map((block, index) => ({
      block,
      index,
      key: sortKey(block.kind === "paragraph" ? paragraphText(block) : "", options.by),
    }));
    const sign = options.order === "ascending" ? 1 : -1;
    keyed.sort((a, b) => {
      const cmp =
        typeof a.key === "number" && typeof b.key === "number"
          ? a.key - b.key
          : collator.compare(String(a.key), String(b.key));
      return cmp * sign || a.index - b.index;
    });
    blocks.splice(first, range.length, ...keyed.map((k) => k.block));
    model.setSelection({
      anchor: { block: start.block, inline: 0, offset: 0 },
      focus: { block: end.block },
    });
  },
  isEnabled: (model) => selectedParagraphs(model).length > 0,
};

const VML_NS = "urn:schemas-microsoft-com:vml";
const OFFICE_NS = "urn:schemas-microsoft-com:office:office";

function qAttr(uri: string, prefix: string, local: string, value: string): XmlAttr {
  return { name: { uri, local, prefix }, value, isNamespaceDecl: false };
}

function nsDecl(prefix: string, uri: string): XmlAttr {
  return {
    name: { uri: "http://www.w3.org/2000/xmlns/", local: prefix, prefix: "xmlns" },
    value: uri,
    isNamespaceDecl: true,
  };
}

/**
 * Borders ▸ Horizontal Line: a paragraph holding Word's horizontal-rule shape
 * (a VML `v:rect` with `o:hr`, ECMA-376 Part 4), inserted after the caret's
 * paragraph.
 */
export const insertHorizontalLineCommand: Command<void> = {
  id: "paragraph.horizontalLine",
  group: "paragraph",
  label: "Horizontal Line",
  run(model) {
    const sel = model.selection;
    if (!sel) return;
    const at = orderSelection(sel).end;
    if (at.cell) throw new Error("Horizontal Line is inserted between body paragraphs.");
    const rect: XmlElement = {
      kind: "element",
      name: { uri: VML_NS, local: "rect", prefix: "v" },
      attrs: [
        nsDecl("v", VML_NS),
        nsDecl("o", OFFICE_NS),
        qAttr("", "", "style", "width:0;height:1.5pt"),
        qAttr(OFFICE_NS, "o", "hralign", "center"),
        qAttr(OFFICE_NS, "o", "hrstd", "t"),
        qAttr(OFFICE_NS, "o", "hr", "t"),
        qAttr("", "", "fillcolor", "#a0a0a0"),
        qAttr("", "", "stroked", "f"),
      ],
      children: [],
      xmlSpace: "default",
      selfClosing: true,
    };
    const pict = makePropsElement("pict");
    (pict.children as XmlElement[]).push(rect);
    const paragraph: WmlParagraph = {
      kind: "paragraph",
      children: [{ kind: "run", pieces: [{ kind: "pict", node: pict }], extras: [] }],
      extras: [],
    };
    model.doc.document.body.blocks.splice(at.block + 1, 0, paragraph);
    model.setSelection({ anchor: { block: at.block + 1 }, focus: { block: at.block + 1 } });
  },
  isEnabled: (model) => !!model.selection && !orderSelection(model.selection).end.cell,
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
  alignDistributedCommand,
  paragraphFormatCommand,
  toggleParagraphSpaceCommand,
  paragraphShadingColorCommand,
  bordersPresetCommand,
  paragraphBordersCommand,
  setTabsCommand,
  sortParagraphsCommand,
  insertHorizontalLineCommand,
];
