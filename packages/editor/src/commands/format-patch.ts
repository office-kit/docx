/**
 * What the Font and Paragraph dialogs (and Modify Style) change, as patches
 * applied to one run / paragraph. A field left `undefined` is not touched, so
 * a dialog that showed a mixed value leaves it mixed.
 *
 * Applied to a run in the document, a toggle that matches what the run's
 * styles already show is written as "no direct formatting" rather than as a
 * redundant `<w:b/>`, which is what Word does; applied to a style
 * (`inherited` omitted) every value is written.
 */

import {
  type BorderLine,
  childElementsOf,
  type ColorValue,
  type FontChoice,
  getElementAttr,
  type ParagraphAlignment,
  setParagraphAlignment,
  setParagraphIndent,
  setParagraphOnOff,
  setParagraphSpacing,
  setParagraphValProp,
  setRunColor,
  setRunFont,
  setRunFormat,
  setRunOnOff,
  setRunUnderline,
  setRunValProp,
  type UnderlineStyle,
  type UnderlineValue,
  type WmlParagraph,
  type WmlRun,
  type XmlElement,
} from "@office-kit/docx";
import type {
  ParagraphToggle,
  ResolvedParagraphFormat,
  ResolvedRunFormat,
  RunToggle,
} from "../resolve.js";

/** The Font dialog's on-off effects, by their `<w:rPr>` element. */
export type FontEffect = "bold" | "italic" | "strike" | RunToggle;

export interface FontPatch {
  readonly font?: FontChoice;
  readonly eastAsiaFont?: FontChoice;
  readonly sizeHalfPoints?: number;
  /** `{ rgb: "auto" }` is Automatic. */
  readonly color?: ColorValue;
  /** `null` removes the underline. */
  readonly underline?: UnderlineValue | null;
  readonly effects?: Partial<Record<FontEffect, boolean>>;
  /** `baseline` removes super/subscript. */
  readonly vertAlign?: "superscript" | "subscript" | "baseline";
  /** Horizontal scale in percent (`w:w`); 100 removes it. */
  readonly scale?: number;
  /** Character spacing in twips (`w:spacing`); 0 removes it. */
  readonly spacing?: number;
  /** Raised / lowered position in half-points (`w:position`); 0 removes it. */
  readonly position?: number;
  /** Kerning from this size in half-points (`w:kern`); `null` turns kerning off. */
  readonly kern?: number | null;
}

const ELEMENT_OF: Readonly<Record<FontEffect, string>> = {
  bold: "b",
  italic: "i",
  strike: "strike",
  dstrike: "dstrike",
  caps: "caps",
  smallCaps: "smallCaps",
  outline: "outline",
  shadow: "shadow",
  emboss: "emboss",
  imprint: "imprint",
  vanish: "vanish",
};

function shows(fmt: ResolvedRunFormat, effect: FontEffect): boolean {
  if (effect === "bold" || effect === "italic" || effect === "strike") return fmt[effect];
  return fmt.toggles.has(effect);
}

/** An integer run property where `neutral` means "remove". */
function setNumber(run: WmlRun, local: string, value: number | undefined, neutral: number): void {
  if (value === undefined) return;
  if (!Number.isInteger(value)) throw new RangeError(`${local} must be a whole number.`);
  setRunValProp(run, local, value === neutral ? undefined : String(value));
}

// Effects that exclude each other in Word's Font dialog: turning one on turns
// the other off.
const EXCLUSIVE: Readonly<Partial<Record<FontEffect, FontEffect>>> = {
  strike: "dstrike",
  dstrike: "strike",
  caps: "smallCaps",
  smallCaps: "caps",
  emboss: "imprint",
  imprint: "emboss",
};

/** Apply a Font dialog patch to a run (or to a style's run properties). */
export function applyFontPatch(run: WmlRun, patch: FontPatch, inherited?: ResolvedRunFormat): void {
  if (patch.font) setRunFont(run, "latin", patch.font);
  if (patch.eastAsiaFont) setRunFont(run, "eastAsia", patch.eastAsiaFont);
  if (patch.sizeHalfPoints !== undefined)
    setRunFormat(run, { fontSizeHalfPoints: patch.sizeHalfPoints });
  if (patch.color) setRunColor(run, patch.color);
  if (patch.underline !== undefined) {
    const inheritedUnderline = inherited?.underline !== undefined && inherited.underline !== "none";
    if (patch.underline === null)
      setRunUnderline(run, inheritedUnderline ? { style: "none" } : undefined);
    else setRunUnderline(run, patch.underline);
  }
  const effects = { ...patch.effects };
  for (const [effect, on] of Object.entries(effects) as Array<[FontEffect, boolean]>) {
    const other = EXCLUSIVE[effect];
    if (on && other && effects[other] === undefined) effects[other] = false;
  }
  for (const [effect, on] of Object.entries(effects) as Array<[FontEffect, boolean | undefined]>) {
    if (on === undefined) continue;
    const local = ELEMENT_OF[effect];
    setRunOnOff(run, local, false);
    if (inherited && shows(inherited, effect) === on) continue;
    if (on) setRunOnOff(run, local, true);
    else setRunValProp(run, local, "0");
  }
  if (patch.vertAlign)
    setRunValProp(run, "vertAlign", patch.vertAlign === "baseline" ? undefined : patch.vertAlign);
  setNumber(run, "w", patch.scale, 100);
  setNumber(run, "spacing", patch.spacing, 0);
  setNumber(run, "position", patch.position, 0);
  if (patch.kern !== undefined)
    setRunValProp(run, "kern", patch.kern === null ? undefined : String(patch.kern));
}

/** The colour stored on a run's `<w:u>`, so changing the line style keeps it. */
export function underlineColorOf(run: WmlRun): ColorValue | undefined {
  const u = run.rPr && childElementsOf(run.rPr).find((c) => c.name.local === "u");
  const rgb = u && getElementAttr(u, "color");
  if (!u || rgb === undefined) return undefined;
  const theme = getElementAttr(u, "themeColor");
  const tint = getElementAttr(u, "themeTint");
  const shade = getElementAttr(u, "themeShade");
  return {
    rgb,
    ...(theme ? { themeColor: theme as NonNullable<ColorValue["themeColor"]> } : {}),
    ...(tint ? { themeTint: Number.parseInt(tint, 16) } : {}),
    ...(shade ? { themeShade: Number.parseInt(shade, 16) } : {}),
  };
}

/** The underline style stored on a run, if any. */
export function underlineStyleOf(run: WmlRun): UnderlineStyle | undefined {
  const u = run.rPr && childElementsOf(run.rPr).find((c) => c.name.local === "u");
  return u ? (getElementAttr(u, "val") as UnderlineStyle | undefined) : undefined;
}

export interface ParagraphPatch {
  readonly alignment?: ParagraphAlignment;
  /** 0–8; `null` is body text. */
  readonly outlineLevel?: number | null;
  readonly indent?: {
    readonly left?: number;
    readonly right?: number;
    readonly firstLine?: number;
    readonly hanging?: number;
  };
  readonly spacing?: {
    readonly before?: number;
    readonly after?: number;
    readonly line?: number;
    readonly lineRule?: "auto" | "exact" | "atLeast";
  };
  readonly toggles?: Partial<Record<ParagraphToggle, boolean>>;
  /** `w:textAlignment` (top / center / baseline / bottom / auto). */
  readonly textAlignment?: string;
}

function direct(pPr: XmlElement | undefined, local: string): XmlElement | undefined {
  return pPr && childElementsOf(pPr).find((c) => c.name.local === local);
}

function intOf(el: XmlElement | undefined, local: string): number | undefined {
  const n = Number(el && getElementAttr(el, local));
  return el && Number.isInteger(n) && getElementAttr(el, local) !== undefined ? n : undefined;
}

/** Only the defined entries, for the library's exact-optional option objects. */
function defined<T extends object>(obj: T): { [K in keyof T]?: Exclude<T[K], undefined> } {
  return Object.fromEntries(Object.entries(obj).filter(([, v]) => v !== undefined)) as {
    [K in keyof T]?: Exclude<T[K], undefined>;
  };
}

/** Apply a Paragraph dialog patch to a paragraph (or a style's paragraph properties). */
export function applyParagraphPatch(
  p: WmlParagraph,
  patch: ParagraphPatch,
  inherited?: ResolvedParagraphFormat,
): void {
  if (patch.alignment) setParagraphAlignment(p, patch.alignment);
  if (patch.outlineLevel !== undefined) {
    setParagraphValProp(
      p,
      "outlineLvl",
      patch.outlineLevel === null ? undefined : String(patch.outlineLevel),
    );
  }
  if (patch.indent) {
    const ind = direct(p.pPr, "ind");
    setParagraphIndent(
      p,
      defined({
        left: patch.indent.left ?? intOf(ind, "left") ?? intOf(ind, "start"),
        right: patch.indent.right ?? intOf(ind, "right") ?? intOf(ind, "end"),
        // A new first-line or hanging value replaces the other.
        firstLine:
          patch.indent.firstLine ??
          (patch.indent.hanging === undefined ? intOf(ind, "firstLine") : undefined),
        hanging:
          patch.indent.hanging ??
          (patch.indent.firstLine === undefined ? intOf(ind, "hanging") : undefined),
      }),
    );
  }
  if (patch.spacing) {
    const spacing = direct(p.pPr, "spacing");
    const lineRule = spacing && getElementAttr(spacing, "lineRule");
    setParagraphSpacing(
      p,
      defined({
        before: patch.spacing.before ?? intOf(spacing, "before"),
        after: patch.spacing.after ?? intOf(spacing, "after"),
        line: patch.spacing.line ?? intOf(spacing, "line"),
        lineRule:
          patch.spacing.lineRule ??
          (lineRule === "auto" || lineRule === "exact" || lineRule === "atLeast"
            ? lineRule
            : undefined),
      }),
    );
  }
  for (const [local, on] of Object.entries(patch.toggles ?? {}) as Array<
    [ParagraphToggle, boolean | undefined]
  >) {
    if (on === undefined) continue;
    setParagraphOnOff(p, local, false);
    if (inherited && inherited.toggles[local] === on) continue;
    if (on) setParagraphOnOff(p, local, true);
    else setParagraphValProp(p, local, "0");
  }
  if (patch.textAlignment !== undefined)
    setParagraphValProp(p, "textAlignment", patch.textAlignment);
}

/** The border Word draws by default from the Borders button: a ½ pt single line. */
export const DEFAULT_BORDER: BorderLine = { style: "single", sizeEighths: 4, spacePt: 1 };
