/**
 * State the Home tab, its dialogs, the Styles pane and the Navigation pane
 * share for one editor session: the colours the split buttons apply, recent
 * fonts, Format Painter, Find options, and which style a dialog edits.
 */

import {
  type ColorValue,
  getParagraphStyle,
  type HighlightColor,
  type UnderlineStyle,
  type WmlRun,
} from "@office-kit/docx";
import {
  createStyleResolver,
  hasEastAsianText,
  orderSelection,
  paragraphAt,
  paragraphsInRange,
  runAtPath,
  runsInRange,
  type EditorModel,
  type FindFormat,
  type ResolvedParagraphFormat,
  type ResolvedRunFormat,
  type commands,
} from "@office-kit/docx-editor";
import type { EditorSession } from "../../../session.svelte";

/** Word keeps the last ten fonts used in the Recently Used Fonts section. */
const RECENT_FONTS = 10;

export class HomeState {
  /** What the Font Color face applies (Word starts with red). */
  fontColorPen = $state<ColorValue>({ rgb: "C00000" });
  highlightPen = $state<HighlightColor>("yellow");
  /** What the Shading face applies; `undefined` is No Color. */
  shadingPen = $state<ColorValue | undefined>(undefined);
  underlinePen = $state<UnderlineStyle>("single");
  recentFonts = $state<string[]>([]);
  /** Format Painter: armed with copied formatting; `sticky` after a double-click. */
  painter = $state<{ format: commands.CopiedFormat; sticky: boolean } | null>(null);
  /** Formatting copied with ⌘⇧C, pasted with ⌘⇧V. */
  copiedFormat = $state<commands.CopiedFormat | null>(null);
  /** Who receives the More Colors dialog's choice. */
  moreColors = $state<{ initial: string; pick: (color: ColorValue) => void } | null>(null);
  /** The style the Modify Style dialog edits; `null` is New Style. */
  styleTarget = $state<string | null>(null);
  find = $state({
    matchCase: false,
    wholeWords: false,
    wildcards: false,
    replace: "",
    format: {} as FindFormat,
    replaceFormat: {} as commands.ReplaceParams["format"] & object,
  });

  useFont(name: string): void {
    this.recentFonts = [name, ...this.recentFonts.filter((f) => f !== name)].slice(0, RECENT_FONTS);
  }
}

const STATES = new WeakMap<EditorSession, HomeState>();

export function homeState(session: EditorSession): HomeState {
  let state = STATES.get(session);
  if (!state) {
    state = new HomeState();
    STATES.set(session, state);
  }
  return state;
}

/**
 * The selection's formatting as Word reports it: the effective value (styles
 * and document defaults included). A caret reads the run it is in (or the
 * paragraph mark in an empty paragraph); a range reads every run it covers.
 */
export function selectionFormats(model: EditorModel): {
  runs: ResolvedRunFormat[];
  /**
   * The font each run shows in the font box: its East Asian font when the run
   * has East Asian text, as Japanese Word shows ＭＳ 明朝 rather than Century.
   */
  fonts: Array<string | undefined>;
  paragraphs: ResolvedParagraphFormat[];
  paragraphStyles: Array<string | undefined>;
} {
  const sel = model.selection;
  if (!sel) return { runs: [], fonts: [], paragraphs: [], paragraphStyles: [] };
  const doc = model.doc;
  const styles = createStyleResolver(doc);
  const ordered = orderSelection(sel);
  // paragraphsInRange takes every cell of a table in range, so a caret uses
  // only its own paragraph.
  const caretPara = ordered.collapsed ? paragraphAt(doc, sel.focus) : undefined;
  const paras = ordered.collapsed
    ? caretPara
      ? [caretPara]
      : []
    : paragraphsInRange(doc, ordered);
  let found: Array<{ run: WmlRun | undefined; format: ResolvedRunFormat }>;
  if (caretPara) {
    const run = runAtPath(doc, sel.focus);
    found = [{ run, format: styles.run(caretPara, run) }];
  } else {
    const owner = new Map(
      paras.flatMap((p) => p.children.filter((c) => c.kind === "run").map((r) => [r, p] as const)),
    );
    found = runsInRange(doc, ordered).flatMap((run) => {
      const para = owner.get(run);
      return para ? [{ run, format: styles.run(para, run) }] : [];
    });
  }
  const runs = found.map((f) => f.format);
  return {
    runs,
    fonts: found.map(({ run, format }) =>
      run && hasEastAsianText(runPlainText(run))
        ? (format.eastAsiaFont ?? format.font)
        : format.font,
    ),
    paragraphs: paras.map((p) => styles.paragraph(p)),
    paragraphStyles: paras.map((p) => getParagraphStyle(p)),
  };
}

function runPlainText(run: WmlRun): string {
  return run.pieces.map((p) => (p.kind === "text" ? p.value : "")).join("");
}

/** One value when every item agrees, else `undefined` (Word's blank "mixed" state). */
export function common<T>(values: readonly T[]): T | undefined {
  const [first] = values;
  return values.every((v) => v === first) ? first : undefined;
}
