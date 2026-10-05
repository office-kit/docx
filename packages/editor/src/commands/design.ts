/**
 * Design tab commands: theme, style set, colours, fonts, paragraph spacing,
 * effects, watermark, page colour and page borders. Each one calls the
 * `@office-kit/docx` design functions.
 */

import {
  applyStyleSet,
  currentStyleSet,
  type DefaultParagraphSpacing,
  type PageBorders,
  type PageColor,
  setDefaultParagraphSpacing,
  setDocumentSettingOnOff,
  setPageColor,
  setSectionProperties,
  setTheme,
  setThemeColors,
  setThemeEffects,
  setThemeFonts,
  setWatermark,
  STYLE_SETS,
  type StyleSetDefinition,
  type ThemeColorScheme,
  type ThemeDefinition,
  type ThemeEffectScheme,
  type ThemeFontScheme,
  THEMES,
  type Watermark,
} from "@office-kit/docx";
import { type SectionTarget, sectionScopeFor } from "./section.js";
import type { Command } from "./types.js";

function first<T>(list: readonly T[], what: string): T {
  const item = list[0];
  if (item === undefined) throw new Error(`No built-in ${what}.`);
  return item;
}

export const applyThemeCommand: Command<{ theme: ThemeDefinition }> = {
  id: "design.theme",
  group: "design",
  label: "Themes",
  run(model, { theme }) {
    setTheme(model.doc, theme);
  },
};

/**
 * Reset to Theme from Template. The editor has no attached template, so this
 * is the theme of Word's Normal template: the Office theme.
 */
export const resetThemeCommand: Command<void> = {
  id: "design.resetTheme",
  group: "design",
  label: "Reset to Theme from Template",
  run(model) {
    setTheme(model.doc, first(THEMES, "theme"));
  },
};

export const applyStyleSetCommand: Command<{ set: StyleSetDefinition }> = {
  id: "design.styleSet",
  group: "design",
  label: "Document Formatting",
  run(model, { set }) {
    applyStyleSet(model.doc, set);
  },
};

/** Reset to the Default Style Set. */
export const resetStyleSetCommand: Command<void> = {
  id: "design.resetStyleSet",
  group: "design",
  label: "Reset to the Default Style Set",
  run(model) {
    applyStyleSet(model.doc, first(STYLE_SETS, "style set"));
  },
};

export const themeColorsCommand: Command<{ colors: ThemeColorScheme }> = {
  id: "design.colors",
  group: "design",
  label: "Colors",
  run(model, { colors }) {
    setThemeColors(model.doc, colors);
  },
};

export const themeFontsCommand: Command<{ fonts: ThemeFontScheme }> = {
  id: "design.fonts",
  group: "design",
  label: "Fonts",
  run(model, { fonts }) {
    setThemeFonts(model.doc, fonts);
  },
};

export const themeEffectsCommand: Command<{ effects: ThemeEffectScheme }> = {
  id: "design.effects",
  group: "design",
  label: "Effects",
  run(model, { effects }) {
    setThemeEffects(model.doc, effects);
  },
};

/**
 * Paragraph Spacing. `"default"` restores the spacing of the document's
 * style set (Word's "Default" entry, "Use the style set's spacing").
 */
export const paragraphSpacingCommand: Command<{ spacing: DefaultParagraphSpacing | "default" }> = {
  id: "design.paragraphSpacing",
  group: "design",
  label: "Paragraph Spacing",
  run(model, { spacing }) {
    if (spacing === "default") {
      const set = currentStyleSet(model.doc) ?? first(STYLE_SETS, "style set");
      setDefaultParagraphSpacing(model.doc, { before: 0, after: set.after, line: set.line });
    } else {
      setDefaultParagraphSpacing(model.doc, spacing);
    }
  },
};

/** Watermark presets, Custom Watermark and Remove Watermark (`undefined`). */
export const watermarkCommand: Command<{ watermark: Watermark | undefined }> = {
  id: "design.watermark",
  group: "design",
  label: "Watermark",
  run(model, { watermark }) {
    setWatermark(model.doc, watermark);
  },
};

/** Page Color; `undefined` is No Color. */
export const pageColorCommand: Command<{ color: PageColor | undefined }> = {
  id: "design.pageColor",
  group: "design",
  label: "Page Color",
  run(model, { color }) {
    setPageColor(model.doc, color);
  },
};

/** The document-wide page border options of Word's Border and Shading Options dialog. */
export interface PageBorderOptions {
  /** Align paragraph borders and table edges with the page border (`w:alignBordersAndEdges`). */
  readonly alignBordersAndEdges?: boolean;
  /** Surround the header (the inverse of `w:bordersDoNotSurroundHeader`). */
  readonly surroundHeader?: boolean;
  /** Surround the footer (the inverse of `w:bordersDoNotSurroundFooter`). */
  readonly surroundFooter?: boolean;
}

/** Borders and Shading ▸ Page Border; `null` borders is the None setting. */
export const pageBordersCommand: Command<{
  borders: PageBorders | null;
  target: SectionTarget;
  options?: PageBorderOptions;
}> = {
  id: "design.pageBorders",
  group: "design",
  label: "Page Borders",
  run(model, { borders, target, options = {} }) {
    setSectionProperties(
      model.doc,
      { pageBorders: borders },
      sectionScopeFor(model, target, "nextPage"),
    );
    if (options.alignBordersAndEdges !== undefined) {
      setDocumentSettingOnOff(model.doc, "alignBordersAndEdges", options.alignBordersAndEdges);
    }
    if (options.surroundHeader !== undefined) {
      setDocumentSettingOnOff(model.doc, "bordersDoNotSurroundHeader", !options.surroundHeader);
    }
    if (options.surroundFooter !== undefined) {
      setDocumentSettingOnOff(model.doc, "bordersDoNotSurroundFooter", !options.surroundFooter);
    }
  },
};

export const designCommands = [
  applyThemeCommand,
  resetThemeCommand,
  applyStyleSetCommand,
  resetStyleSetCommand,
  themeColorsCommand,
  themeFontsCommand,
  themeEffectsCommand,
  paragraphSpacingCommand,
  watermarkCommand,
  pageColorCommand,
  pageBordersCommand,
];
