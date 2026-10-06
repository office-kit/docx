/**
 * `@office-kit/docx-editor` — an MS Office-like WYSIWYG editing core for
 * word-kit `.docx` documents.
 *
 * The editor never serializes OOXML itself: every mutation runs through a
 * {@link Command} that calls the `@office-kit/docx` public API, keeping that
 * library the single source of truth. Completeness against the WordprocessingML
 * spec is tracked and enforced by the capability ledger under `./coverage`.
 *
 * @packageDocumentation
 */

import { createDocx, type Docx, openDocx } from "@office-kit/docx";
import { EditorModel } from "./model.js";

export { EditorModel, type ChangeListener, type EditorSnapshot } from "./model.js";
export { type Command, type FeatureGroup, runCommand } from "./commands/types.js";
export type { SectionTarget } from "./commands/section.js";
export type {
  PastedBlock,
  PastedParagraph,
  PastedRun,
  PastedRunFormat,
  PastedTable,
} from "./commands/structure.js";
export { parseClipboardHtml } from "./clipboard-html.js";
export { ALL_COMMANDS, COMMAND_IDS, commandsInGroup, getCommand } from "./commands/registry.js";
// Named command objects + group arrays for typed UI wiring.
export * as commands from "./commands/index.js";
export {
  type CellCoord,
  type DocPosition,
  type OrderedSelection,
  type Selection,
  type StoryRef,
  caretAt,
  orderSelection,
  sameStory,
  storyKeyOf,
} from "./selection.js";
export { renderDocumentHtml, renderBlocksHtml, paragraphPlainText } from "./render.js";
export { hasEastAsianText, isEastAsianFont } from "./east-asian-fonts.js";
export { PAGE_FIELDS, fieldFormatSwitch, fieldType } from "./render-fields.js";
export * from "./layout/index.js";
export { highlightCss } from "./highlight.js";
export { EMU_PER_PX, imageDataUrl, pictureSvg, type PictureSvgInput } from "./render-drawing.js";
export { chartSvg } from "./render-chart.js";
export { CROP_SHAPES, presetPath } from "./preset-geometry.js";
export {
  floatFrameStart,
  floatOrigin,
  layoutFloatingObjects,
  type PageBox,
  readFloat,
} from "./float-layout.js";
export { drawingIndexAt, drawingPositionOf } from "./drawing-access.js";
export { renderPictHtml, type VmlRenderContext, vmlPathPreviewSvg } from "./render-vml.js";
export { type SvgSubpath, vmlPathToSvg } from "./vml-path.js";
export { type InkPoint, type RecognizedShape, recognizeInkShape } from "./ink-shape.js";
export {
  type OutlineOp,
  type OutlineRow,
  applyOutlineOp,
  outlineToNodes,
  outlineRows,
} from "./smartart-outline.js";
export {
  type ResolvedParagraphFormat,
  type ResolvedRunFormat,
  type PageGeometry,
  type StyleResolver,
  createStyleResolver,
  pageGeometry,
} from "./resolve.js";
export { positionFromDom, readDomSelection, runPoint, STORY_ATTR } from "./dom-selection.js";
export { layoutScaledText } from "./scale-layout.js";
export { layoutTabStops } from "./tab-layout.js";
export {
  adjacentCellPosition,
  cellStart,
  selectInTable,
  type TableSelection,
  type TableSelectTarget,
  tableSelection,
} from "./table-selection.js";
export {
  type BorderSpec,
  type ResolvedCell,
  type ResolvedTableFormat,
  resolveTable,
  shadingColor,
} from "./table-format.js";
export { type TablePreviewCell, tableStylePreviews } from "./table-preview.js";
export { type TablePropertiesSnapshot, tablePropertiesSnapshot } from "./table-properties.js";
export { isEmptyParagraph, runAtPath, setSimpleRunText } from "./text-edit.js";
export {
  isTrackingRevisions,
  type Reviewer,
  reviewerOf,
  setReviewer,
  trackRevisionsSetting,
} from "./track-changes.js";
export {
  adjacentReviewMark,
  type ReviewMark,
  reviewMarks,
  revisionIdsAtSelection,
} from "./review-nav.js";
export { BODY_TEXT_LEVEL, outlineLevelOf } from "./outline.js";
export { AUTHOR_COLORS } from "./render-revisions.js";
export { isEditingLocked, PROTECTED_MESSAGE, protectionRefusal } from "./protection-guard.js";
export { type RawNode, rawTrees, allRawElements, xmlParts, partRawTree } from "./raw-tree.js";
export { blocks, bodyOf, paragraphAt, paragraphsInRange } from "./doc-access.js";
export { releasePendingFormat, type RunRef, runsInRange } from "./selection-runs.js";
export {
  type PageBackground,
  type ResolvedSection,
  resolvePageBackground,
  resolveSectionLayout,
} from "./section-layout.js";
export { resolveTheme } from "./theme.js";
export * from "./insert-queries.js";
export { renderMath } from "./render-math.js";
export { symbolGlyph } from "./render-fields.js";
export {
  bodyParagraphs,
  compileFind,
  type FindFormat,
  type FindMatch,
  findMatches,
  type FindOptions,
  selectionOf,
} from "./find.js";
export {
  BULLET_PRESETS,
  type ListPreset,
  MULTILEVEL_PRESETS,
  NUMBERING_PRESETS,
} from "./list-presets.js";
export { bulletGlyph, formatListNumber } from "./list-numbering.js";
export {
  DEFAULT_THEME_COLORS,
  type ThemePalette,
  type ThemeSwatch,
  themeColorGrid,
  themePalette,
} from "./theme-color.js";
export { type CaseMode, changeCase } from "./text-case.js";
export { type FontPatch, type ParagraphPatch } from "./commands/format-patch.js";
export {
  type ParagraphToggle,
  readThemeFonts,
  type ResolvedBorder,
  type ResolvedShading,
  type RunToggle,
  type ThemeFonts,
} from "./resolve.js";

// Coverage / capability ledger — the completeness-guarantee surface.
export {
  type Capability,
  type Disposition,
  classify,
  coverageSummary,
  LEDGER,
} from "./capability/ledger.js";
export { renderCoverageMarkdown } from "./capability/report.js";

/** Create an editor over a fresh, empty document. */
export function createEditor(): EditorModel {
  return new EditorModel(createDocx({ paragraphs: [""] }));
}

/** Create an editor over an existing `.docx` (bytes). */
export function openEditor(bytes: Uint8Array): EditorModel {
  return new EditorModel(openDocx(bytes));
}

/** Create an editor over an already-open {@link Docx}. */
export function editorFor(doc: Docx): EditorModel {
  return new EditorModel(doc);
}
