/** Page layout: sections, pagination, numbering — the pure half of Print Layout. */

export { createFlagResolver, type ParagraphFlags } from "./flags.js";
export {
  collectNoteReferences,
  endnotesAtSectionEnd,
  type NoteOccurrence,
  numberNotes,
} from "./notes.js";
export { formatNumber } from "./number-format.js";
export {
  type Fragment,
  type LaidOutColumn,
  type LaidOutPage,
  type LayoutBlock,
  type LayoutNote,
  type LayoutRow,
  type PageKind,
  paginate,
  type PaginatorOptions,
  type PaginatorSection,
  type Region,
  type SectionStart,
  sectionPageCounts,
  type VerticalAlign,
} from "./paginate.js";
export {
  type BorderLine,
  type ColumnSpec,
  type DocumentGrid,
  documentSections,
  type LayoutSettings,
  layoutSettings,
  type LineNumbering,
  type Margins,
  type NoteProperties,
  type PageBorders,
  type PageNumbering,
  type SectionModel,
} from "./sections.js";
