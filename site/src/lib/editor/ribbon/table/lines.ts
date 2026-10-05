import type { MessageKey } from "../../i18n/index.svelte";

/** Line Style choices (ST_Border values), in the order of Word's list. */
export const LINE_STYLES: readonly { style: string; label: MessageKey }[] = [
  { style: "single", label: "tbl.line.single" },
  { style: "dotted", label: "tbl.line.dotted" },
  { style: "dashed", label: "tbl.line.dashed" },
  { style: "dotDash", label: "tbl.line.dotDash" },
  { style: "double", label: "tbl.line.double" },
  { style: "triple", label: "tbl.line.triple" },
  { style: "thickThinSmallGap", label: "tbl.line.thickThin" },
  { style: "wave", label: "tbl.line.wave" },
];

/** Line Weight choices: `size` in eighths of a point. */
export const LINE_WEIGHTS: readonly { size: number; label: string }[] = [
  { size: 2, label: "¼ pt" },
  { size: 4, label: "½ pt" },
  { size: 6, label: "¾ pt" },
  { size: 8, label: "1 pt" },
  { size: 12, label: "1½ pt" },
  { size: 18, label: "2¼ pt" },
  { size: 24, label: "3 pt" },
  { size: 36, label: "4½ pt" },
  { size: 48, label: "6 pt" },
];

// Dialog ids (see Dialog.svelte), shared by the button that opens a dialog and the dialog.
export const DIALOGS = {
  insertTable: "tbl-insert",
  convertTextToTable: "tbl-text-to-table",
  bordersShading: "tbl-borders-shading",
  modifyStyle: "tbl-modify-style",
  newStyle: "tbl-new-style",
  properties: "tbl-properties",
  deleteCells: "tbl-delete-cells",
  splitCells: "tbl-split-cells",
  cellMargins: "tbl-cell-margins",
  sort: "tbl-sort",
  convertToText: "tbl-to-text",
  formula: "tbl-formula",
} as const;
