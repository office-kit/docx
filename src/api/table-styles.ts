/**
 * Table styles: applying a style and its Table Style Options (`w:tblLook`),
 * Word's built-in table style gallery, and editing a table style's
 * conditional formatting (Modify Table Style).
 */

import type { XmlElement } from "../internal/xml/index.js";
import {
  type BuiltInTableStyle,
  BUILT_IN_TABLE_STYLE_LIST,
  builtInTableStyleElement,
  normalTableStyle,
  OFFICE_THEME_COLORS,
  type TableStyleCategory,
  type TableStyleRegion,
  type ThemeColors,
} from "../internal/wordprocessingml/table-style-defs.js";
import {
  BORDER_ORDER,
  borderAttrs,
  PPR_ORDER,
  RPR_ORDER,
  STYLE_ORDER,
  TBL_PR_ORDER,
  TBL_STYLE_PR_ORDER,
  TC_PR_ORDER,
  ensureWChild,
  removeWChild,
  upsertWChild,
  wAttrOf,
  wChild,
  wChildren,
  wEl,
} from "../internal/wordprocessingml/table-xml.js";
import type { WmlTable } from "../internal/wordprocessingml/index.js";
import { addStyle, type Docx, getRawPartRoot, stylesPart, xmlPartNames } from "./docx.js";
import type { TableBorder } from "./table-tools.js";

export type { BuiltInTableStyle, TableStyleCategory, TableStyleRegion };

/** Word's built-in table styles in gallery order (Plain Tables, Grid Tables, List Tables). */
export const BUILT_IN_TABLE_STYLES: readonly BuiltInTableStyle[] = BUILT_IN_TABLE_STYLE_LIST;

/** Apply a table style by id (`w:tblStyle`); undefined removes it (Clear). */
export function setTableStyle(table: WmlTable, styleId: string | undefined): void {
  table.tblPr ??= wEl("tblPr");
  if (styleId === undefined) removeWChild(table.tblPr, "tblStyle");
  else upsertWChild(table.tblPr, wEl("tblStyle", { val: styleId }), TBL_PR_ORDER);
}

/** The table's style id, if it has one. */
export function getTableStyle(table: WmlTable): string | undefined {
  return wAttrOf(wChild(table.tblPr, "tblStyle"), "val");
}

/**
 * Which parts of the table style apply (Table Design ▸ Table Style Options,
 * `w:tblLook` §17.4.56).
 */
export interface TableLook {
  readonly headerRow: boolean;
  readonly totalRow: boolean;
  readonly firstColumn: boolean;
  readonly lastColumn: boolean;
  readonly bandedRows: boolean;
  readonly bandedColumns: boolean;
}

// The ECMA-376 1st edition bitmask in tblLook/@w:val (§17.4.56), which Word
// still writes next to the ISO/IEC 29500 attributes.
const LOOK_BITS = {
  firstRow: 0x0020,
  lastRow: 0x0040,
  firstColumn: 0x0080,
  lastColumn: 0x0100,
  noHBand: 0x0200,
  noVBand: 0x0400,
} as const;
const HEX_WIDTH = 4;
const ON = new Set(["1", "true", "on"]);

/**
 * The table's style options. Without a `w:tblLook`, nothing is emphasized
 * but rows and columns are banded (every flag of the bitmask is 0).
 */
export function getTableLook(table: WmlTable): TableLook {
  const look = wChild(table.tblPr, "tblLook");
  const bits = Number.parseInt(wAttrOf(look, "val") ?? "0", 16) || 0;
  const flag = (name: keyof typeof LOOK_BITS): boolean => {
    const attr = wAttrOf(look, name);
    return attr === undefined ? (bits & LOOK_BITS[name]) !== 0 : ON.has(attr);
  };
  return {
    headerRow: flag("firstRow"),
    totalRow: flag("lastRow"),
    firstColumn: flag("firstColumn"),
    lastColumn: flag("lastColumn"),
    bandedRows: !flag("noHBand"),
    bandedColumns: !flag("noVBand"),
  };
}

/** Set the table's style options (`w:tblLook`), in both the bitmask and attribute forms. */
export function setTableLook(table: WmlTable, look: TableLook): void {
  const flags = {
    firstRow: look.headerRow,
    lastRow: look.totalRow,
    firstColumn: look.firstColumn,
    lastColumn: look.lastColumn,
    noHBand: !look.bandedRows,
    noVBand: !look.bandedColumns,
  };
  let bits = 0;
  for (const [name, on] of Object.entries(flags)) {
    if (on) bits |= LOOK_BITS[name as keyof typeof LOOK_BITS];
  }
  const attrs: Record<string, string> = {
    val: bits.toString(16).toUpperCase().padStart(HEX_WIDTH, "0"),
  };
  for (const [name, on] of Object.entries(flags)) attrs[name] = on ? "1" : "0";
  table.tblPr ??= wEl("tblPr");
  upsertWChild(table.tblPr, wEl("tblLook", attrs), TBL_PR_ORDER);
}

const THEME_PART_DIR = "/word/theme/";
const THEME_KEYS = [
  "dk1",
  "lt1",
  "accent1",
  "accent2",
  "accent3",
  "accent4",
  "accent5",
  "accent6",
] as const;
const HEX = /^[0-9A-Fa-f]{6}$/;

/** The document theme's colors, falling back to the Office theme for each missing one. */
function themeColors(doc: Docx): ThemeColors {
  const name = xmlPartNames(doc).find((n) => n.startsWith(THEME_PART_DIR));
  const root = name ? getRawPartRoot(doc, name) : undefined;
  const scheme = findLocal(findLocal(root, "themeElements"), "clrScheme");
  const out: Record<string, string> = { ...OFFICE_THEME_COLORS };
  for (const key of THEME_KEYS) {
    const slot = findLocal(scheme, key);
    const color = slot?.children.find((c): c is XmlElement => c.kind === "element");
    // <a:srgbClr val> or <a:sysClr lastClr> (dk1/lt1 are usually system colors).
    const attr = color?.name.local === "sysClr" ? "lastClr" : "val";
    const value = color?.attrs.find((a) => a.name.local === attr)?.value;
    if (value && HEX.test(value)) out[key] = value.toUpperCase();
  }
  return {
    dk1: out.dk1!,
    lt1: out.lt1!,
    accent1: out.accent1!,
    accent2: out.accent2!,
    accent3: out.accent3!,
    accent4: out.accent4!,
    accent5: out.accent5!,
    accent6: out.accent6!,
  };
}

function findLocal(el: XmlElement | undefined, local: string): XmlElement | undefined {
  return el?.children.find((c): c is XmlElement => c.kind === "element" && c.name.local === local);
}

const NORMAL_TABLE = "TableNormal";

function defaultTableStyleId(doc: Docx): string | undefined {
  for (const style of stylesPart(doc)?.styles ?? []) {
    if (wAttrOf(style, "type") === "table" && wAttrOf(style, "default") === "1") {
      return wAttrOf(style, "styleId");
    }
  }
  return undefined;
}

/**
 * The definition Word writes for a built-in table style, colored from the
 * document's theme. Throws for an id that is not a built-in table style.
 */
export function builtInTableStyle(doc: Docx, styleId: string): XmlElement {
  const el = builtInTableStyleElement(
    styleId,
    themeColors(doc),
    defaultTableStyleId(doc) ?? NORMAL_TABLE,
  );
  if (!el) throw new RangeError(`"${styleId}" is not a built-in table style.`);
  return el;
}

/**
 * Add a built-in table style's definition to styles.xml (and Normal Table,
 * which it is based on) unless the document already defines that style id.
 * Word does the same the first time a table uses a gallery style.
 */
export function addBuiltInTableStyle(doc: Docx, styleId: string): void {
  const definition = builtInTableStyle(doc, styleId);
  let part = stylesPart(doc);
  if (!part) {
    // addStyle creates styles.xml (with Word's Normal Table) on first use.
    addStyle(doc, { type: "table", styleId: NORMAL_TABLE, name: "Normal Table", default: true });
    part = stylesPart(doc);
    const index = part?.styles.findIndex((s) => wAttrOf(s, "styleId") === NORMAL_TABLE) ?? -1;
    if (part && index >= 0) part.styles[index] = normalTableStyle();
  }
  if (!part) throw new Error("styles.xml could not be created.");
  const ids = new Set(part.styles.map((s) => wAttrOf(s, "styleId")));
  if (!defaultTableStyleId(doc)) part.styles.push(normalTableStyle());
  if (!ids.has(styleId)) part.styles.push(definition);
  doc.stylesDirty = true;
}

/** Formatting for one region of a table style (Modify Table Style ▸ Apply formatting to). */
export interface TableStyleFormatting {
  readonly bold?: boolean;
  readonly italic?: boolean;
  /** Text color, hex RGB or `auto`. */
  readonly color?: string;
  /** Cell shading fill, hex RGB; `auto` for none. */
  readonly fill?: string;
  /** Paragraph alignment of the region's text. */
  readonly alignment?: "left" | "center" | "right" | "both";
  /** Cell vertical alignment. */
  readonly verticalAlignment?: "top" | "center" | "bottom";
  /**
   * Borders around and inside the region. A side mapped to `null` is set to
   * no border; an omitted side is left as it is.
   */
  readonly borders?: Partial<
    Record<"top" | "left" | "bottom" | "right" | "insideH" | "insideV", TableBorder | null>
  >;
}

const COLOR = /^(auto|[0-9A-Fa-f]{6})$/;

function findTableStyle(doc: Docx, styleId: string): XmlElement {
  const style = stylesPart(doc)?.styles.find(
    (s) => wAttrOf(s, "styleId") === styleId && wAttrOf(s, "type") === "table",
  );
  if (!style) throw new RangeError(`No table style "${styleId}".`);
  return style;
}

/**
 * Change a table style's formatting for a region (`wholeTable` edits the
 * style's own properties; the others its `<w:tblStylePr w:type>`).
 */
export function setTableStyleFormatting(
  doc: Docx,
  styleId: string,
  region: TableStyleRegion,
  formatting: TableStyleFormatting,
): void {
  for (const c of [formatting.color, formatting.fill]) {
    if (c !== undefined && !COLOR.test(c)) throw new RangeError(`Invalid color "${c}".`);
  }
  const style = findTableStyle(doc, styleId);
  let target: XmlElement;
  let order: readonly string[];
  if (region === "wholeTable") {
    target = style;
    order = STYLE_ORDER;
  } else {
    const existing = wChildren(style, "tblStylePr").find((e) => wAttrOf(e, "type") === region);
    target = existing ?? wEl("tblStylePr", { type: region });
    if (!existing) (style.children as XmlElement[]).push(target);
    order = TBL_STYLE_PR_ORDER;
  }
  const rPr = (): XmlElement => ensureWChild(target, "rPr", order);
  if (formatting.bold !== undefined) {
    for (const local of ["b", "bCs"]) {
      upsertWChild(rPr(), wEl(local, formatting.bold ? {} : { val: "0" }), RPR_ORDER);
    }
  }
  if (formatting.italic !== undefined) {
    for (const local of ["i", "iCs"]) {
      upsertWChild(rPr(), wEl(local, formatting.italic ? {} : { val: "0" }), RPR_ORDER);
    }
  }
  if (formatting.color !== undefined) {
    upsertWChild(rPr(), wEl("color", { val: formatting.color.toUpperCase() }), RPR_ORDER);
  }
  if (formatting.alignment !== undefined) {
    upsertWChild(
      ensureWChild(target, "pPr", order),
      wEl("jc", { val: formatting.alignment }),
      PPR_ORDER,
    );
  }
  const tcPr = (): XmlElement => ensureWChild(target, "tcPr", order);
  if (formatting.fill !== undefined) {
    upsertWChild(
      tcPr(),
      wEl("shd", { val: "clear", color: "auto", fill: formatting.fill.toUpperCase() }),
      TC_PR_ORDER,
    );
  }
  if (formatting.verticalAlignment !== undefined) {
    upsertWChild(tcPr(), wEl("vAlign", { val: formatting.verticalAlignment }), TC_PR_ORDER);
  }
  if (formatting.borders) {
    // The whole table's borders live in its tblPr; a region's in its tcPr.
    const container =
      region === "wholeTable"
        ? ensureWChild(ensureWChild(target, "tblPr", order), "tblBorders", TBL_PR_ORDER)
        : ensureWChild(tcPr(), "tcBorders", TC_PR_ORDER);
    for (const [side, border] of Object.entries(formatting.borders)) {
      if (border === undefined) continue;
      upsertWChild(container, wEl(side, borderAttrs(border ?? undefined)), BORDER_ORDER);
    }
  }
  doc.stylesDirty = true;
}
