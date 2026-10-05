import type { XmlAttr, XmlElement } from "../xml/index.js";
import { WML_NS } from "./namespaces.js";
import { normalizeVerticalMerges } from "./table-grid.js";
import { TBL_PR_ORDER, TC_PR_ORDER, TR_PR_ORDER, upsertWChild } from "./table-xml.js";
import type {
  WmlParagraph,
  WmlRun,
  WmlRunPiece,
  WmlTable,
  WmlTableCell,
  WmlTableRow,
} from "./types.js";

/**
 * Construct a `<w:t>`-bearing run containing the given plain text.
 * `xml:space="preserve"` is applied automatically when the text has
 * leading or trailing whitespace.
 *
 * Tab characters (`\t`) and newlines (`\n`, `\r\n`) in the input are
 * segmented into `<w:tab/>` and `<w:br/>` elements respectively, so
 * the visible rendering in Word matches what the caller wrote.
 */
export function buildTextRun(text: string): WmlRun {
  return { kind: "run", pieces: splitTextIntoPieces(text), extras: [] };
}

/**
 * Split a plain string into a sequence of WmlRunPiece values, replacing
 * `\t` with a tab piece and `\n` (or `\r\n`) with a break piece. Empty
 * input produces a single empty text piece so downstream serializers
 * always emit something for the run.
 */
function splitTextIntoPieces(text: string): WmlRunPiece[] {
  if (text === "") {
    return [{ kind: "text", value: "", preserveSpace: false }];
  }
  const pieces: WmlRunPiece[] = [];
  let buffer = "";
  const flush = (): void => {
    if (buffer.length === 0) return;
    pieces.push({ kind: "text", value: buffer, preserveSpace: /^\s|\s$/.test(buffer) });
    buffer = "";
  };
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (ch === "\t") {
      flush();
      pieces.push({ kind: "tab" });
    } else if (ch === "\r" && text[i + 1] === "\n") {
      flush();
      pieces.push({ kind: "break" });
      i++; // consume the \n half of \r\n
    } else if (ch === "\n" || ch === "\r") {
      flush();
      pieces.push({ kind: "break" });
    } else {
      buffer += ch;
    }
  }
  flush();
  return pieces;
}

/** Construct a paragraph with a single text run. */
export function buildTextParagraph(text: string): WmlParagraph {
  return { kind: "paragraph", children: [buildTextRun(text)], extras: [] };
}

export interface BuildTableOptions {
  /**
   * Total table width in twips (1/20 of a point, 1/1440 of an inch).
   * Distributed evenly across columns. Defaults to 9000 twips (~6.25").
   */
  readonly totalWidthTwips?: number;
}

/**
 * Construct a basic table from a row-major matrix of strings. Each cell
 * becomes a single paragraph with one run containing the supplied text.
 * The number of columns is inferred from the longest row; shorter rows are
 * padded with empty cells.
 */
export function buildTextTable(
  rows: ReadonlyArray<ReadonlyArray<string>>,
  options: BuildTableOptions = {},
): WmlTable {
  const colCount = rows.reduce((max, row) => Math.max(max, row.length), 0);
  const totalWidth = options.totalWidthTwips ?? 9000;
  const colWidth = colCount > 0 ? Math.floor(totalWidth / colCount) : totalWidth;
  const tblGrid: XmlElement = {
    kind: "element",
    name: { uri: WML_NS, local: "tblGrid", prefix: "w" },
    attrs: [],
    children: Array.from({ length: colCount }, () => ({
      kind: "element" as const,
      name: { uri: WML_NS, local: "gridCol", prefix: "w" },
      attrs: [wmlAttr("w", String(colWidth))],
      children: [],
      xmlSpace: "default" as const,
      selfClosing: true,
    })),
    xmlSpace: "default",
    selfClosing: false,
  };
  const tblPr: XmlElement = {
    kind: "element",
    name: { uri: WML_NS, local: "tblPr", prefix: "w" },
    attrs: [],
    children: [
      {
        kind: "element",
        name: { uri: WML_NS, local: "tblW", prefix: "w" },
        attrs: [wmlAttr("w", String(totalWidth)), wmlAttr("type", "dxa")],
        children: [],
        xmlSpace: "default",
        selfClosing: true,
      },
      {
        kind: "element",
        name: { uri: WML_NS, local: "tblLook", prefix: "w" },
        attrs: [
          wmlAttr("val", "04A0"),
          wmlAttr("firstRow", "1"),
          wmlAttr("lastRow", "0"),
          wmlAttr("firstColumn", "1"),
          wmlAttr("lastColumn", "0"),
          wmlAttr("noHBand", "0"),
          wmlAttr("noVBand", "1"),
        ],
        children: [],
        xmlSpace: "default",
        selfClosing: true,
      },
    ],
    xmlSpace: "default",
    selfClosing: false,
  };
  const tableRows: WmlTableRow[] = rows.map((rowTexts): WmlTableRow => {
    const cells: WmlTableCell[] = [];
    for (let c = 0; c < colCount; c++) {
      const text = rowTexts[c] ?? "";
      cells.push({
        tcPr: cellPropertiesWithWidth(colWidth),
        paragraphs: [buildTextParagraph(text)],
        extras: [],
      });
    }
    return { cells, extras: [] };
  });
  return {
    kind: "table",
    tblPr,
    tblGrid,
    rows: tableRows,
    extras: [],
  };
}

function cellPropertiesWithWidth(widthTwips: number): XmlElement {
  return {
    kind: "element",
    name: { uri: WML_NS, local: "tcPr", prefix: "w" },
    attrs: [],
    children: [
      {
        kind: "element",
        name: { uri: WML_NS, local: "tcW", prefix: "w" },
        attrs: [wmlAttr("w", String(widthTwips)), wmlAttr("type", "dxa")],
        children: [],
        xmlSpace: "default",
        selfClosing: true,
      },
    ],
    xmlSpace: "default",
    selfClosing: false,
  };
}

export type TableBorderStyle = "none" | "single" | "double" | "dashed" | "dotted" | "thick";

export interface TableBordersOptions {
  /** Stroke style — defaults to `"single"`. Pass `"none"` to clear borders. */
  readonly style?: TableBorderStyle;
  /** Stroke width in eighths of a point — Word's `w:sz`. Defaults to 4 (½ pt). */
  readonly sizeEighthsOfPoint?: number;
  /** Hex RGB without a leading `#`. Defaults to `"auto"`. */
  readonly color?: string;
  /**
   * When `true` (default), also draws the interior gridlines
   * (`<w:insideH/>` and `<w:insideV/>`). Disable for an outer-only frame.
   */
  readonly inside?: boolean;
}

/**
 * Set uniform borders on every side of `table`. Replaces any existing
 * `<w:tblBorders>` block; the rest of `<w:tblPr>` (width, look, etc.)
 * is preserved.
 *
 * Example: a half-point single-line frame around the whole table —
 *   `setTableBorders(table, {});`
 *
 * Example: thick double border, outer only —
 *   `setTableBorders(table, { style: "double", sizeEighthsOfPoint: 12, inside: false });`
 *
 * Example: no borders at all —
 *   `setTableBorders(table, { style: "none" });`
 */
export function setTableBorders(table: WmlTable, options: TableBordersOptions = {}): void {
  const style = options.style ?? "single";
  const sz = String(options.sizeEighthsOfPoint ?? 4);
  const color = options.color ?? "auto";
  const inside = options.inside ?? true;

  const tblPr = table.tblPr ?? {
    kind: "element",
    name: { uri: WML_NS, local: "tblPr", prefix: "w" },
    attrs: [],
    children: [],
    xmlSpace: "default",
    selfClosing: false,
  };
  const children = tblPr.children as XmlElement[];
  // Drop any existing tblBorders before re-adding ours.
  for (let i = children.length - 1; i >= 0; i--) {
    const c = children[i];
    if (c && c.kind === "element" && c.name.uri === WML_NS && c.name.local === "tblBorders") {
      children.splice(i, 1);
    }
  }
  const borderAttrs = (): XmlAttr[] => [
    wmlAttr("val", style),
    wmlAttr("sz", sz),
    wmlAttr("space", "0"),
    wmlAttr("color", color),
  ];
  const sides = ["top", "left", "bottom", "right"];
  if (inside) sides.push("insideH", "insideV");
  const bordersEl: XmlElement = {
    kind: "element",
    name: { uri: WML_NS, local: "tblBorders", prefix: "w" },
    attrs: [],
    children: sides.map((s) => wmlEmpty(s, borderAttrs())),
    xmlSpace: "default",
    selfClosing: false,
  };
  upsertWChild(tblPr, bordersEl, TBL_PR_ORDER);

  if (!table.tblPr) table.tblPr = tblPr;
}

export interface TableCellShadingOptions {
  /**
   * Hex RGB fill colour without a leading `#` (e.g. `"DDDDDD"`). Defaults
   * to `"auto"` which lets Word pick from the active style.
   */
  readonly fill?: string;
  /**
   * Pattern overlaid on the fill. `"clear"` (default) is a flat fill;
   * `"solid"` is equivalent to `"clear"` with the foreground colour
   * forced; the rest are crosshatches commonly used for table emphasis.
   */
  readonly pattern?:
    | "clear"
    | "solid"
    | "horzStripe"
    | "vertStripe"
    | "diagStripe"
    | "diagCross"
    | "thinHorzStripe"
    | "thinVertStripe";
  /**
   * Pattern stroke colour (only meaningful when `pattern !== "clear"`).
   * Defaults to `"auto"`.
   */
  readonly color?: string;
}

/**
 * Apply cell shading to a table cell. Replaces any existing `<w:shd>` on
 * the cell's `<w:tcPr>`; other tcPr children (width, vertical alignment,
 * cell margins) are preserved.
 *
 * Example — light grey header cell:
 *   `setTableCellShading(headerCell, { fill: "E0E0E0" });`
 */
export function setTableCellShading(cell: WmlTableCell, options: TableCellShadingOptions): void {
  const fill = options.fill ?? "auto";
  const pattern = options.pattern ?? "clear";
  const color = options.color ?? "auto";
  const tcPr = cell.tcPr ?? {
    kind: "element",
    name: { uri: WML_NS, local: "tcPr", prefix: "w" },
    attrs: [],
    children: [],
    xmlSpace: "default",
    selfClosing: false,
  };
  const children = tcPr.children as XmlElement[];
  for (let i = children.length - 1; i >= 0; i--) {
    const c = children[i];
    if (c && c.kind === "element" && c.name.uri === WML_NS && c.name.local === "shd") {
      children.splice(i, 1);
    }
  }
  upsertWChild(
    tcPr,
    wmlEmpty("shd", [wmlAttr("val", pattern), wmlAttr("color", color), wmlAttr("fill", fill)]),
    TC_PR_ORDER,
  );
  if (!cell.tcPr) cell.tcPr = tcPr;
}

/**
 * Vertical alignment for content inside a table cell. Maps to `<w:vAlign>`.
 */
export type TableCellVerticalAlign = "top" | "center" | "bottom";

/** Set the vertical alignment of content inside a cell. */
export function setTableCellVerticalAlign(cell: WmlTableCell, align: TableCellVerticalAlign): void {
  const tcPr = cell.tcPr ?? {
    kind: "element",
    name: { uri: WML_NS, local: "tcPr", prefix: "w" },
    attrs: [],
    children: [],
    xmlSpace: "default",
    selfClosing: false,
  };
  const children = tcPr.children as XmlElement[];
  for (let i = children.length - 1; i >= 0; i--) {
    const c = children[i];
    if (c && c.kind === "element" && c.name.uri === WML_NS && c.name.local === "vAlign") {
      children.splice(i, 1);
    }
  }
  upsertWChild(tcPr, wmlEmpty("vAlign", [wmlAttr("val", align)]), TC_PR_ORDER);
  if (!cell.tcPr) cell.tcPr = tcPr;
}

export type TableRowHeightRule = "atLeast" | "exact" | "auto";

/**
 * Set an explicit row height in twips. `rule` defaults to `"atLeast"`,
 * meaning the row grows past `heightTwips` if its content overflows.
 * `"exact"` clips overflowing content; `"auto"` lets Word resize freely
 * and ignores `heightTwips` (Word's UI calls this "automatic").
 */
export function setTableRowHeight(
  row: WmlTableRow,
  heightTwips: number,
  rule: TableRowHeightRule = "atLeast",
): void {
  const trPr = row.trPr ?? {
    kind: "element",
    name: { uri: WML_NS, local: "trPr", prefix: "w" },
    attrs: [],
    children: [],
    xmlSpace: "default",
    selfClosing: false,
  };
  const children = trPr.children as XmlElement[];
  for (let i = children.length - 1; i >= 0; i--) {
    const c = children[i];
    if (c && c.kind === "element" && c.name.uri === WML_NS && c.name.local === "trHeight") {
      children.splice(i, 1);
    }
  }
  upsertWChild(
    trPr,
    wmlEmpty("trHeight", [wmlAttr("val", String(heightTwips)), wmlAttr("hRule", rule)]),
    TR_PR_ORDER,
  );
  if (!row.trPr) row.trPr = trPr;
}

/**
 * Mark a table row as a *header row* — its content is repeated at the
 * top of every page when the table breaks across pages. Word's UI
 * equivalent is "Repeat as header row at the top of each page".
 *
 * Pass `false` to clear the marker.
 */
export function setTableRowAsHeader(row: WmlTableRow, isHeader = true): void {
  const trPr = row.trPr ?? {
    kind: "element",
    name: { uri: WML_NS, local: "trPr", prefix: "w" },
    attrs: [],
    children: [],
    xmlSpace: "default",
    selfClosing: false,
  };
  const children = trPr.children as XmlElement[];
  for (let i = children.length - 1; i >= 0; i--) {
    const c = children[i];
    if (c && c.kind === "element" && c.name.uri === WML_NS && c.name.local === "tblHeader") {
      children.splice(i, 1);
    }
  }
  if (isHeader) upsertWChild(trPr, wmlEmpty("tblHeader", []), TR_PR_ORDER);
  if (!row.trPr) row.trPr = trPr;
}

/**
 * Every value `<w:highlight w:val>` accepts: ST_HighlightColor (ECMA-376
 * Part 1 §17.18.40), in Word's palette order. The list is closed and has no
 * RGB form; an arbitrary color behind text is run shading (`<w:shd>`).
 * `"none"` is an explicit "no highlight" that also overrides a highlight
 * inherited from a style.
 */
export const HIGHLIGHT_COLORS = [
  "yellow",
  "green",
  "cyan",
  "magenta",
  "blue",
  "red",
  "darkBlue",
  "darkCyan",
  "darkGreen",
  "darkMagenta",
  "darkRed",
  "darkYellow",
  "darkGray",
  "lightGray",
  "black",
  "white",
  "none",
] as const;

export type HighlightColor = (typeof HIGHLIGHT_COLORS)[number];

const HIGHLIGHT_COLOR_SET: ReadonlySet<string> = new Set(HIGHLIGHT_COLORS);

// ST_HexColor (wml.xsd): `auto`, or ST_HexColorRGB — hexBinary of length 3,
// i.e. six hex digits in either case.
const HEX_COLOR = /^(?:auto|[0-9A-Fa-f]{6})$/;

/**
 * Throw if `formatting` would write invalid OOXML. Every writer calls this
 * before touching the document, so a rejected call changes nothing. Only
 * schema rules are enforced; application limits (Word's 1–1638 pt font-size
 * box) are not, since other values are valid files.
 */
function assertWritableRunFormatting(formatting: RunFormatting): void {
  if (formatting.highlight !== undefined && !HIGHLIGHT_COLOR_SET.has(formatting.highlight)) {
    throw new RangeError(
      `highlight must be one of ${HIGHLIGHT_COLORS.join(", ")} (ST_HighlightColor), got ${JSON.stringify(formatting.highlight)}.`,
    );
  }
  // `typeof` first: RegExp#test stringifies, so a JS caller's 123456 would pass.
  const color: unknown = formatting.color;
  if (color !== undefined && !(typeof color === "string" && HEX_COLOR.test(color))) {
    throw new RangeError(
      `color must be six hex digits (no "#") or "auto" (ST_HexColor), got ${typeof color === "string" ? JSON.stringify(color) : typeof color}.`,
    );
  }
  // ST_HpsMeasure (wml.xsd) as a number is ST_UnsignedDecimalNumber: a
  // non-negative integer, 0 included. Beyond MAX_SAFE_INTEGER a JS number
  // stops being an exact integer, so that is the practical upper bound.
  const size = formatting.fontSizeHalfPoints;
  if (size !== undefined && !(Number.isSafeInteger(size) && size >= 0)) {
    throw new RangeError(
      `fontSizeHalfPoints must be a non-negative integer (ST_HpsMeasure), got ${typeof size === "number" ? String(size) : typeof size}.`,
    );
  }
}

export interface RunFormatting {
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly strike?: boolean;
  readonly underline?: "single" | "double" | "thick" | "dotted" | "wave" | "none";
  /**
   * Six hex digits without a leading `#` (e.g. `"FF0000"`), or `"auto"`.
   * Writers throw a `RangeError` for anything else, before changing anything;
   * `getRunFormat` returns the file's value unvalidated.
   */
  readonly color?: string;
  /**
   * Text highlight. Writers (`appendTextRun`, `setRunFormat`,
   * `setParagraphText`, `setTableCellText`) accept only a
   * {@link HighlightColor} and throw a `RangeError` otherwise, before changing
   * anything. `getRunFormat` returns whatever the file holds, so a value
   * written by another tool is read back unvalidated (the same holds for
   * `color` and `fontSizeHalfPoints`).
   */
  readonly highlight?: string;
  /**
   * Font size in half-points (e.g. 24 = 12pt). Writers accept a non-negative
   * integer (0 is schema-valid) and throw a `RangeError` otherwise, before
   * changing anything.
   */
  readonly fontSizeHalfPoints?: number;
  /** Font family applied to ASCII / hAnsi runs. */
  readonly font?: string;
  /** Font family applied to East Asian (CJK) text. */
  readonly fontEastAsia?: string;
}

/** Append a styled text run to an existing paragraph and return it. */
export function appendTextRun(
  paragraph: WmlParagraph,
  text: string,
  formatting: RunFormatting = {},
): WmlRun {
  assertWritableRunFormatting(formatting);
  const pieces = splitTextIntoPieces(text);
  const rPrChildren: XmlElement[] = [];
  if (formatting.font || formatting.fontEastAsia) {
    const attrs: XmlAttr[] = [];
    if (formatting.font) {
      attrs.push(wmlAttr("ascii", formatting.font));
      attrs.push(wmlAttr("hAnsi", formatting.font));
    }
    if (formatting.fontEastAsia) attrs.push(wmlAttr("eastAsia", formatting.fontEastAsia));
    rPrChildren.push(wmlEmpty("rFonts", attrs));
  }
  if (formatting.bold) rPrChildren.push(wmlEmpty("b", []));
  if (formatting.italic) rPrChildren.push(wmlEmpty("i", []));
  if (formatting.strike) rPrChildren.push(wmlEmpty("strike", []));
  if (formatting.underline) rPrChildren.push(wmlEmpty("u", [wmlAttr("val", formatting.underline)]));
  if (formatting.color) rPrChildren.push(wmlEmpty("color", [wmlAttr("val", formatting.color)]));
  if (formatting.highlight)
    rPrChildren.push(wmlEmpty("highlight", [wmlAttr("val", formatting.highlight)]));
  if (formatting.fontSizeHalfPoints !== undefined) {
    rPrChildren.push(wmlEmpty("sz", [wmlAttr("val", String(formatting.fontSizeHalfPoints))]));
    rPrChildren.push(wmlEmpty("szCs", [wmlAttr("val", String(formatting.fontSizeHalfPoints))]));
  }
  const run: WmlRun = rPrChildren.length
    ? {
        kind: "run",
        rPr: {
          kind: "element",
          name: { uri: WML_NS, local: "rPr", prefix: "w" },
          attrs: [],
          children: rPrChildren,
          xmlSpace: "default",
          selfClosing: false,
        },
        pieces,
        extras: [],
      }
    : { kind: "run", pieces, extras: [] };
  paragraph.children.push(run);
  return run;
}

/**
 * Apply formatting to an existing run, replacing any matching `<w:rPr>`
 * children. Properties left `undefined` are not touched, so this is a
 * patch — call `clearRunFormat` first if you want a full reset.
 *
 * Boolean flags (`bold`, `italic`, `strike`) accept `false` to remove
 * the corresponding `<w:b/>` / `<w:i/>` / `<w:strike/>` element.
 * `underline: "none"` removes `<w:u/>`. Setting a value to a string or
 * number replaces (or inserts) the relevant rPr child while preserving
 * the run's other rPr children, including formatting the library
 * doesn't yet model.
 */
export function setRunFormat(run: WmlRun, formatting: RunFormatting): void {
  assertWritableRunFormatting(formatting);
  if (
    formatting.bold === undefined &&
    formatting.italic === undefined &&
    formatting.strike === undefined &&
    formatting.underline === undefined &&
    formatting.color === undefined &&
    formatting.highlight === undefined &&
    formatting.fontSizeHalfPoints === undefined &&
    formatting.font === undefined &&
    formatting.fontEastAsia === undefined
  ) {
    return;
  }
  const rPr: XmlElement = run.rPr ?? {
    kind: "element",
    name: { uri: WML_NS, local: "rPr", prefix: "w" },
    attrs: [],
    children: [],
    xmlSpace: "default",
    selfClosing: false,
  };
  const children = rPr.children as XmlElement[];

  const removeLocal = (local: string): void => {
    for (let i = children.length - 1; i >= 0; i--) {
      const c = children[i];
      if (c && c.kind === "element" && c.name.uri === WML_NS && c.name.local === local) {
        children.splice(i, 1);
      }
    }
  };
  const setEmpty = (local: string, attrs: XmlAttr[]): void => {
    removeLocal(local);
    children.push(wmlEmpty(local, attrs));
  };

  if (formatting.bold !== undefined) {
    if (formatting.bold) setEmpty("b", []);
    else removeLocal("b");
  }
  if (formatting.italic !== undefined) {
    if (formatting.italic) setEmpty("i", []);
    else removeLocal("i");
  }
  if (formatting.strike !== undefined) {
    if (formatting.strike) setEmpty("strike", []);
    else removeLocal("strike");
  }
  if (formatting.underline !== undefined) {
    if (formatting.underline === "none") removeLocal("u");
    else setEmpty("u", [wmlAttr("val", formatting.underline)]);
  }
  if (formatting.color !== undefined) {
    setEmpty("color", [wmlAttr("val", formatting.color)]);
  }
  if (formatting.highlight !== undefined) {
    setEmpty("highlight", [wmlAttr("val", formatting.highlight)]);
  }
  if (formatting.fontSizeHalfPoints !== undefined) {
    setEmpty("sz", [wmlAttr("val", String(formatting.fontSizeHalfPoints))]);
    setEmpty("szCs", [wmlAttr("val", String(formatting.fontSizeHalfPoints))]);
  }
  if (formatting.font !== undefined || formatting.fontEastAsia !== undefined) {
    // Build/replace <w:rFonts>. Preserve any attrs the caller didn't override.
    const existing = children.find(
      (c): c is XmlElement => c.kind === "element" && c.name.local === "rFonts",
    );
    const attrs: XmlAttr[] = existing ? existing.attrs.slice() : [];
    const setAttr = (local: string, value: string): void => {
      const idx = attrs.findIndex((a) => a.name.local === local);
      const next = wmlAttr(local, value);
      if (idx >= 0) attrs[idx] = next;
      else attrs.push(next);
    };
    if (formatting.font !== undefined) {
      setAttr("ascii", formatting.font);
      setAttr("hAnsi", formatting.font);
    }
    if (formatting.fontEastAsia !== undefined) {
      setAttr("eastAsia", formatting.fontEastAsia);
    }
    removeLocal("rFonts");
    children.push(wmlEmpty("rFonts", attrs));
  }

  if (!run.rPr && children.length > 0) {
    run.rPr = rPr;
  }
}

/**
 * Drop the run's entire `<w:rPr>` block, leaving the run text unformatted
 * (it picks up the document/paragraph default style).
 */
export function clearRunFormat(run: WmlRun): void {
  delete run.rPr;
}

// --- Generic run/paragraph property access ------------------------------------
//
// `setRunFormat` / `setParagraphAlignment` cover the common formatting. These
// generic helpers reach *any* on-off (`<w:x/>`) or single-value (`<w:x
// w:val="…"/>`) child of `<w:rPr>` / `<w:pPr>`, so callers (e.g. the editor's
// property commands) can toggle the long tail of WordprocessingML formatting
// (caps, smallCaps, vanish, keepNext, widowControl, outlineLvl, …) without a
// bespoke function per element. Complex children (rFonts, ind, spacing, borders,
// shading, tabs, numPr, framePr) keep their dedicated builders.

function ensureRPr(run: WmlRun): XmlElement {
  if (!run.rPr) {
    run.rPr = {
      kind: "element",
      name: { uri: WML_NS, local: "rPr", prefix: "w" },
      attrs: [],
      children: [],
      xmlSpace: "default",
      selfClosing: false,
    };
  }
  return run.rPr;
}

/** Remove every direct child named `local` from a properties element. */
function removePropChild(container: XmlElement, local: string): void {
  const children = container.children as XmlElement[];
  for (let i = children.length - 1; i >= 0; i--) {
    const c = children[i];
    if (c && c.kind === "element" && c.name.uri === WML_NS && c.name.local === local) {
      children.splice(i, 1);
    }
  }
}

// Table property containers are xsd:sequences; a child appended at the end
// can land out of schema order, which Word may reject as corrupt.
const ORDERED_CONTAINERS: Readonly<Record<string, readonly string[]>> = {
  tblPr: TBL_PR_ORDER,
  trPr: TR_PR_ORDER,
  tcPr: TC_PR_ORDER,
};

function putPropChild(container: XmlElement, child: XmlElement): void {
  const order =
    container.name.uri === WML_NS ? ORDERED_CONTAINERS[container.name.local] : undefined;
  if (order) upsertWChild(container, child, order);
  else (container.children as XmlElement[]).push(child);
}

/** Add or remove an on-off property element (`<w:b/>`, `<w:caps/>`, …). */
function setOnOffChild(container: XmlElement, local: string, on: boolean): void {
  removePropChild(container, local);
  if (on) putPropChild(container, wmlEmpty(local, []));
}

/** Set a single-value property element (`<w:x w:val="…"/>`); undefined removes it. */
function setValChild(container: XmlElement, local: string, val: string | undefined): void {
  removePropChild(container, local);
  if (val !== undefined) putPropChild(container, wmlEmpty(local, [wmlAttr("val", val)]));
}

/** Read whether a property element is present and its `w:val`, if any. */
function readProp(
  container: XmlElement | undefined,
  local: string,
): { present: boolean; val?: string } {
  if (!container) return { present: false };
  for (const c of container.children) {
    if (c.kind === "element" && c.name.uri === WML_NS && c.name.local === local) {
      const val = c.attrs.find((a) => a.name.uri === WML_NS && a.name.local === "val")?.value;
      return val !== undefined ? { present: true, val } : { present: true };
    }
  }
  return { present: false };
}

/** Toggle an on-off `<w:rPr>` child (creates `<w:rPr>` if absent). */
export function setRunOnOff(run: WmlRun, local: string, on: boolean): void {
  setOnOffChild(ensureRPr(run), local, on);
}

/** Set a single-value `<w:rPr>` child; `undefined` removes it. */
export function setRunValProp(run: WmlRun, local: string, val: string | undefined): void {
  setValChild(ensureRPr(run), local, val);
}

/** Read an `<w:rPr>` child's presence / value. */
export function getRunProp(run: WmlRun, local: string): { present: boolean; val?: string } {
  return readProp(run.rPr, local);
}

/** Toggle an on-off `<w:pPr>` child (creates `<w:pPr>` if absent). */
export function setParagraphOnOff(p: WmlParagraph, local: string, on: boolean): void {
  setOnOffChild(ensurePPr(p), local, on);
}

/** Set a single-value `<w:pPr>` child; `undefined` removes it. */
export function setParagraphValProp(p: WmlParagraph, local: string, val: string | undefined): void {
  setValChild(ensurePPr(p), local, val);
}

/** Read a `<w:pPr>` child's presence / value. */
export function getParagraphProp(
  p: WmlParagraph,
  local: string,
): { present: boolean; val?: string } {
  return readProp(p.pPr, local);
}

// --- Container-level generic property access ----------------------------------
//
// The same on-off / single-value machinery, but operating on an arbitrary
// properties element (`<w:tcPr>`, `<w:trPr>`, `<w:tblPr>`, `<w:sectPr>`, `<w:lvl>`
// …) passed in by the caller. The editor uses these for the long tail of table,
// row, cell, and section formatting, ensuring the container exists first.

/** Toggle an on-off child on any properties element. */
export function setElementOnOff(container: XmlElement, local: string, on: boolean): void {
  setOnOffChild(container, local, on);
}

/** Set a single-value child on any properties element; `undefined` removes it. */
export function setElementValProp(
  container: XmlElement,
  local: string,
  val: string | undefined,
): void {
  setValChild(container, local, val);
}

/** Read a child's presence / value from any properties element. */
export function getElementProp(
  container: XmlElement | undefined,
  local: string,
): { present: boolean; val?: string } {
  return readProp(container, local);
}

/** Build an empty properties element (`<w:tcPr/>`, `<w:sectPr/>`, …). */
export function makePropsElement(local: string): XmlElement {
  return {
    kind: "element",
    name: { uri: WML_NS, local, prefix: "w" },
    attrs: [],
    children: [],
    xmlSpace: "default",
    selfClosing: false,
  };
}

// --- Raw XML node editing -----------------------------------------------------
//
// The universal escape hatch: set/read any attribute or child on any element in
// the AST. The editor's raw-XML inspector uses these to make every element —
// including DrawingML / OMML / VML that live inline in document.xml — editable,
// without a bespoke command per OOXML element.

/** Set (or, with `undefined`, remove) an attribute by local name on any element. */
export function setElementAttr(el: XmlElement, local: string, value: string | undefined): void {
  const attrs = el.attrs as XmlAttr[];
  const index = attrs.findIndex((a) => a.name.local === local);
  if (value === undefined) {
    if (index >= 0) attrs.splice(index, 1);
    return;
  }
  if (index >= 0) {
    (attrs[index] as { value: string }).value = value;
  } else {
    attrs.push({ name: { uri: "", local, prefix: "" }, value, isNamespaceDecl: false });
  }
}

/** Read an attribute value by local name from any element. */
export function getElementAttr(el: XmlElement, local: string): string | undefined {
  return el.attrs.find((a) => a.name.local === local)?.value;
}

/** The direct child elements of an element (text/comment nodes filtered out). */
export function childElementsOf(el: XmlElement): XmlElement[] {
  return el.children.filter((c): c is XmlElement => c.kind === "element");
}

/** Append a child element (`<w:local/>` in the WML namespace) and return it. */
export function appendChildElement(parent: XmlElement, local: string): XmlElement {
  const child = makePropsElement(local);
  (parent.children as XmlElement[]).push(child);
  return child;
}

// Half-points per unit of ST_PositiveUniversalMeasure, as exact fractions
// [numerator, denominator]: 1pt = 2, 1pc = 1pi = 12pt, 1in = 72pt,
// 1cm = 72 / 2.54 pt, 1mm = 72 / 25.4 pt.
const HALF_POINTS_PER_UNIT: Readonly<Record<string, readonly [bigint, bigint]>> = {
  pt: [2n, 1n],
  pc: [24n, 1n],
  pi: [24n, 1n],
  in: [144n, 1n],
  cm: [7200n, 127n],
  mm: [720n, 127n],
};
const HALF_POINT_COUNT = /^\+?[0-9]+$/;
const UNIVERSAL_MEASURE = /^([0-9]+)(?:\.([0-9]+))?(mm|cm|in|pt|pc|pi)$/;

/**
 * An ST_HpsMeasure value as half-points: a plain count, or a universal measure
 * ("12pt", "1in", …) converted exactly. `undefined` when the value is invalid
 * or is not a whole number of half-points ("3mm"); the XML keeps it as is.
 */
function halfPointsOf(val: string): number | undefined {
  let exact: bigint | undefined;
  if (HALF_POINT_COUNT.test(val)) exact = BigInt(val);
  else {
    const m = UNIVERSAL_MEASURE.exec(val);
    const unit = m && HALF_POINTS_PER_UNIT[m[3]!];
    if (!m || !unit) return undefined;
    const fraction = m[2] ?? "";
    const numerator = BigInt(m[1]! + fraction) * unit[0];
    const denominator = 10n ** BigInt(fraction.length) * unit[1];
    if (numerator % denominator !== 0n) return undefined;
    exact = numerator / denominator;
  }
  return exact <= BigInt(Number.MAX_SAFE_INTEGER) ? Number(exact) : undefined;
}

/**
 * Read back the formatting a run currently has on its `<w:rPr>`. Returns a
 * `RunFormatting` with only the keys that are actually present, so callers
 * can round-trip via `setRunFormat(other, getRunFormat(run))`. Values are
 * returned as the file holds them, unvalidated: a `highlight` outside
 * {@link HIGHLIGHT_COLORS} or a malformed `color` (written by another tool) is
 * read back as-is, and that copy is rejected by `setRunFormat` rather than
 * spread further. A font size given in units (`w:sz="12pt"`) is converted to
 * half-points; one that is not a whole number of half-points is left out.
 */
export function getRunFormat(run: WmlRun): RunFormatting {
  // Build a writable scratch object; only keys we actually observe will
  // be set, then cast to the readonly `RunFormatting` on return.
  const out: Record<string, unknown> = {};
  if (!run.rPr) return out as RunFormatting;
  for (const child of run.rPr.children) {
    if (child.kind !== "element" || child.name.uri !== WML_NS) continue;
    const local = child.name.local;
    if (local === "b") out.bold = true;
    else if (local === "i") out.italic = true;
    else if (local === "strike") out.strike = true;
    else if (local === "u") {
      const v = child.attrs.find((a) => a.name.local === "val")?.value;
      if (
        v === "single" ||
        v === "double" ||
        v === "thick" ||
        v === "dotted" ||
        v === "wave" ||
        v === "none"
      ) {
        out.underline = v;
      }
    } else if (local === "color") {
      const v = child.attrs.find((a) => a.name.local === "val")?.value;
      if (v !== undefined) out.color = v;
    } else if (local === "highlight") {
      const v = child.attrs.find((a) => a.name.local === "val")?.value;
      if (v !== undefined) out.highlight = v;
    } else if (local === "sz") {
      const v = child.attrs.find((a) => a.name.local === "val")?.value;
      const n = v !== undefined ? halfPointsOf(v) : undefined;
      if (n !== undefined) out.fontSizeHalfPoints = n;
    } else if (local === "rFonts") {
      const ascii = child.attrs.find((a) => a.name.local === "ascii")?.value;
      const east = child.attrs.find((a) => a.name.local === "eastAsia")?.value;
      if (ascii !== undefined) out.font = ascii;
      if (east !== undefined) out.fontEastAsia = east;
    }
  }
  return out as RunFormatting;
}

/**
 * Collapse runs in `paragraph` that are immediately adjacent and share
 * the same `<w:rPr>`. Only runs whose pieces are entirely text /
 * delText / tab / break (no drawings, instrText, etc.) are eligible —
 * special pieces stop the merge boundary. Other inline children
 * (`<w:hyperlink>`, raw passthroughs, …) interrupt adjacency.
 *
 * Returns the number of merges performed. The merged run is the earlier
 * one — the later run is removed and its pieces appended.
 *
 * Useful after templating has fragmented a paragraph into many tiny
 * runs that share the same formatting (Word's spell-checker is the
 * common culprit).
 */
export function mergeAdjacentRuns(paragraph: WmlParagraph): number {
  let merges = 0;
  const out: typeof paragraph.children = [];
  for (const child of paragraph.children) {
    const prev = out[out.length - 1];
    if (
      child.kind === "run" &&
      prev &&
      prev.kind === "run" &&
      isMergeableRun(child) &&
      isMergeableRun(prev) &&
      prev.revision === child.revision &&
      sameRPr(prev, child)
    ) {
      prev.pieces.push(...child.pieces);
      merges++;
      continue;
    }
    out.push(child);
  }
  paragraph.children = out;
  return merges;
}

function isSafeRunPiece(p: WmlRunPiece): boolean {
  return p.kind === "text" || p.kind === "delText" || p.kind === "tab" || p.kind === "break";
}

function isMergeableRun(r: WmlRun): boolean {
  return r.extras.length === 0 && r.pieces.every(isSafeRunPiece);
}

function sameRPr(a: WmlRun, b: WmlRun): boolean {
  if (a.rPr === undefined && b.rPr === undefined) return true;
  if (a.rPr === undefined || b.rPr === undefined) return false;
  return elementsEqual(a.rPr, b.rPr);
}

function elementsEqual(a: XmlElement, b: XmlElement): boolean {
  if (a.name.uri !== b.name.uri || a.name.local !== b.name.local) return false;
  if (a.attrs.length !== b.attrs.length) return false;
  // Attribute order isn't semantic, but for rPr comparison Word treats
  // ordering as cosmetic — we compare attrs as an unordered set keyed by
  // namespaced name + value.
  const aKeys = a.attrs.map((x) => `${x.name.uri}|${x.name.local}=${x.value}`).toSorted();
  const bKeys = b.attrs.map((x) => `${x.name.uri}|${x.name.local}=${x.value}`).toSorted();
  for (let i = 0; i < aKeys.length; i++) if (aKeys[i] !== bKeys[i]) return false;
  if (a.children.length !== b.children.length) return false;
  for (let i = 0; i < a.children.length; i++) {
    const ca = a.children[i];
    const cb = b.children[i];
    if (!ca || !cb || ca.kind !== cb.kind) return false;
    if (ca.kind === "element" && cb.kind === "element") {
      if (!elementsEqual(ca, cb)) return false;
    } else if (ca.kind === "text" && cb.kind === "text") {
      if (ca.value !== cb.value) return false;
    } else if (ca.kind === "cdata" && cb.kind === "cdata") {
      if (ca.value !== cb.value) return false;
    } else {
      // Comments / PIs in rPr are exotic enough that we treat them as
      // "not equal" rather than try to compare in detail.
      return false;
    }
  }
  return true;
}

export type ParagraphAlignment = "left" | "center" | "right" | "both" | "distribute";

/** Set the paragraph's `<w:jc w:val="..."/>` justification. */
export function setParagraphAlignment(p: WmlParagraph, alignment: ParagraphAlignment): void {
  const pPr = ensurePPr(p);
  const idx = pPr.children.findIndex(
    (c) => c.kind === "element" && c.name.uri === WML_NS && c.name.local === "jc",
  );
  const newEl = wmlEmpty("jc", [wmlAttr("val", alignment)]);
  if (idx >= 0) (pPr.children as XmlElement[])[idx] = newEl;
  else (pPr.children as XmlElement[]).push(newEl);
}

export interface ParagraphIndent {
  /** Left indent in twips. */
  readonly left?: number;
  /** Right indent in twips. */
  readonly right?: number;
  /** First-line indent in twips. */
  readonly firstLine?: number;
  /** Hanging indent in twips. */
  readonly hanging?: number;
  /**
   * The same indents in hundredths of a character (`w:leftChars` …,
   * §17.3.1.12), as Japanese and Chinese Word set them ("2 字"). Where both
   * are given, the character value wins.
   */
  readonly leftChars?: number;
  readonly rightChars?: number;
  readonly firstLineChars?: number;
  readonly hangingChars?: number;
}

const INDENT_ATTRS = [
  "left",
  "leftChars",
  "right",
  "rightChars",
  "hanging",
  "hangingChars",
  "firstLine",
  "firstLineChars",
] as const satisfies readonly (keyof ParagraphIndent)[];

/** Set `<w:ind>` on the paragraph's pPr. */
export function setParagraphIndent(p: WmlParagraph, indent: ParagraphIndent): void {
  const pPr = ensurePPr(p);
  const attrs: XmlAttr[] = [];
  for (const name of INDENT_ATTRS) {
    const value = indent[name];
    if (value !== undefined) attrs.push(wmlAttr(name, String(value)));
  }
  const idx = pPr.children.findIndex(
    (c) => c.kind === "element" && c.name.uri === WML_NS && c.name.local === "ind",
  );
  const newEl = wmlEmpty("ind", attrs);
  if (idx >= 0) (pPr.children as XmlElement[])[idx] = newEl;
  else (pPr.children as XmlElement[]).push(newEl);
}

export interface ParagraphSpacing {
  /** Space before the paragraph in twips. */
  readonly before?: number;
  /** Space after the paragraph in twips. */
  readonly after?: number;
  /** Line spacing value in twips. */
  readonly line?: number;
  /** Line spacing rule: `"auto"`, `"exact"`, or `"atLeast"`. */
  readonly lineRule?: "auto" | "exact" | "atLeast";
  /**
   * Space before / after in hundredths of a line (`w:beforeLines`,
   * `w:afterLines`, §17.3.1.33), as East Asian Word sets them ("0.5 行").
   * Where both are given, the line value wins.
   */
  readonly beforeLines?: number;
  readonly afterLines?: number;
}

/** Set `<w:spacing>` on the paragraph's pPr. */
export function setParagraphSpacing(p: WmlParagraph, spacing: ParagraphSpacing): void {
  const pPr = ensurePPr(p);
  const attrs: XmlAttr[] = [];
  if (spacing.before !== undefined) attrs.push(wmlAttr("before", String(spacing.before)));
  if (spacing.beforeLines !== undefined) {
    attrs.push(wmlAttr("beforeLines", String(spacing.beforeLines)));
  }
  if (spacing.after !== undefined) attrs.push(wmlAttr("after", String(spacing.after)));
  if (spacing.afterLines !== undefined) {
    attrs.push(wmlAttr("afterLines", String(spacing.afterLines)));
  }
  if (spacing.line !== undefined) attrs.push(wmlAttr("line", String(spacing.line)));
  if (spacing.lineRule !== undefined) attrs.push(wmlAttr("lineRule", spacing.lineRule));
  const idx = pPr.children.findIndex(
    (c) => c.kind === "element" && c.name.uri === WML_NS && c.name.local === "spacing",
  );
  const newEl = wmlEmpty("spacing", attrs);
  if (idx >= 0) (pPr.children as XmlElement[])[idx] = newEl;
  else (pPr.children as XmlElement[]).push(newEl);
}

export interface ParagraphBordersOptions {
  /** Stroke style — defaults to `"single"`. Pass `"none"` to clear borders. */
  readonly style?: TableBorderStyle;
  /** Stroke width in eighths of a point. Defaults to 4 (½ pt). */
  readonly sizeEighthsOfPoint?: number;
  /** Hex RGB without a leading `#`. Defaults to `"auto"`. */
  readonly color?: string;
  /**
   * Distance in points (`w:space`) between text and border. Defaults to 4.
   * Valid range is 0–31 per the spec; values outside are clamped silently
   * by Word.
   */
  readonly spacePt?: number;
  /**
   * Sides to apply. Defaults to all four (`top`, `left`, `bottom`, `right`)
   * — pass a subset to draw, eg, only a bottom rule under headings.
   */
  readonly sides?: ReadonlyArray<"top" | "left" | "bottom" | "right">;
}

/**
 * Set borders on a paragraph's pPr. Replaces any existing
 * `<w:pBdr>` block.
 *
 * Common patterns:
 *
 * - Bottom rule under a heading:
 *   `setParagraphBorders(p, { sides: ["bottom"], sizeEighthsOfPoint: 12 });`
 *
 * - Full box around a callout paragraph:
 *   `setParagraphBorders(p, { color: "808080" });`
 *
 * - Clear an existing border block:
 *   `setParagraphBorders(p, { style: "none" });`
 */
export function setParagraphBorders(
  paragraph: WmlParagraph,
  options: ParagraphBordersOptions = {},
): void {
  const pPr = ensurePPr(paragraph);
  const style = options.style ?? "single";
  const sz = String(options.sizeEighthsOfPoint ?? 4);
  const color = options.color ?? "auto";
  const space = String(options.spacePt ?? 4);
  const sides = options.sides ?? ["top", "left", "bottom", "right"];
  const children = pPr.children as XmlElement[];
  for (let i = children.length - 1; i >= 0; i--) {
    const c = children[i];
    if (c && c.kind === "element" && c.name.uri === WML_NS && c.name.local === "pBdr") {
      children.splice(i, 1);
    }
  }
  const sideAttrs = (): XmlAttr[] => [
    wmlAttr("val", style),
    wmlAttr("sz", sz),
    wmlAttr("space", space),
    wmlAttr("color", color),
  ];
  const pBdr: XmlElement = {
    kind: "element",
    name: { uri: WML_NS, local: "pBdr", prefix: "w" },
    attrs: [],
    children: sides.map((s) => wmlEmpty(s, sideAttrs())),
    xmlSpace: "default",
    selfClosing: false,
  };
  // pBdr appears after pStyle and numPr per the schema; insert after the
  // last "structural" pPr child if any are present.
  const tailLocals = new Set(["pStyle", "keepNext", "keepLines", "numPr"]);
  let insertAt = 0;
  for (let i = 0; i < children.length; i++) {
    const c = children[i];
    if (c && c.kind === "element" && tailLocals.has(c.name.local)) insertAt = i + 1;
  }
  children.splice(insertAt, 0, pBdr);
}

/** Every `w:shd/@w:val` pattern: ST_Shd (ECMA-376 Part 1 §17.18.78). */
export const SHADING_PATTERNS = [
  "nil",
  "clear",
  "solid",
  "horzStripe",
  "vertStripe",
  "reverseDiagStripe",
  "diagStripe",
  "horzCross",
  "diagCross",
  "thinHorzStripe",
  "thinVertStripe",
  "thinReverseDiagStripe",
  "thinDiagStripe",
  "thinHorzCross",
  "thinDiagCross",
  "pct5",
  "pct10",
  "pct12",
  "pct15",
  "pct20",
  "pct25",
  "pct30",
  "pct35",
  "pct37",
  "pct40",
  "pct45",
  "pct50",
  "pct55",
  "pct60",
  "pct62",
  "pct65",
  "pct70",
  "pct75",
  "pct80",
  "pct85",
  "pct87",
  "pct90",
  "pct95",
] as const;

export type ShadingPattern = (typeof SHADING_PATTERNS)[number];

/** Every `w:themeColor` / `w:themeFill` value: ST_ThemeColor (§17.18.97). */
export const THEME_COLORS = [
  "dark1",
  "light1",
  "dark2",
  "light2",
  "accent1",
  "accent2",
  "accent3",
  "accent4",
  "accent5",
  "accent6",
  "hyperlink",
  "followedHyperlink",
  "none",
  "background1",
  "text1",
  "background2",
  "text2",
] as const;

export type ThemeColor = (typeof THEME_COLORS)[number];

/** Shading of a paragraph or run (`<w:shd>`, §17.3.1.31 / §17.3.2.32). */
export interface ShadingOptions {
  /**
   * Hex RGB fill colour, or `"auto"` (the default). With `themeFill` this is
   * the theme colour's resolved value, which consumers without the theme use.
   */
  readonly fill?: string;
  /** Pattern overlaid on the fill. Defaults to `"clear"` (flat fill). */
  readonly pattern?: ShadingPattern;
  /** Pattern stroke colour. Defaults to `"auto"`. */
  readonly color?: string;
  /** The theme colour the fill comes from (`w:themeFill`). */
  readonly themeFill?: ThemeColor;
  /** Tint applied to `themeFill`, 0–255 (`w:themeFillTint`). */
  readonly themeFillTint?: number;
  /** Shade applied to `themeFill`, 0–255 (`w:themeFillShade`). */
  readonly themeFillShade?: number;
}

/** Shading of a paragraph; the same shape as run shading. */
export type ParagraphShadingOptions = ShadingOptions;

const THEME_COLOR_SET: ReadonlySet<string> = new Set(THEME_COLORS);

/**
 * A theme tint / shade as ST_UcharHexNumber: two hex digits. Throws a
 * `RangeError` for anything that is not an integer in 0–255.
 */
export function ucharHex(value: number, what: string): string {
  if (!Number.isInteger(value) || value < 0 || value > 255) {
    throw new RangeError(`${what} must be an integer in 0–255, got ${String(value)}.`);
  }
  return value.toString(16).toUpperCase().padStart(2, "0");
}

/** Throw a `RangeError` unless `value` is an ST_ThemeColor. */
export function assertThemeColor(value: string, what: string): void {
  if (!THEME_COLOR_SET.has(value)) {
    throw new RangeError(
      `${what} must be one of ${THEME_COLORS.join(", ")}, got ${JSON.stringify(value)}.`,
    );
  }
}

/** Build a `<w:shd>` element; the theme attributes are validated. */
export function buildShading(options: ShadingOptions): XmlElement {
  const attrs = [
    wmlAttr("val", options.pattern ?? "clear"),
    wmlAttr("color", options.color ?? "auto"),
    wmlAttr("fill", options.fill ?? "auto"),
  ];
  if (options.themeFill !== undefined) {
    assertThemeColor(options.themeFill, "themeFill");
    attrs.push(wmlAttr("themeFill", options.themeFill));
  }
  if (options.themeFillTint !== undefined)
    attrs.push(wmlAttr("themeFillTint", ucharHex(options.themeFillTint, "themeFillTint")));
  if (options.themeFillShade !== undefined)
    attrs.push(wmlAttr("themeFillShade", ucharHex(options.themeFillShade, "themeFillShade")));
  return wmlEmpty("shd", attrs);
}

/** Apply background shading to a paragraph (Word's "highlight" — but applied
 * to the whole paragraph rather than a run). Replaces any existing
 * `<w:shd>` on pPr.
 */
export function setParagraphShading(paragraph: WmlParagraph, options: ShadingOptions = {}): void {
  const shd = buildShading(options);
  const pPr = ensurePPr(paragraph);
  removePropChild(pPr, "shd");
  (pPr.children as XmlElement[]).push(shd);
}

function ensurePPr(p: WmlParagraph): XmlElement {
  if (p.pPr) return p.pPr;
  const pPr: XmlElement = {
    kind: "element",
    name: { uri: WML_NS, local: "pPr", prefix: "w" },
    attrs: [],
    children: [],
    xmlSpace: "default",
    selfClosing: false,
  };
  p.pPr = pPr;
  return pPr;
}

/** Read `<w:pStyle w:val>` from a paragraph's pPr, or `undefined`. */
export function getParagraphStyle(p: WmlParagraph): string | undefined {
  return refChildVal(p.pPr, "pStyle");
}

/**
 * Set or clear a paragraph's `<w:pStyle w:val="…">`. Pass a styleId
 * (e.g. `"Heading1"`) to apply, or `undefined` to drop the style
 * reference altogether (paragraph falls back to Normal).
 *
 * Note that this only writes the `<w:pStyle>` reference; the styles
 * part is unaffected. Use {@link addStyle} (in `@office-kit/docx`) first
 * if you're applying a custom style that doesn't exist yet.
 */
export function setParagraphStyle(p: WmlParagraph, styleId: string | undefined): void {
  if (styleId === undefined) {
    if (!p.pPr) return;
    const children = p.pPr.children as XmlElement[];
    for (let i = children.length - 1; i >= 0; i--) {
      const c = children[i];
      if (c && c.kind === "element" && c.name.uri === WML_NS && c.name.local === "pStyle") {
        children.splice(i, 1);
      }
    }
    return;
  }
  const pPr = ensurePPr(p);
  const children = pPr.children as XmlElement[];
  const idx = children.findIndex(
    (c) => c.kind === "element" && c.name.uri === WML_NS && c.name.local === "pStyle",
  );
  const newEl = wmlEmpty("pStyle", [wmlAttr("val", styleId)]);
  if (idx >= 0) children[idx] = newEl;
  else children.unshift(newEl); // pStyle conventionally appears first.
}

/** Read `<w:jc w:val>` from a paragraph's pPr as a typed value, or `undefined`. */
export function getParagraphAlignment(p: WmlParagraph): ParagraphAlignment | undefined {
  const v = refChildVal(p.pPr, "jc");
  if (v === "left" || v === "center" || v === "right" || v === "both" || v === "distribute")
    return v;
  return undefined;
}

/** Read the numbering reference (`<w:numPr>` → `numId` / `ilvl`) if present. */
export function getParagraphNumbering(
  p: WmlParagraph,
): { numId: number; ilvl: number } | undefined {
  if (!p.pPr) return undefined;
  const numPr = p.pPr.children.find(
    (c) => c.kind === "element" && c.name.uri === WML_NS && c.name.local === "numPr",
  );
  if (!numPr || numPr.kind !== "element") return undefined;
  const numIdEl = numPr.children.find(
    (c) => c.kind === "element" && c.name.uri === WML_NS && c.name.local === "numId",
  );
  const ilvlEl = numPr.children.find(
    (c) => c.kind === "element" && c.name.uri === WML_NS && c.name.local === "ilvl",
  );
  if (!numIdEl || numIdEl.kind !== "element") return undefined;
  const numIdAttr = numIdEl.attrs.find((a) => a.name.uri === WML_NS && a.name.local === "val");
  if (!numIdAttr) return undefined;
  const numIdVal = Number.parseInt(numIdAttr.value, 10);
  const ilvlAttr =
    ilvlEl?.kind === "element"
      ? ilvlEl.attrs.find((a) => a.name.uri === WML_NS && a.name.local === "val")
      : undefined;
  const ilvlVal = ilvlAttr ? Number.parseInt(ilvlAttr.value, 10) : 0;
  if (!Number.isFinite(numIdVal)) return undefined;
  return { numId: numIdVal, ilvl: Number.isFinite(ilvlVal) ? ilvlVal : 0 };
}

/**
 * Replace a paragraph's content with a single styled text run. Existing
 * runs and inline children are removed; existing pPr is kept.
 */
export function setParagraphText(
  p: WmlParagraph,
  text: string,
  formatting: RunFormatting = {},
): void {
  assertWritableRunFormatting(formatting);
  p.children = [];
  appendTextRun(p, text, formatting);
}

function refChildVal(parent: XmlElement | undefined, local: string): string | undefined {
  if (!parent) return undefined;
  for (const c of parent.children) {
    if (c.kind === "element" && c.name.uri === WML_NS && c.name.local === local) {
      const attr = c.attrs.find((a) => a.name.uri === WML_NS && a.name.local === "val");
      return attr?.value;
    }
  }
  return undefined;
}

/**
 * Replace the text of a specific cell in a table with `text`. The cell's
 * existing paragraphs are replaced with a single paragraph containing one
 * styled text run. Row / column indices are 0-based; throws if out of
 * range.
 */
export function setTableCellText(
  table: WmlTable,
  row: number,
  col: number,
  text: string,
  formatting: RunFormatting = {},
): void {
  assertWritableRunFormatting(formatting);
  const tableRow = table.rows[row];
  if (!tableRow) {
    throw new Error(
      `setTableCellText: row ${row} is out of range (table has ${table.rows.length})`,
    );
  }
  const cell = tableRow.cells[col];
  if (!cell) {
    throw new Error(
      `setTableCellText: column ${col} is out of range (row ${row} has ${tableRow.cells.length})`,
    );
  }
  const paragraph = buildTextParagraph(text);
  cell.paragraphs = [paragraph];
  if (Object.keys(formatting).length > 0) {
    // appendTextRun already pushed the plain run; we need to wrap with the
    // styled run instead. Rebuild via appendTextRun on a fresh paragraph.
    paragraph.children = [];
    appendTextRun(paragraph, text, formatting);
  }
}

/**
 * Read the plain text of a single table cell (joining its paragraphs with
 * newlines). Throws if `row` / `col` is out of range.
 */
export function getTableCellText(table: WmlTable, row: number, col: number): string {
  const tableRow = table.rows[row];
  if (!tableRow) {
    throw new Error(`getTableCellText: row ${row} is out of range`);
  }
  const cell = tableRow.cells[col];
  if (!cell) {
    throw new Error(`getTableCellText: column ${col} is out of range`);
  }
  return cell.paragraphs
    .map((p) => {
      let acc = "";
      for (const child of p.children) {
        if (child.kind !== "run") continue;
        for (const piece of child.pieces) {
          if (piece.kind === "text" || piece.kind === "delText") acc += piece.value;
        }
      }
      return acc;
    })
    .join("\n");
}

/**
 * Append a new row at the end of a table. The new row's cell count
 * matches the existing rows. If `texts` is shorter, the remaining cells
 * are empty.
 */
export function appendTableRow(table: WmlTable, texts: readonly string[]): WmlTableRow {
  const expectedCols = table.rows[0]?.cells.length ?? texts.length;
  const cells: WmlTableCell[] = [];
  for (let c = 0; c < expectedCols; c++) {
    const text = texts[c] ?? "";
    cells.push({
      paragraphs: [buildTextParagraph(text)],
      extras: [],
    });
  }
  const row: WmlTableRow = { cells, extras: [] };
  table.rows.push(row);
  return row;
}

/**
 * Remove the row at `index` from a table (0-based). Returns true if
 * removed.
 */
export function removeTableRow(table: WmlTable, index: number): boolean {
  if (index < 0 || index >= table.rows.length) return false;
  table.rows.splice(index, 1);
  // A vertical merge that started in the removed row must restart below it.
  normalizeVerticalMerges(table);
  return true;
}

function wmlEmpty(local: string, attrs: XmlAttr[]): XmlElement {
  return {
    kind: "element",
    name: { uri: WML_NS, local, prefix: "w" },
    attrs,
    children: [],
    xmlSpace: "default",
    selfClosing: true,
  };
}

function wmlAttr(local: string, value: string): XmlAttr {
  return {
    name: { uri: WML_NS, local, prefix: "w" },
    value,
    isNamespaceDecl: false,
  };
}
