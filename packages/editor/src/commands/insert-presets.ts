/**
 * The built-in designs behind Insert ▸ Cover Page, Header and Footer. Word's
 * own gallery entries are building blocks shipped with Office (text boxes and
 * content controls bound to document properties); these presets reproduce
 * their look with ordinary paragraphs — borders, shading, theme-like colors —
 * and live fields (TITLE, SUBJECT, AUTHOR, DATE, PAGE) where Word binds a
 * property, so the content is plain WordprocessingML any consumer can render.
 */

import {
  buildComplexField,
  coreProperties,
  type Docx,
  setParagraphAlignment,
  setParagraphBorders,
  setParagraphShading,
  setParagraphSpacing,
  setParagraphStyle,
  setRunFormat,
  type RunFormatting,
  type WmlInline,
  type WmlParagraph,
  type WmlRun,
} from "@office-kit/docx";

/** Word's gallery placeholder text for an empty header / footer slot. */
export const TYPE_HERE = "[Type here]";

// Office theme accent colors (Accent 1 blue, Text 2 navy, Accent 2 orange).
const ACCENT_BLUE = "156082";
const TEXT_NAVY = "0E2841";
const ACCENT_ORANGE = "E97132";
const LIGHT_FILL = "DCEAF7";
const WHITE = "FFFFFF";

function run(text: string, format: RunFormatting = {}): WmlRun {
  const r: WmlRun = {
    kind: "run",
    pieces: text.split("\t").flatMap((part, i) => [
      ...(i > 0 ? [{ kind: "tab" } as const] : []),
      ...(part ? [{ kind: "text", value: part, preserveSpace: /^\s|\s$/.test(part) } as const] : []),
    ]),
    extras: [],
  };
  if (Object.keys(format).length) setRunFormat(r, format);
  return r;
}

/** A complex field whose shown result is `result`, with the given formatting. */
function field(instruction: string, result: string, format: RunFormatting = {}): WmlRun[] {
  const runs = buildComplexField(instruction, result);
  if (Object.keys(format).length) for (const r of runs) setRunFormat(r, format);
  return runs;
}

function paragraph(children: WmlInline[], style?: string): WmlParagraph {
  const p: WmlParagraph = { kind: "paragraph", children, extras: [] };
  if (style) setParagraphStyle(p, style);
  return p;
}

/** Property-bound text: the property's value, or Word's bracketed prompt. */
function propertyField(
  doc: Docx,
  type: "TITLE" | "SUBJECT" | "AUTHOR",
  prompt: string,
  format: RunFormatting = {},
): WmlRun[] {
  const core = coreProperties(doc);
  const value = type === "TITLE" ? core.title : type === "SUBJECT" ? core.subject : core.creator;
  return field(type, value || prompt, format);
}

// --- headers and footers ----------------------------------------------------------

export const HEADER_FOOTER_PRESETS = [
  "blank",
  "blankThreeColumns",
  "austin",
  "banded",
  "facet",
  "ion",
  "retrospect",
] as const;
export type HeaderFooterPresetId = (typeof HEADER_FOOTER_PRESETS)[number];

const HEADER_STYLE = "Header";
const FOOTER_STYLE = "Footer";

function pageField(format: RunFormatting = {}): WmlRun[] {
  return field("PAGE", "1", format);
}

/** The paragraphs of a header (`kind: "header"`) or footer preset. */
export function headerFooterPreset(
  doc: Docx,
  kind: "header" | "footer",
  preset: HeaderFooterPresetId,
): WmlParagraph[] {
  const style = kind === "header" ? HEADER_STYLE : FOOTER_STYLE;
  switch (preset) {
    case "blank":
      return [paragraph([run(TYPE_HERE)], style)];
    case "blankThreeColumns":
      // The Header / Footer style's center and right tab stops lay out the
      // three slots across the text width.
      return [paragraph([run(`${TYPE_HERE}\t${TYPE_HERE}\t${TYPE_HERE}`)], style)];
    case "austin": {
      const p =
        kind === "header"
          ? paragraph(propertyField(doc, "TITLE", "[Document title]", { color: ACCENT_BLUE }), style)
          : paragraph(pageField({ color: ACCENT_BLUE }), style);
      setParagraphBorders(p, {
        style: "single",
        sizeEighthsOfPoint: 12,
        color: ACCENT_BLUE,
        spacePt: 4,
        sides: [kind === "header" ? "bottom" : "top"],
      });
      setParagraphAlignment(p, kind === "header" ? "right" : "center");
      return [p];
    }
    case "banded": {
      const p =
        kind === "header"
          ? paragraph(propertyField(doc, "TITLE", "[Document title]", { color: WHITE, bold: true }), style)
          : paragraph(pageField({ color: WHITE, bold: true }), style);
      setParagraphShading(p, { fill: ACCENT_BLUE });
      setParagraphAlignment(p, "center");
      return [p];
    }
    case "facet": {
      if (kind === "header") {
        const p = paragraph(propertyField(doc, "TITLE", "[Document title]", { color: TEXT_NAVY }), style);
        setParagraphAlignment(p, "right");
        return [p];
      }
      const p = paragraph(
        [
          ...propertyField(doc, "TITLE", "[Document title]", { color: TEXT_NAVY }),
          run(" | "),
          ...propertyField(doc, "SUBJECT", "[Document subtitle]", { color: TEXT_NAVY }),
          run("\t\t"),
          ...pageField({ color: TEXT_NAVY }),
        ],
        style,
      );
      return [p];
    }
    case "ion": {
      const p =
        kind === "header"
          ? paragraph(
              [
                ...propertyField(doc, "TITLE", "[Document title]", { color: ACCENT_BLUE, bold: true }),
                run("\t\t"),
                ...propertyField(doc, "AUTHOR", "[Author name]", { color: ACCENT_BLUE }),
              ],
              style,
            )
          : paragraph([run("\t\t"), ...pageField({ color: ACCENT_BLUE, bold: true, fontSizeHalfPoints: 28 })], style);
      return [p];
    }
    case "retrospect": {
      const p =
        kind === "header"
          ? paragraph(
              [
                ...propertyField(doc, "TITLE", "[Document title]", { color: ACCENT_ORANGE }),
                run("\t\t"),
                ...field('DATE \\@ "MMMM d, yyyy"', "[Date]", { color: ACCENT_ORANGE }),
              ],
              style,
            )
          : paragraph(
              [
                ...propertyField(doc, "AUTHOR", "[Author name]", { color: ACCENT_ORANGE }),
                run("\t\t"),
                ...pageField({ color: ACCENT_ORANGE }),
              ],
              style,
            );
      setParagraphBorders(p, {
        style: "single",
        sizeEighthsOfPoint: 18,
        color: ACCENT_ORANGE,
        spacePt: 4,
        sides: [kind === "header" ? "bottom" : "top"],
      });
      return [p];
    }
    default: {
      const unhandled: never = preset;
      throw new Error(`Unknown preset ${String(unhandled)}`);
    }
  }
}

// --- cover pages ---------------------------------------------------------------------

export const COVER_PAGE_PRESETS = ["austin", "banded", "facet", "ion", "retrospect", "sideline"] as const;
export type CoverPagePresetId = (typeof COVER_PAGE_PRESETS)[number];

const TITLE_SIZE = 72;
const SUBTITLE_SIZE = 36;
const DETAIL_SIZE = 24;
// Vertical room above a cover title, in twips (roughly a third of a page).
const COVER_TOP_SPACE = 4320;

function coverParagraph(
  inlines: WmlInline[],
  options: {
    align?: "left" | "center" | "right";
    before?: number;
    after?: number;
    fill?: string;
    border?: { color: string; side: "top" | "bottom" | "left"; size: number };
  } = {},
): WmlParagraph {
  const p = paragraph(inlines);
  if (options.border) {
    setParagraphBorders(p, {
      style: "single",
      sizeEighthsOfPoint: options.border.size,
      color: options.border.color,
      spacePt: 8,
      sides: [options.border.side],
    });
  }
  if (options.fill) setParagraphShading(p, { fill: options.fill });
  setParagraphSpacing(p, { before: options.before ?? 0, after: options.after ?? 120 });
  if (options.align) setParagraphAlignment(p, options.align);
  return p;
}

/** The paragraphs of a cover page design (a page break follows them). */
export function coverPagePreset(doc: Docx, preset: CoverPagePresetId): WmlParagraph[] {
  const title = (format: RunFormatting): WmlRun[] =>
    propertyField(doc, "TITLE", "[Document title]", { fontSizeHalfPoints: TITLE_SIZE, ...format });
  const subtitle = (format: RunFormatting): WmlRun[] =>
    propertyField(doc, "SUBJECT", "[Document subtitle]", { fontSizeHalfPoints: SUBTITLE_SIZE, ...format });
  const author = (format: RunFormatting): WmlRun[] =>
    propertyField(doc, "AUTHOR", "[Author name]", { fontSizeHalfPoints: DETAIL_SIZE, ...format });
  const date = (format: RunFormatting): WmlRun[] =>
    field('DATE \\@ "MMMM d, yyyy"', "[Date]", { fontSizeHalfPoints: DETAIL_SIZE, ...format });
  switch (preset) {
    case "austin":
      return [
        coverParagraph(title({ color: ACCENT_BLUE }), {
          before: COVER_TOP_SPACE,
          border: { color: ACCENT_BLUE, side: "left", size: 48 },
        }),
        coverParagraph(subtitle({ color: TEXT_NAVY }), { border: { color: ACCENT_BLUE, side: "left", size: 48 } }),
        coverParagraph(author({ color: TEXT_NAVY }), { before: COVER_TOP_SPACE }),
      ];
    case "banded":
      return [
        coverParagraph(title({ color: WHITE, bold: true }), {
          before: COVER_TOP_SPACE,
          align: "center",
          fill: ACCENT_BLUE,
        }),
        coverParagraph(subtitle({ color: TEXT_NAVY }), { align: "center" }),
        coverParagraph([...author({ color: TEXT_NAVY }), run("  |  "), ...date({ color: TEXT_NAVY })], {
          before: COVER_TOP_SPACE,
          align: "center",
        }),
      ];
    case "facet":
      return [
        coverParagraph(title({ color: TEXT_NAVY, bold: true }), {
          before: COVER_TOP_SPACE,
          align: "right",
        }),
        coverParagraph(subtitle({ color: ACCENT_BLUE }), { align: "right" }),
        coverParagraph(author({ color: TEXT_NAVY }), {
          before: COVER_TOP_SPACE,
          align: "right",
          border: { color: ACCENT_BLUE, side: "top", size: 24 },
        }),
      ];
    case "ion":
      return [
        coverParagraph(date({ color: WHITE }), { align: "right", fill: ACCENT_BLUE }),
        coverParagraph(title({ color: ACCENT_BLUE }), { before: COVER_TOP_SPACE }),
        coverParagraph(subtitle({ color: TEXT_NAVY }), {}),
        coverParagraph(author({ color: TEXT_NAVY }), { before: COVER_TOP_SPACE, fill: LIGHT_FILL }),
      ];
    case "retrospect":
      return [
        coverParagraph(title({ color: ACCENT_ORANGE, bold: true }), {
          before: COVER_TOP_SPACE,
          align: "center",
          border: { color: ACCENT_ORANGE, side: "bottom", size: 24 },
        }),
        coverParagraph(subtitle({ color: TEXT_NAVY }), { align: "center" }),
        coverParagraph(date({ color: ACCENT_ORANGE }), { before: COVER_TOP_SPACE, align: "center" }),
        coverParagraph(author({ color: TEXT_NAVY }), { align: "center" }),
      ];
    case "sideline":
      return [
        coverParagraph(author({ color: TEXT_NAVY }), {
          border: { color: ACCENT_BLUE, side: "left", size: 24 },
        }),
        coverParagraph(title({ color: ACCENT_BLUE }), {
          before: COVER_TOP_SPACE,
          border: { color: ACCENT_BLUE, side: "left", size: 24 },
        }),
        coverParagraph(subtitle({ color: TEXT_NAVY }), {
          border: { color: ACCENT_BLUE, side: "left", size: 24 },
        }),
        coverParagraph(date({ color: TEXT_NAVY }), { before: COVER_TOP_SPACE }),
      ];
    default: {
      const unhandled: never = preset;
      throw new Error(`Unknown cover page ${String(unhandled)}`);
    }
  }
}
