/**
 * The blank document "New" opens, per UI language. Word's Normal template
 * differs by language; the Japanese one is set for Japanese text, and a
 * Japanese user expects it.
 */

import {
  createDocx,
  type Docx,
  PAGE_SIZE_A4,
  setDefaultParagraphSpacing,
  setPageMargins,
  setPageSize,
  setParagraphAlignment,
  setRunFormat,
  setSectionProperties,
  updateStyleFormatting,
} from "@office-kit/docx";
import type { LocaleId } from "./i18n/locales";

// Japanese Word for Mac 16's Normal.dotm: A4 with 35 / 30 / 30 / 30 mm
// margins, a line grid of 18 pt, and 游明朝 10.5 pt justified with no
// space between paragraphs.
const JA_MARGINS = {
  top: 1985,
  bottom: 1701,
  left: 1701,
  right: 1701,
  header: 851,
  footer: 992,
  gutter: 0,
};
const JA_LINE_PITCH = 360;
const JA_FONT = "游明朝";
const JA_SIZE_HALF_POINTS = 21;
const SINGLE_LINE = 240;

export function newDocument(locale: LocaleId): Docx {
  const doc = createDocx({ paragraphs: [""] });
  if (locale !== "ja") return doc;
  setPageSize(doc, PAGE_SIZE_A4);
  setPageMargins(doc, JA_MARGINS);
  setSectionProperties(doc, { documentGrid: { type: "lines", linePitch: JA_LINE_PITCH } });
  setDefaultParagraphSpacing(doc, { before: 0, after: 0, line: SINGLE_LINE });
  updateStyleFormatting(doc, "Normal", {
    run: (run) =>
      setRunFormat(run, {
        font: JA_FONT,
        fontEastAsia: JA_FONT,
        fontSizeHalfPoints: JA_SIZE_HALF_POINTS,
      }),
    paragraph: (p) => setParagraphAlignment(p, "both"),
  });
  return doc;
}
