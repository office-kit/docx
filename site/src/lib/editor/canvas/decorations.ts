/**
 * What Word draws on a page besides the text: crop marks at the corners of
 * the text area, page borders (`w:pgBorders`, §17.6.10), the page colour,
 * gridlines (View ▸ Gridlines, spaced by `w:docGrid`), line numbers
 * (`w:lnNumType`, §17.6.8), text watermarks, and the per-page field results
 * (PAGE, NUMPAGES, SECTIONPAGES, SECTION) and note numbers.
 */

import {
  fieldFormatSwitch,
  formatNumber,
  type LaidOutPage,
  PAGE_FIELDS,
  type BorderLine,
  type SectionModel,
  STORY_ATTR,
} from "@office-kit/docx-editor";
import type { PageGeometry } from "./page-layout";
import { eighthPtToPx, ptToPx, twipsToPx } from "./units";
import { textRects, zoomOf } from "./measure";

const CROP_MARK_PX = ptToPx(13.5);
// Word's gridline spacing when the document grid sets none: one 12 pt line.
const DEFAULT_GRID_PITCH_TWIPS = 240;
// `w:lnNumType` without `w:distance`: Word's automatic 0.25".
const DEFAULT_LINE_NUMBER_DISTANCE_TWIPS = 360;
const LINE_NUMBER_WIDTH_PX = 40;
// Vertical tolerance when grouping client rects into lines.
const LINE_SLACK_PX = 1;

const LINE_BORDER_CSS: Readonly<Record<string, string>> = {
  single: "solid",
  thick: "solid",
  double: "double",
  triple: "double",
  dotted: "dotted",
  dashed: "dashed",
  dashSmallGap: "dashed",
  dotDash: "dashed",
  dotDotDash: "dotted",
  thinThickSmallGap: "double",
  thickThinSmallGap: "double",
  thinThickThinSmallGap: "double",
  thinThickMediumGap: "double",
  thickThinMediumGap: "double",
  thinThickThinMediumGap: "double",
  thinThickLargeGap: "double",
  thickThinLargeGap: "double",
  thinThickThinLargeGap: "double",
  wave: "solid",
  doubleWave: "double",
  dashDotStroked: "dashed",
  threeDEmboss: "ridge",
  threeDEngrave: "groove",
  outset: "outset",
  inset: "inset",
};
const HEX = /^[0-9A-Fa-f]{6}$/;

function borderCss(line: BorderLine | undefined): string {
  if (!line) return "none";
  const color = HEX.test(line.color) ? `#${line.color}` : "#000";
  const css = LINE_BORDER_CSS[line.style];
  // Art borders (apples, stars …, §17.18.2) are pictures Word repeats along
  // the edge; their w:sz is in points. Approximated by a ridge line.
  if (!css) return `${Math.max(1, ptToPx(line.size))}px ridge ${color}`;
  const width = Math.max(css === "double" ? 3 : 1, eighthPtToPx(Math.max(line.size, 2)));
  return `${width}px ${css} ${color}`;
}

function div(cls: string): HTMLElement {
  const e = document.createElement("div");
  e.className = cls;
  e.setAttribute("contenteditable", "false");
  return e;
}

function cropMarks(box: HTMLElement, g: PageGeometry): void {
  const corners = [
    { x: g.left - CROP_MARK_PX, y: g.bodyTop - CROP_MARK_PX, sides: ["right", "bottom"] },
    { x: g.width - g.right, y: g.bodyTop - CROP_MARK_PX, sides: ["left", "bottom"] },
    { x: g.left - CROP_MARK_PX, y: g.height - g.bodyBottom, sides: ["right", "top"] },
    { x: g.width - g.right, y: g.height - g.bodyBottom, sides: ["left", "top"] },
  ];
  for (const c of corners) {
    const mark = div("wk-crop");
    mark.style.left = `${c.x}px`;
    mark.style.top = `${c.y}px`;
    for (const side of c.sides) mark.style.setProperty(`border-${side}-style`, "solid");
    box.appendChild(mark);
  }
}

function pageBorders(
  box: HTMLElement,
  page: LaidOutPage,
  section: SectionModel,
  g: PageGeometry,
): void {
  const b = section.borders;
  if (!b) return;
  if (b.display === "firstPage" && page.numberInSection !== 1) return;
  if (b.display === "notFirstPage" && page.numberInSection === 1) return;
  const frame = div(`wk-pgborder ${b.zOrder === "back" ? "back" : "front"}`);
  const space = (line: BorderLine | undefined): number => ptToPx(line?.space ?? 0);
  if (b.offsetFrom === "page") {
    frame.style.top = `${space(b.top)}px`;
    frame.style.left = `${space(b.left)}px`;
    frame.style.right = `${space(b.right)}px`;
    frame.style.bottom = `${space(b.bottom)}px`;
  } else {
    frame.style.top = `${g.bodyTop - space(b.top)}px`;
    frame.style.left = `${g.left - space(b.left)}px`;
    frame.style.right = `${g.right - space(b.right)}px`;
    frame.style.bottom = `${g.bodyBottom - space(b.bottom)}px`;
  }
  frame.style.borderTop = borderCss(b.top);
  frame.style.borderLeft = borderCss(b.left);
  frame.style.borderRight = borderCss(b.right);
  frame.style.borderBottom = borderCss(b.bottom);
  box.appendChild(frame);
}

function gridlines(box: HTMLElement, section: SectionModel, g: PageGeometry): void {
  const grid = div("wk-grid");
  grid.style.top = `${g.bodyTop}px`;
  grid.style.left = `${g.left}px`;
  grid.style.width = `${g.textWidth}px`;
  grid.style.height = `${g.height - g.bodyTop - g.bodyBottom}px`;
  grid.style.setProperty(
    "--pitch",
    `${twipsToPx(section.grid?.linePitch ?? DEFAULT_GRID_PITCH_TWIPS)}px`,
  );
  box.appendChild(grid);
}

export interface DecorationContext {
  readonly page: LaidOutPage;
  readonly section: SectionModel;
  readonly geometry: PageGeometry;
  readonly showGridlines: boolean;
}

export function decoratePage(box: HTMLElement, ctx: DecorationContext): void {
  cropMarks(box, ctx.geometry);
  pageBorders(box, ctx.page, ctx.section, ctx.geometry);
  if (ctx.showGridlines) gridlines(box, ctx.section, ctx.geometry);
}

export interface FieldContext {
  readonly page: LaidOutPage;
  readonly section: SectionModel;
  readonly totalPages: number;
  readonly sectionPages: number;
  readonly noteMarks: ReadonlyMap<string, string>;
}

/** Write the page-dependent field results and note numbers into a page box. */
export function fillPageFields(box: HTMLElement, ctx: FieldContext): void {
  let previous: { el: Element; instr: string } | undefined;
  for (const field of box.querySelectorAll<HTMLElement>("[data-wk-field]")) {
    const type = field.dataset.wkField ?? "";
    if (!PAGE_FIELDS.has(type)) continue;
    const instr = field.dataset.wkInstr ?? "";
    // A result spread over several runs shows the number once, in the first.
    const continuation =
      previous !== undefined &&
      previous.instr === instr &&
      previous.el.parentElement === field.parentElement &&
      previous.el.nextElementSibling === field;
    previous = { el: field, instr };
    const format = fieldFormatSwitch(instr);
    const value =
      type === "PAGE"
        ? formatNumber(ctx.page.number, format ?? ctx.section.pageNumbers.format)
        : type === "NUMPAGES"
          ? formatNumber(ctx.totalPages, format)
          : type === "SECTIONPAGES"
            ? formatNumber(ctx.sectionPages, format)
            : formatNumber(ctx.section.index + 1, format);
    const text = continuation ? "" : value;
    if (field.textContent !== text) field.textContent = text;
  }
  for (const mark of box.querySelectorAll<HTMLElement>(".wk-noteref")) {
    const key =
      mark.dataset.wkNote ?? mark.closest(`[${STORY_ATTR}]`)?.getAttribute(STORY_ATTR) ?? "";
    const text = ctx.noteMarks.get(key) ?? "";
    if (mark.textContent !== text) mark.textContent = text;
  }
}

/**
 * Line numbers in the left margin of every section that turns them on.
 * Lines are read from the rendered page (only visible lines count), so this
 * runs after the page boxes are in the document.
 */
export function numberLines(
  root: HTMLElement,
  pages: readonly LaidOutPage[],
  sections: readonly SectionModel[],
  suppressed: (block: number) => boolean,
): void {
  if (!sections.some((s) => s.lineNumbers)) return;
  let count = 0;
  let lastSection = -1;
  const boxes = root.querySelectorAll<HTMLElement>(".wk-pagebox");
  pages.forEach((page, p) => {
    const box = boxes[p];
    if (!box) return;
    const zoom = zoomOf(box);
    const boxRect = box.getBoundingClientRect();
    const regions = box.querySelectorAll<HTMLElement>(".wk-regions > .wk-region");
    page.regions.forEach((region, r) => {
      const numbering = sections[region.section]?.lineNumbers;
      const regionEl = regions[r];
      if (!numbering || !regionEl) return;
      if (numbering.restart === "newPage" && r === 0) count = 0;
      if (numbering.restart === "newSection" && region.section !== lastSection) count = 0;
      lastSection = region.section;
      const distance = twipsToPx(numbering.distance ?? DEFAULT_LINE_NUMBER_DISTANCE_TWIPS);
      for (const col of regionEl.querySelectorAll<HTMLElement>(":scope > .wk-col")) {
        const colRect = col.getBoundingClientRect();
        for (const para of col.querySelectorAll<HTMLElement>(".wk-p")) {
          if (para.closest("table")) continue;
          const block = Number.parseInt(para.dataset.wkBlock ?? "", 10);
          if (suppressed(block)) continue;
          const clip = para.closest(".wk-slice")?.getBoundingClientRect() ?? colRect;
          const tops: number[] = [];
          let bottom = Number.NEGATIVE_INFINITY;
          for (const rect of textRects(para).toSorted((a, b) => a.top - b.top)) {
            if (rect.height === 0 || rect.top < bottom - LINE_SLACK_PX) continue;
            bottom = rect.bottom;
            const mid = (rect.top + rect.bottom) / 2;
            if (mid >= clip.top && mid <= clip.bottom) tops.push(rect.top);
          }
          for (const top of tops) {
            const n = numbering.start + count;
            count++;
            if (n % numbering.countBy !== 0) continue;
            const label = div("wk-linenum");
            label.textContent = String(n);
            label.style.top = `${(top - boxRect.top) / zoom}px`;
            label.style.left = `${(colRect.left - boxRect.left) / zoom - distance - LINE_NUMBER_WIDTH_PX}px`;
            label.style.width = `${LINE_NUMBER_WIDTH_PX}px`;
            box.appendChild(label);
          }
        }
      }
    });
  });
}
