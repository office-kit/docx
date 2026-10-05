import { describe, expect, it } from "vitest";
import {
  type LaidOutPage,
  type LayoutBlock,
  paginate,
  type PaginatorOptions,
  type PaginatorSection,
  sectionPageCounts,
} from "./paginate.js";

const OPTIONS: PaginatorOptions = { evenAndOddHeaders: false, noteSeparatorHeight: 10 };

function section(
  overrides: Partial<PaginatorSection> & { height?: number } = {},
): PaginatorSection {
  const height = overrides.height ?? 100;
  return {
    start: "nextPage",
    columns: 1,
    titlePage: false,
    vAlign: "top",
    bodyHeight: () => height,
    ...overrides,
  };
}

/** Block indices on each page, in order (a split block appears on both pages). */
function blocksPerPage(pages: readonly LaidOutPage[]): number[][] {
  return pages.map((p) =>
    p.regions.flatMap((r) => r.columns.flatMap((c) => c.fragments.map((f) => f.block))),
  );
}

const para = (height: number, extra: Partial<LayoutBlock> = {}): LayoutBlock => ({
  height,
  section: 0,
  ...extra,
});

/** Lines of equal height. */
const lines = (count: number, lineHeight: number) => () =>
  Array.from({ length: count }, (_, i) => (i + 1) * lineHeight);

describe("paginate", () => {
  it("fills pages in order and starts a new page when a block does not fit", () => {
    const pages = paginate([para(40), para(40), para(40)], [section()], OPTIONS);
    expect(blocksPerPage(pages)).toEqual([[0, 1], [2]]);
    expect(pages.map((p) => p.number)).toEqual([1, 2]);
  });

  it("gives an empty document one page", () => {
    expect(paginate([], [section()], OPTIONS)).toHaveLength(1);
  });

  it("honours pageBreakBefore except at the top of a page", () => {
    const pages = paginate(
      [para(10, { pageBreakBefore: true }), para(10), para(10, { pageBreakBefore: true })],
      [section()],
      OPTIONS,
    );
    expect(blocksPerPage(pages)).toEqual([[0, 1], [2]]);
  });

  it("splits a paragraph at a forced page break", () => {
    const pages = paginate(
      [para(30, { breaks: [{ offset: 10, kind: "page" }] })],
      [section()],
      OPTIONS,
    );
    expect(pages).toHaveLength(2);
    expect(pages[1]?.regions[0]?.columns[0]?.fragments[0]).toEqual({
      kind: "slice",
      block: 0,
      from: 10,
      to: 30,
    });
  });

  it("splits a long paragraph by its lines", () => {
    const pages = paginate([para(60), para(80, { lines: lines(8, 10) })], [section()], OPTIONS);
    expect(pages[0]?.regions[0]?.columns[0]?.fragments[1]).toEqual({
      kind: "slice",
      block: 1,
      from: 0,
      to: 40,
    });
    expect(pages[1]?.regions[0]?.columns[0]?.fragments[0]).toEqual({
      kind: "slice",
      block: 1,
      from: 40,
      to: 80,
    });
  });

  it("applies widow/orphan control", () => {
    // 4 lines fit; 5 would leave one line alone on the next page.
    const widow = paginate(
      [para(55), para(50, { lines: lines(5, 10), widowControl: true })],
      [section()],
      OPTIONS,
    );
    expect(widow[0]?.regions[0]?.columns[0]?.fragments[1]).toMatchObject({ to: 30 });
    // Only one line would fit: the whole paragraph moves.
    const orphan = paginate(
      [para(85), para(50, { lines: lines(5, 10), widowControl: true })],
      [section()],
      OPTIONS,
    );
    expect(blocksPerPage(orphan)).toEqual([[0], [1]]);
  });

  it("keeps lines together", () => {
    const pages = paginate(
      [para(60), para(50, { lines: lines(5, 10), keepLines: true })],
      [section()],
      OPTIONS,
    );
    expect(blocksPerPage(pages)).toEqual([[0], [1]]);
  });

  it("keeps a heading with the next paragraph", () => {
    const pages = paginate(
      [para(70), para(20, { keepNext: true }), para(30)],
      [section()],
      OPTIONS,
    );
    expect(blocksPerPage(pages)).toEqual([[0], [1, 2]]);
  });

  it("splits tables between rows and repeats header rows", () => {
    const rows = [
      { height: 10, header: true },
      ...Array.from({ length: 15 }, () => ({ height: 10 })),
    ];
    const pages = paginate([para(20), { height: 160, section: 0, rows }], [section()], OPTIONS);
    const second = pages[1]?.regions[0]?.columns[0]?.fragments[0];
    expect(pages[0]?.regions[0]?.columns[0]?.fragments[1]).toMatchObject({
      kind: "rows",
      from: 0,
      to: 8,
      headerRows: 0,
    });
    expect(second).toMatchObject({ kind: "rows", from: 8, to: 16, headerRows: 1, height: 90 });
  });

  it("flows into columns and balances them before a continuous break", () => {
    const pages = paginate(
      [
        { height: 30, section: 0 },
        { height: 30, section: 0 },
        { height: 30, section: 0 },
        { height: 30, section: 0 },
        { height: 10, section: 1 },
      ],
      [section({ columns: 2 }), section({ start: "continuous" })],
      OPTIONS,
    );
    expect(pages).toHaveLength(1);
    const [cols, after] = pages[0]?.regions ?? [];
    expect(cols?.columns.map((c) => c.fragments.length)).toEqual([2, 2]);
    expect(after?.top).toBe(60);
  });

  it("starts odd/even sections on the right side with a blank page", () => {
    const pages = paginate(
      [para(10), { height: 10, section: 1 }],
      [section(), section({ start: "oddPage" })],
      OPTIONS,
    );
    expect(pages.map((p) => [p.number, p.blank])).toEqual([
      [1, false],
      [2, true],
      [3, false],
    ]);
  });

  it("picks first / even page kinds and restarts numbering", () => {
    const pages = paginate(
      [para(100), para(100), para(100), { height: 100, section: 1 }],
      [section({ titlePage: true }), section({ pageNumberStart: 1 })],
      { ...OPTIONS, evenAndOddHeaders: true },
    );
    expect(pages.map((p) => [p.kind, p.number])).toEqual([
      ["first", 1],
      ["even", 2],
      ["default", 3],
      ["default", 1],
    ]);
    expect(sectionPageCounts(pages)).toEqual(
      new Map([
        [0, 3],
        [1, 1],
      ]),
    );
  });

  it("shrinks the body by footnotes on the page that references them", () => {
    const pages = paginate(
      [para(50, { notes: [{ key: "footnote:1", height: 30 }] }), para(20)],
      [section()],
      OPTIONS,
    );
    expect(pages[0]?.notes).toEqual(["footnote:1"]);
    expect(pages[0]?.noteHeight).toBe(40);
    expect(blocksPerPage(pages)).toEqual([[0], [1]]);
  });

  it("uses the body height of first pages (taller header)", () => {
    const pages = paginate(
      [para(60), para(60)],
      [section({ titlePage: true, bodyHeight: (kind) => (kind === "first" ? 50 : 200) })],
      OPTIONS,
    );
    expect(blocksPerPage(pages)).toEqual([[0], [1]]);
  });

  it("aligns content vertically", () => {
    const [centered] = paginate([para(40)], [section({ vAlign: "center" })], OPTIONS);
    expect(centered?.regions[0]?.offset).toBe(30);
    const [justified] = paginate(
      [para(20), para(20), para(20)],
      [section({ vAlign: "both" })],
      OPTIONS,
    );
    expect(justified?.regions[0]?.columns[0]?.gap).toBe(20);
  });

  it("paginates a 10 000-block document quickly (perf smoke)", () => {
    const blocks: LayoutBlock[] = Array.from({ length: 10_000 }, (_, i) =>
      i % 50 === 49
        ? { height: 300, section: 0, rows: Array.from({ length: 30 }, () => ({ height: 10 })) }
        : para(18 + (i % 5), { lines: lines(2, 9 + (i % 5) / 2), widowControl: true }),
    );
    const start = performance.now();
    const pages = paginate(blocks, [section({ height: 700 })], OPTIONS);
    const elapsed = performance.now() - start;
    expect(pages.length).toBeGreaterThan(200);
    expect(elapsed).toBeLessThan(250);
  });
});
