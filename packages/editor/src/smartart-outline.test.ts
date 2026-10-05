import { describe, expect, it } from "vitest";
import { applyOutlineOp, outlineToNodes, outlineRows } from "./smartart-outline.js";

const TREE = [{ text: "A", children: [{ text: "A1" }] }, { text: "B" }];

describe("SmartArt outline", () => {
  it("round-trips the node tree through rows", () => {
    expect(outlineRows(TREE)).toEqual([
      { level: 0, text: "A" },
      { level: 1, text: "A1" },
      { level: 0, text: "B" },
    ]);
    expect(outlineToNodes(outlineRows(TREE))).toEqual(TREE);
  });

  it("demotes, promotes and moves rows with their subtrees", () => {
    const rows = outlineRows(TREE);
    const demoted = applyOutlineOp(rows, 2, "demote", "");
    expect(outlineToNodes(demoted.rows)).toEqual([
      { text: "A", children: [{ text: "A1" }, { text: "B" }] },
    ]);
    expect(applyOutlineOp(demoted.rows, 2, "promote", "").rows).toEqual(rows);
    const moved = applyOutlineOp(rows, 2, "moveUp", "");
    expect(outlineToNodes(moved.rows)).toEqual([
      { text: "B" },
      { text: "A", children: [{ text: "A1" }] },
    ]);
    expect(applyOutlineOp(moved.rows, 0, "moveDown", "").rows).toEqual(rows);
  });

  it("adds shapes after, below and above", () => {
    const rows = outlineRows(TREE);
    expect(applyOutlineOp(rows, 0, "addAfter", "x")).toMatchObject({ index: 2 });
    expect(outlineToNodes(applyOutlineOp(rows, 2, "addBelow", "x").rows).at(-1)).toEqual({
      text: "B",
      children: [{ text: "x" }],
    });
    expect(outlineToNodes(applyOutlineOp(rows, 2, "addAbove", "x").rows).at(-1)).toEqual({
      text: "x",
      children: [{ text: "B" }],
    });
  });
});
