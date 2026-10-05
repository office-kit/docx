import type { IconSet } from "./types.js";

// Table Design / Table Layout and Insert ▸ Table. Names carry a `tbl` prefix
// so they cannot collide with other areas' icons.

const ACCENT = "#2B78E4";
const RED = "#D13438";
const GRID = "M3 3.5h14v13H3zM3 7.833h14M3 12.167h14M7.667 3.5v13M12.333 3.5v13";
const FAINT = "#B4B4B4";

// A cell box with text lines placed for one of the Alignment grid's nine positions.
function alignIcon(
  vertical: 0 | 1 | 2,
  horizontal: 0 | 1 | 2,
): readonly { d: string; stroke?: string }[] {
  const y = [6, 9, 12][vertical] ?? 6;
  const x = [5.5, 7.5, 9.5][horizontal] ?? 5.5;
  return [{ d: "M3 3.5h14v13H3z", stroke: FAINT }, { d: `M${x} ${y}h5M${x} ${y + 2.5}h5` }];
}

export default {
  tblShading: [
    { d: "M5 9.5 10 4.5l5 5-5 5zM10 4.5 8 2.5M15 9.5H5" },
    {
      d: "M16.5 12s-1.2 1.6-1.2 2.4a1.2 1.2 0 0 0 2.4 0c0-.8-1.2-2.4-1.2-2.4z",
      fill: "currentColor",
    },
  ],
  tblBorderStyles: [
    { d: "M3 3.5h9v9H3z", stroke: FAINT },
    { d: "M3 12.5h9", stroke: ACCENT },
    { d: "M11 16.5 17 10.5l-1.5-1.5L9.5 15z" },
  ],
  tblPen: "M4 16 5 12.5 13.5 4 16 6.5 7.5 15zM12 5.5 14.5 8",
  tblBorders: [{ d: GRID, stroke: FAINT }, { d: "M3 16.5h14" }],
  tblBorderBottom: [{ d: GRID, stroke: FAINT }, { d: "M3 16.5h14" }],
  tblBorderTop: [{ d: GRID, stroke: FAINT }, { d: "M3 3.5h14" }],
  tblBorderLeft: [{ d: GRID, stroke: FAINT }, { d: "M3 3.5v13" }],
  tblBorderRight: [{ d: GRID, stroke: FAINT }, { d: "M17 3.5v13" }],
  tblBorderNone: [{ d: GRID, stroke: FAINT }],
  tblBorderAll: GRID,
  tblBorderOutside: [{ d: GRID, stroke: FAINT }, { d: "M3 3.5h14v13H3z" }],
  tblBorderInside: [
    { d: "M3 3.5h14v13H3z", stroke: FAINT },
    { d: "M3 7.833h14M3 12.167h14M7.667 3.5v13M12.333 3.5v13" },
  ],
  tblBorderInsideH: [{ d: GRID, stroke: FAINT }, { d: "M3 7.833h14M3 12.167h14" }],
  tblBorderInsideV: [{ d: GRID, stroke: FAINT }, { d: "M7.667 3.5v13M12.333 3.5v13" }],
  tblBorderDiagDown: [{ d: "M3 3.5h14v13H3z", stroke: FAINT }, { d: "M3 3.5l14 13" }],
  tblBorderDiagUp: [{ d: "M3 3.5h14v13H3z", stroke: FAINT }, { d: "M3 16.5l14-13" }],
  tblBordersShading: [
    { d: "M3 3.5h14v13H3z" },
    { d: "M5.5 6h9v8h-9z", fill: "#DDE7F7", stroke: FAINT },
  ],
  tblBorderPainter: [
    { d: "M4 3.5h9v4H4zM13 5.5h2.5v4h-6v2" },
    { d: "M8.5 11.5h2v6h-2z", fill: "currentColor" },
  ],
  tblSelect: [
    { d: GRID, stroke: FAINT },
    { d: "M9 8l6 2.5-2.5 1 2.5 3-1.2 1-2.5-3-1.8 2z", fill: "#fff" },
  ],
  tblGridlines: [{ d: "M3 3.5h14v13H3zM3 8h14M3 12.5h14M8 3.5v13M12.5 3.5v13", stroke: ACCENT }],
  tblProperties: [{ d: GRID }, { d: "M3 3.5h14v4.333H3z", fill: "#DDE7F7" }],
  tblDraw: [
    { d: "M3 3.5h9M3 3.5v9M3 8h9M7.5 3.5v4.5", stroke: FAINT },
    { d: "M9 16.5 10 13l7-7 2.5 2.5-7 7z" },
  ],
  tblEraser: [
    { d: "M3 3.5h14v13H3z", stroke: FAINT },
    { d: "M6 14.5l6-6 3 3-4 4H8zM9 11.5l3 3", stroke: RED },
  ],
  tblDelete: [{ d: GRID }, { d: "M11.5 11.5l6 6M17.5 11.5l-6 6", stroke: RED }],
  tblInsertAbove: [
    { d: "M3 9.5h14v7H3zM3 13h14M8 9.5v7M12.5 9.5v7" },
    { d: "M3 3h14v4.5H3z", fill: "#DDE7F7", stroke: ACCENT },
  ],
  tblInsertBelow: [
    { d: "M3 3.5h14v7H3zM3 7h14M8 3.5v7M12.5 3.5v7" },
    { d: "M3 12.5h14V17H3z", fill: "#DDE7F7", stroke: ACCENT },
  ],
  tblInsertLeft: [
    { d: "M9.5 3h7.5v14H9.5zM9.5 8h7.5M9.5 12.5h7.5M13.25 3v14" },
    { d: "M3 3h4.5v14H3z", fill: "#DDE7F7", stroke: ACCENT },
  ],
  tblInsertRight: [
    { d: "M3 3h7.5v14H3zM3 8h7.5M3 12.5h7.5M6.75 3v14" },
    { d: "M12.5 3H17v14h-4.5z", fill: "#DDE7F7", stroke: ACCENT },
  ],
  tblMerge: [
    { d: "M3 3.5h14v13H3zM3 7.833h14M3 12.167h14", stroke: FAINT },
    { d: "M4.5 10h11M6.5 8 4.5 10l2 2M13.5 8l2 2-2 2", stroke: ACCENT },
  ],
  tblSplit: [
    { d: "M3 3.5h14v13H3z", stroke: FAINT },
    { d: "M10 3.5v13", stroke: ACCENT },
    { d: "M8 8l-2 2 2 2M12 8l2 2-2 2" },
  ],
  tblSplitTable: [
    { d: "M3 2.5h14v6H3zM3 11.5h14v6H3zM10 2.5v6M10 11.5v6" },
    { d: "M2 10h16", stroke: ACCENT },
  ],
  tblAutoFit: [{ d: GRID, stroke: FAINT }, { d: "M2 10h16M4.5 8 2 10l2.5 2M15.5 8l2.5 2-2.5 2" }],
  tblDistributeRows: [
    { d: "M5 3.5h12v13H5zM5 7.833h12M5 12.167h12" },
    { d: "M2.5 3.5v13M1.5 5l1-1.5 1 1.5M1.5 15l1 1.5 1-1.5", stroke: ACCENT },
  ],
  tblDistributeColumns: [
    { d: "M3 5.5h14v11H3zM7.667 5.5v11M12.333 5.5v11" },
    { d: "M3 2.5h14M4.5 1.5 3 2.5l1.5 1M15.5 1.5 17 2.5l-1.5 1", stroke: ACCENT },
  ],
  tblAlignTL: alignIcon(0, 0),
  tblAlignTC: alignIcon(0, 1),
  tblAlignTR: alignIcon(0, 2),
  tblAlignCL: alignIcon(1, 0),
  tblAlignCC: alignIcon(1, 1),
  tblAlignCR: alignIcon(1, 2),
  tblAlignBL: alignIcon(2, 0),
  tblAlignBC: alignIcon(2, 1),
  tblAlignBR: alignIcon(2, 2),
  tblTextDirection: [
    { d: "M3 3.5h14v13H3z", stroke: FAINT },
    { d: "M8 6v8M12 6v8M6.5 12.5 8 14l1.5-1.5M10.5 7.5 12 6l1.5 1.5" },
  ],
  tblCellMargins: [{ d: "M3 3.5h14v13H3z" }, { d: "M6 6.5h8v7H6z", stroke: ACCENT }],
  tblSort: [
    { d: "M5 3.5v13M3 14.5l2 2 2-2" },
    { d: "M10 3.5h3.5l-3.5 5h3.5M10 16.5l1.75-5 1.75 5M10.6 14.8h2.3", stroke: ACCENT },
  ],
  tblRepeatHeader: [
    { d: "M3 3.5h14v13H3zM3 12.167h14M7.667 7.833v8.667M12.333 7.833v8.667" },
    { d: "M3 3.5h14v4.333H3z", fill: "#DDE7F7", stroke: ACCENT },
  ],
  tblConvertToText: [
    { d: "M2.5 3.5h7v7h-7zM2.5 7h7M6 3.5v7", stroke: FAINT },
    { d: "M11 12h6.5M11 14.5h6.5M11 17h4.5M12 6h4.5M15 4.5l1.5 1.5L15 7.5" },
  ],
  tblConvertTextToTable: [
    { d: "M2.5 3.5h6M2.5 6h6M2.5 8.5h4" },
    { d: "M10.5 9.5h7v7h-7zM10.5 13h7M14 9.5v7", stroke: ACCENT },
  ],
  tblFormula: "M8.5 16.5c1.5 0 2-1 2.3-3l1.2-7c.3-2 .9-3 2.5-3M8.5 8.5h6M13 11.5l4 5M17 11.5l-4 5",
  tblQuickTables: [{ d: GRID }, { d: "M3 3.5h14v4.333H3z", fill: ACCENT, stroke: ACCENT }],
  tblNewStyle: [{ d: GRID }, { d: "M14 13v5M11.5 15.5h5", stroke: ACCENT }],
} as const satisfies IconSet;
