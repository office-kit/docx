import type { IconSet } from "./types.js";

const BLUE = "#2b6cc4";
const GREEN = "#3a8a3a";
const RED = "#d13438";

// View tab, the Outlining tab and the status bar's view buttons.
export default {
  viewPrintLayout: "M5 2.5h10v15H5zM7.5 5.5h5M7.5 8h5M7.5 10.5h5M7.5 13h3",
  viewWebLayout: [
    { d: "M2.5 3.5h11v11h-11zM4.5 6.5h7M4.5 9h4" },
    {
      d: "M14 9a4 4 0 1 1 0 8 4 4 0 0 1 0-8zM10 13h8M14 9c-1.5 2.5-1.5 5.5 0 8M14 9c1.5 2.5 1.5 5.5 0 8",
      stroke: BLUE,
    },
  ],
  viewOutline: "M3 4.5h1M6 4.5h11M5 8.5h1M8 8.5h9M5 12.5h1M8 12.5h9M3 16.5h1M6 16.5h11",
  viewDraft: "M3 5h14M3 8.5h14M3 12h14M3 15.5h9",
  viewFocus: "M3 7.5V3h4.5M12.5 3H17v4.5M17 12.5V17h-4.5M7.5 17H3v-4.5",
  viewImmersiveReader: [
    { d: "M2.5 4.5c2.5-1 5-1 7.5.5v12c-2.5-1.5-5-1.5-7.5-.5zM17.5 4.5c-2.5-1-5-1-7.5.5" },
    { d: "M13 9.5a2 2 0 0 1 0 3M15 8a4.5 4.5 0 0 1 0 6", stroke: BLUE },
  ],
  viewZoom: "M8.5 3a5.5 5.5 0 1 1 0 11 5.5 5.5 0 0 1 0-11zM12.5 12.5 17 17",
  viewZoom100: [
    { d: "M4 2.5h7l3.5 3.5v11.5H4zM11 2.5V6h3.5" },
    { d: "M8.5 11h9v6h-9z", stroke: BLUE, fill: BLUE },
  ],
  viewOnePage: "M5.5 2.5h9v15h-9zM8 5.5h4M8 8h4M8 10.5h4",
  viewMultiplePages:
    "M2.5 4.5h6.5v11H2.5zM11 4.5h6.5v11H11zM4 7h3.5M4 9.5h3.5M12.5 7h3.5M12.5 9.5h3.5",
  viewPageWidth: [
    { d: "M5 2.5h10v15H5z" },
    { d: "M2 10h16M4 8l-2 2 2 2M16 8l2 2-2 2", stroke: BLUE },
  ],
  viewNewWindow: [{ d: "M3 6h14v11H3zM3 9h14" }, { d: "M6 1v6M3 4h6", stroke: GREEN }],
  viewArrangeAll: "M3 3h14v6.5H3zM3 10.5h14V17H3zM3 5h14M3 12.5h14",
  viewSplit: "M3 3h14v14H3zM3 5h14M3 10h14M3 12h14",
  viewSwitchWindows: [{ d: "M6 3h11v9H6zM6 5h11" }, { d: "M3 7v10h11", stroke: BLUE }],
  viewMacros: [
    { d: "M3 4h11v12H3zM3 6.5h11M5 9h7M5 11.5h7" },
    { d: "M13 11h5v6h-5zM14.5 13h2", stroke: RED },
  ],
  viewProperties:
    "M4.5 2.5h7L15 6v11.5H4.5zM11.5 2.5V6H15M7 9.5h1M9.5 9.5h3M7 12h1M9.5 12h3M7 14.5h1M9.5 14.5h3",
  outlinePromoteFull: "M15 5H5M8 2 5 5l3 3M3 2v6M5 11h12M5 14.5h12M5 18h12",
  outlinePromote: "M14 5H5M8 2 5 5l3 3M5 11h12M5 14.5h12M5 18h12",
  outlineDemote: "M5 5h9M11 2l3 3-3 3M5 11h12M5 14.5h12M5 18h12",
  outlineDemoteBody: "M5 5h9M11 2l3 3-3 3M16 2v6M5 11h12M5 14.5h12M5 18h12",
  outlineMoveUp: "M10 17V4M6 8l4-4 4 4",
  outlineMoveDown: "M10 3v13M6 12l4 4 4-4",
  outlineExpand: "M10 3v14M3 10h14",
  outlineCollapse: "M3 10h14",
  outlineClose: [{ d: "M3 3h14v14H3z" }, { d: "M7 7l6 6M13 7l-6 6", stroke: RED }],
} as const satisfies IconSet;
