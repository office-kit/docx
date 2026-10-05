import type { IconSet } from "./types.js";

// Word's ribbon accent blue (the margin guides, column rules, break marks).
const ACCENT = "#2F6FBA";

export default {
  layTextDirection: [{ d: "M4 3h7M7.5 3v9" }, { d: "M14.5 4v12M12.5 14l2 2 2-2", stroke: ACCENT }],
  margins: [
    { d: "M4 2.5h12v15H4z" },
    { d: "M4 5.5h12M4 14.5h12M7 2.5v15M13 2.5v15", stroke: ACCENT },
  ],
  pageSize: [{ d: "M5 4.5h9v13H5z" }, { d: "M8 2.5h9v13", stroke: ACCENT }],
  layColumns: [
    {
      d: "M3 4h5.5M3 7h5.5M3 10h5.5M3 13h5.5M3 16h5.5M11.5 4H17M11.5 7H17M11.5 10H17M11.5 13H17M11.5 16H17",
    },
    { d: "M10 3v14", stroke: ACCENT },
  ],
  breaks: [
    { d: "M5 2.5v5h10v-5M5 17.5v-5h10v5" },
    { d: "M2.5 10h3M8.5 10h3M14.5 10h3", stroke: ACCENT },
  ],
  lineNumbers: [
    { d: "M8 4.5h9M8 10h9M8 15.5h9" },
    { d: "M3.5 3.5l1-1v4M3 8.5h2v1.5H3V11.5h2M3 14h2v3H3M3 15.5h2", stroke: ACCENT },
  ],
  hyphenation: [{ d: "M3 5h14M3 9h14M3 13h9M3 17h7" }, { d: "M13.5 13H17", stroke: ACCENT }],
  layIndentLeft: [{ d: "M8 5h9M8 9h9M8 13h9M8 17h9" }, { d: "M2.5 8l3 3-3 3", stroke: ACCENT }],
  layIndentRight: [{ d: "M3 5h9M3 9h9M3 13h9M3 17h9" }, { d: "M17.5 8l-3 3 3 3", stroke: ACCENT }],
  laySpacingBefore: [
    { d: "M8 12h9M8 16h9" },
    { d: "M4 3v8M2 5l2-2 2 2M2 9l2 2 2-2", stroke: ACCENT },
  ],
  laySpacingAfter: [
    { d: "M8 4h9M8 8h9" },
    { d: "M4 9v8M2 11l2-2 2 2M2 15l2 2 2-2", stroke: ACCENT },
  ],
} as const satisfies IconSet;
