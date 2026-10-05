import type { IconSet } from "./types.js";

// Word's References accent: the blue of its link and table-of-contents glyphs.
const BLUE = "#2b6cd4";

export default {
  addText: [{ d: "M3 5h9M3 9h9M3 13h6" }, { d: "M15 11v6M12 14h6", stroke: BLUE }],
  updateTable: [
    { d: "M3 4h8M3 8h6M3 12h5" },
    { d: "M17 11a4 4 0 1 1-1.2-2.8M17 6.5v2.2h-2.2", stroke: BLUE },
  ],
  nextFootnote: [{ d: "M3 4h14M3 8h14M3 12h8" }, { d: "M12 15h5M15 13l2 2-2 2", stroke: BLUE }],
  showNotes: [
    { d: "M4 2.5h12v15H4z" },
    { d: "M6.5 12.5h7M6.5 15h5", stroke: BLUE },
    { d: "M6.5 5.5h7M6.5 8h7" },
  ],
  noteOptions: [
    { d: "M3 4h14M3 8h14M3 12h6" },
    { d: "M14 12.5a2.5 2.5 0 1 1 0 5 2.5 2.5 0 0 1 0-5zM14 11v1.5M14 17.5V19", stroke: BLUE },
  ],
  citation: [
    { d: "M4 3h12v14H4z" },
    {
      d: "M7 7.5c-1 .5-1.2 1.5-1 2.5h1.5v1.8H6M11.5 7.5c-1 .5-1.2 1.5-1 2.5H12v1.8h-1.5",
      stroke: BLUE,
    },
    { d: "M7 14.5h6" },
  ],
  citations: [
    { d: "M3.5 3h13v14h-13zM12.5 3v14" },
    { d: "M5.5 6.5h5M5.5 9h5M5.5 11.5h3", stroke: BLUE },
  ],
  bibliography: [
    { d: "M4 3h8l4 4v10H4zM12 3v4h4" },
    { d: "M6.5 10h7M6.5 12.5h7M6.5 15h4", stroke: BLUE },
  ],
  manageSources: [
    { d: "M3 4.5h9v12H3zM6 2.5h9v12" },
    { d: "M5 8h5M5 10.5h5M5 13h3", stroke: BLUE },
  ],
  caption: [
    { d: "M3 3h14v9H3z" },
    { d: "M5.5 10 8 7l2 2 1.5-1.5L15 10", stroke: BLUE },
    { d: "M3 15h14M3 17.5h9" },
  ],
  tableOfFigures: [
    { d: "M2.5 3h7v6h-7z" },
    { d: "M4 7.5l1.5-2 1.5 1.5 1-1 1 1.5", stroke: BLUE },
    { d: "M11.5 4.5h6M11.5 7.5h6M2.5 12h15M2.5 15h15" },
  ],
  crossReference: [
    { d: "M3 3h8v6H3zM9 11h8v6H9z" },
    { d: "M7 9v4.5h2M13 11V6.5h-2", stroke: BLUE },
  ],
  markEntry: [
    { d: "M3 5h14M3 9h7M3 13h14M3 17h9" },
    { d: "M11.5 7.5h5v3h-5z", stroke: BLUE, fill: "#dce8fb" },
  ],
  insertIndex: [
    { d: "M3 3h6v14H3zM11 3h6v14h-6z" },
    { d: "M4.5 5.5h3M12.5 5.5h3", stroke: BLUE },
    { d: "M4.5 8h3M4.5 10.5h3M12.5 8h3M12.5 10.5h3" },
  ],
  markCitation: [
    { d: "M3 5h14M3 9h14M3 13h7" },
    { d: "M12 11.5h5v6h-5zM13.5 13.5h2M13.5 15.5h2", stroke: BLUE },
  ],
  toa: [{ d: "M10 3v13M5 16.5h10M4 6h12" }, { d: "M4 6l-2 5h4zM16 6l-2 5h4z", stroke: BLUE }],
} as const satisfies IconSet;
