import type { IconSet } from "./types.js";

// Word's accent colors on Insert-tab icons.
const BLUE = "#2b7cd3";
const RED = "#d13438";
const GREEN = "#2e9e44";
const ORANGE = "#e3913b";

export default {
  // A page with a title band and lines (Cover Page).
  coverPage: [
    { d: "M5 2.5h10v15H5z" },
    { d: "M7 5h6M7 7h4", stroke: BLUE },
    { d: "M7 11h6M7 13h6M7 15h4" },
  ],
  // An empty folded page (Blank Page).
  blankPage: [{ d: "M5 2.5h7l3 3v12H5zM12 2.5v3h3" }],
  // Two pages with an arrow between them (Cross-reference).
  crossReference: [
    { d: "M3 3.5h7v10H3zM10 6.5h7v10h-7" },
    { d: "M5.5 9.5h9M12.5 7.5l2 2-2 2", stroke: RED },
  ],
  // Speech bubble with a green plus (New Comment).
  newComment: [{ d: "M3 5h14v9.5H9l-4 3v-3H3z" }, { d: "M14.5 1v6M11.5 4h6", stroke: GREEN }],
  // A page with an accent band at the top (Header) / bottom (Footer).
  header: [
    { d: "M5 2.5h10v15H5z" },
    { d: "M5 3h10v3.5H5z", fill: ORANGE, stroke: ORANGE },
    { d: "M7 9h6M7 11h6M7 13h4" },
  ],
  footer: [
    { d: "M5 2.5h10v15H5z" },
    { d: "M5 13.5h10V17H5z", fill: ORANGE, stroke: ORANGE },
    { d: "M7 5h6M7 7h6M7 9h4" },
  ],
  // A large initial "A" beside text lines (Drop Cap).
  dropCap: [
    { d: "M3 11 5.5 3.5 8 11M4 8.5h3", stroke: BLUE },
    { d: "M10 4h7M10 7h7M10 10h7M3 13.5h14M3 16.5h11" },
  ],
  // A grey field box with a code mark (Field).
  field: [
    { d: "M2.5 5.5h15v9h-15z" },
    { d: "M6 8.5 5 10l1 1.5M14 8.5l1 1.5-1 1.5M8.5 12l3-4", stroke: BLUE },
  ],
  // Calendar with a clock (Date & Time).
  dateTime: [
    { d: "M3 4.5h11v11H3zM3 7.5h11M6 3v3M11 3v3" },
    { d: "M14.5 11a3.5 3.5 0 1 0 0.01 0zM14.5 12.5v2l1.2 0.8", stroke: BLUE },
  ],
  // A framed object (Object).
  object: [{ d: "M3 3.5h14v13H3z" }, { d: "M6 13l3-4 2 2.5 1.5-1.5L15 13z", stroke: BLUE }],
  // Pi (Equation).
  equation: [{ d: "M4 5.5h12M7.5 5.5v10M12.5 5.5v8.5c0 1 0.6 1.5 1.5 1.5" }],
  // Omega (Advanced Symbol).
  symbol: [{ d: "M4 16.5h3.5v-1.6A6 6 0 1 1 12.5 14.9v1.6H16" }],
  // An X over a signature rule (Signature Line).
  signatureLine: [
    { d: "M3 14.5h14" },
    { d: "M4 9.5l3 3M7 9.5l-3 3", stroke: BLUE },
    { d: "M9 11c1.5-3 2.5-3 3 0s1.5 1.5 3-1" },
  ],
} as const satisfies IconSet;
