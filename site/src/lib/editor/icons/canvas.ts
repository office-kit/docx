import type { IconSet } from "./types.js";

// Word's header/footer glyphs: a page outline with the header or footer band
// in Word's blue.
const PAGE = "M5 2.5h10v15H5z";
const BLUE = "#2b67c6";

/** Icons of the Header & Footer contextual tab. */
export default {
  header: [
    { d: PAGE },
    { d: "M6 4h8v2.5H6z", fill: BLUE, stroke: BLUE },
    { d: "M7 10h6M7 12.5h6M7 15h4" },
  ],
  footer: [
    { d: PAGE },
    { d: "M6 13.5h8V16H6z", fill: BLUE, stroke: BLUE },
    { d: "M7 5h6M7 7.5h6M7 10h4" },
  ],
  goToHeader: [
    { d: PAGE },
    { d: "M6 4h8v2.5H6z", fill: BLUE, stroke: BLUE },
    { d: "M10 15.5v-6M7.5 12l2.5-2.5 2.5 2.5" },
  ],
  goToFooter: [
    { d: PAGE },
    { d: "M6 13.5h8V16H6z", fill: BLUE, stroke: BLUE },
    { d: "M10 4.5v6M7.5 8l2.5 2.5L12.5 8" },
  ],
  previousSection: [{ d: PAGE }, { d: "M10 14.5v-9M6.5 9l3.5-3.5L13.5 9", stroke: BLUE }],
  nextSection: [{ d: PAGE }, { d: "M10 5.5v9M6.5 11l3.5 3.5 3.5-3.5", stroke: BLUE }],
  linkToPrevious: [
    { d: "M8.5 11.5l3-3" },
    { d: "M9.5 6.5l1.5-1.5a2.5 2.5 0 013.5 3.5L13 10", stroke: BLUE },
    { d: "M10.5 13.5L9 15a2.5 2.5 0 01-3.5-3.5L7 10", stroke: BLUE },
  ],
  closeHeaderFooter: [
    { d: "M3.5 3.5h13v13h-13z", fill: "#d13438", stroke: "#d13438" },
    { d: "M7 7l6 6M13 7l-6 6", stroke: "#fff" },
  ],
} as const satisfies IconSet;
