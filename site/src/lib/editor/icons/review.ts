import type { IconSet } from "./types.js";

// Word's accents on the Review tab.
const BLUE = "#2b6cc4";
const GREEN = "#3a8a3a";
const RED = "#d13438";
const ORANGE = "#e3893b";

// Review tab and the status bar's proofing / accessibility indicators.
export default {
  reviewEditor: [
    { d: "M4 16.5 13.5 7l2.5 2.5L6.5 19H4z", stroke: BLUE },
    { d: "M12.5 8 15 10.5M14.5 6l1-1a1.5 1.5 0 0 1 2.5 2.5l-1 1" },
    { d: "M2.5 13.5h6M2.5 10.5h8" },
  ],
  reviewSpelling: [
    {
      d: "M2.5 9 4.5 3l2 6M3.2 7h2.6M8.5 3v6h1.8a1.5 1.5 0 0 0 0-3H8.5h1.5a1.5 1.5 0 0 0 0-3H8.5M17 4a2.5 2.5 0 1 0 0 4.5",
    },
    { d: "M5 13.5 8 16.5 15 10", stroke: GREEN },
  ],
  reviewThesaurus:
    "M3 4.5h6a1 1 0 0 1 1 1v11a1 1 0 0 0-1-1H3zM17 4.5h-6a1 1 0 0 0-1 1v11a1 1 0 0 1 1-1h6zM4.5 7.5h3M4.5 10h3M12.5 7.5h3M12.5 10h3",
  reviewWordCount: [
    { d: "M3 5h14M3 8.5h14M3 12h6" },
    { d: "M11 13.5l1-.5v4M13.5 13.5h2l-2 3h2M17 13.5h1.5l-1 1.2 1 1.3H17", stroke: BLUE },
  ],
  reviewReadAloud: [
    { d: "M3 17 8 3h1l5 14M5 12h7" },
    { d: "M15 6a3 3 0 0 1 0 4M16.5 4a6 6 0 0 1 0 8", stroke: BLUE },
  ],
  reviewAccessibility: [
    { d: "M4 2.5h7l3.5 3.5v5M4 2.5v15h6M11 2.5V6h3.5" },
    { d: "M14 11a1 1 0 1 0 0 .01M11 13h6M14 13v2.5l-2 2.5M14 15.5l2 2.5", stroke: BLUE },
  ],
  reviewTranslate: [
    { d: "M2.5 4.5h7M6 3v1.5M4 4.5a6 6 0 0 0 4.5 5M8 4.5A6 6 0 0 1 3 10", stroke: BLUE },
    { d: "M10 17l3-7.5 3 7.5M11 14.5h4", stroke: GREEN },
  ],
  reviewLanguage: [
    { d: "M3 17 7.5 4h1L13 17M5 12.5h6" },
    { d: "M14 4v6M12.5 5.5h3M17 3.5v13", stroke: BLUE },
  ],
  reviewNewComment: [{ d: "M3 5h14v9.5H9l-4 3v-3H3z" }, { d: "M14.5 1v6M11.5 4h6", stroke: GREEN }],
  reviewDeleteComment: [
    { d: "M3 5h11v8H8l-3 2.5V13H3z" },
    { d: "M13 11l5 5M18 11l-5 5", stroke: RED },
  ],
  reviewResolve: [{ d: "M3 5h11v8H8l-3 2.5V13H3z" }, { d: "M12 14l2 2 4-4.5", stroke: GREEN }],
  reviewPrevComment: [
    { d: "M5 5h12v8h-6l-3 2.5V13H5z" },
    { d: "M5.5 9H1.5M3.5 7 1.5 9l2 2", stroke: BLUE },
  ],
  reviewNextComment: [
    { d: "M3 5h12v8H9l-3 2.5V13H3z" },
    { d: "M14.5 9h4M16.5 7l2 2-2 2", stroke: BLUE },
  ],
  reviewShowComments: "M3 4h14v9.5H9l-4 3v-3H3zM6 7.5h8M6 10h5",
  reviewTrackChanges: [
    { d: "M4 2.5h7l3.5 3.5v3M4 2.5v15h5M11 2.5V6h3.5M6.5 8h5M6.5 11h3" },
    { d: "M10 18l.5-2.5 6-6 2 2-6 6z", stroke: BLUE },
  ],
  reviewDisplay: [
    { d: "M3 2.5h6l2.5 2.5v8.5H3zM5 6h4M5 8.5h4.5M5 11h3" },
    { d: "M9.5 7h6l2.5 2.5v8H9.5", stroke: RED },
  ],
  reviewMarkupOptions: "M4.5 2.5h7L15 6v11.5H4.5zM11.5 2.5V6H15M7 9h5M7 11.5h5M7 14h3",
  reviewReviewingPane: [
    { d: "M2.5 4h15v12h-15z" },
    { d: "M5 10h10M12 7l3 3-3 3M5 6.5v7", stroke: BLUE },
  ],
  reviewAccept: [
    { d: "M4 2.5h7l3.5 3.5v6M4 2.5v15h5M11 2.5V6h3.5" },
    { d: "M9.5 15.5 12 18l6-6.5", stroke: GREEN },
  ],
  reviewReject: [
    { d: "M4 2.5h7l3.5 3.5v6M4 2.5v15h5M11 2.5V6h3.5" },
    { d: "M11 12l6 6M17 12l-6 6", stroke: RED },
  ],
  reviewPrevChange: [
    { d: "M5 2.5h7l3.5 3.5v11.5H5zM12 2.5V6h3.5" },
    { d: "M13 11H7.5M9.5 9 7.5 11l2 2", stroke: BLUE },
  ],
  reviewNextChange: [
    { d: "M5 2.5h7l3.5 3.5v11.5H5zM12 2.5V6h3.5" },
    { d: "M7.5 11H13M11 9l2 2-2 2", stroke: BLUE },
  ],
  reviewCompare: [
    { d: "M2.5 3.5h7v13h-7zM4.5 6.5h3M4.5 9h3M4.5 11.5h3", stroke: BLUE },
    { d: "M11.5 3.5h6v13h-6zM13.5 6.5h2M13.5 9h2M13.5 11.5h2" },
  ],
  reviewBlockAuthors: [
    { d: "M8 3a3 3 0 1 1 0 6 3 3 0 0 1 0-6zM2.5 17a5.5 5.5 0 0 1 9-4" },
    { d: "M15 11a3 3 0 1 1 0 6 3 3 0 0 1 0-6zM13 15l4-2", stroke: RED },
  ],
  reviewProtect: [
    { d: "M3.5 2.5h7l3.5 3.5v3M3.5 2.5v15h5M10.5 2.5V6H14" },
    { d: "M11 12.5h7v5.5h-7zM12.5 12.5v-1.5a2 2 0 0 1 4 0v1.5", stroke: ORANGE },
  ],
  reviewReadOnly: [
    { d: "M4 16 13 7l2.5 2.5-9 9H4z" },
    { d: "M15 13.5a3 3 0 1 1 0 .01M13 15.5l4-4", stroke: RED },
  ],
  reviewHideInk: [
    { d: "M3 3.5h14v13H3z" },
    { d: "M5 13c2-4 3-4 4-1s2 3 3 0 2-3 3-1", stroke: RED },
  ],
  statusProofing: [
    { d: "M3.5 3.5h6a1 1 0 0 1 1 1v12a1 1 0 0 0-1-1h-6zM16.5 3.5h-4.5" },
    { d: "M11.5 13l2 2 4-4.5", stroke: GREEN },
  ],
  statusTracking: [
    { d: "M7 4a2.5 2.5 0 1 1 0 5 2.5 2.5 0 0 1 0-5zM2.5 16a4.5 4.5 0 0 1 9 0" },
    { d: "M13.5 15.5l.5-2 4-4 1.5 1.5-4 4z", stroke: BLUE },
  ],
} as const satisfies IconSet;
