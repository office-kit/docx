import type { IconSet } from "./types.js";

// Word's Mailings accents: the blue merge chevrons and the yellow envelope flap.
const BLUE = "#2b6cd4";
const YELLOW = "#e8a33c";

export default {
  envelope: [{ d: "M2.5 5h15v10h-15z" }, { d: "M2.5 5 10 11l7.5-6", stroke: YELLOW }],
  labels: [{ d: "M3 2.5h14v15H3z" }, { d: "M5 5h10v3H5zM5 10h10v3H5z", stroke: BLUE }],
  startMerge: [
    { d: "M4 2.5h8l3.5 3.5v11.5H4zM12 2.5V6h3.5" },
    { d: "M6.5 10h3l1.5 2-1.5 2h-3M11.5 12h2", stroke: BLUE },
  ],
  recipients: [
    { d: "M7 4.5a2.5 2.5 0 1 1 0 5 2.5 2.5 0 0 1 0-5zM2.5 16c0-2.8 2-4.5 4.5-4.5s4.5 1.7 4.5 4.5" },
    { d: "M13 5a2 2 0 1 1 0 4M14 11.5c2 .3 3.5 1.8 3.5 4.5", stroke: BLUE },
  ],
  editRecipients: [
    { d: "M3 3.5h11v13H3zM3 7h11M3 10.5h11M7 3.5v13" },
    { d: "M17.5 9 12 14.5l-.5 2 2-.5 5.5-5.5z", stroke: BLUE },
  ],
  filterRecipients: [
    { d: "M3 4.5h11v12H3zM3 8h11M3 11.5h11" },
    { d: "M12 11h6l-2.3 3v3l-1.4-.8V14z", stroke: BLUE },
  ],
  mergeField: [
    { d: "M3 4h14M3 8h4M13 8h4M3 12h14M3 16h9" },
    { d: "M9.5 6.5 8 8l1.5 1.5M10.5 6.5 12 8l-1.5 1.5", stroke: BLUE },
  ],
  rules: [
    { d: "M3 4h8M3 8h6M3 12h8M3 16h5" },
    { d: "M15 3.5v3l-2 2 2 2v3M13 15.5h4", stroke: BLUE },
  ],
  matchFields: [
    { d: "M2.5 4h6v3h-6zM2.5 13h6v3h-6zM11.5 4h6v3h-6zM11.5 13h6v3h-6z" },
    { d: "M8.5 5.5h3M8.5 14.5h3", stroke: BLUE },
  ],
  updateLabels: [
    { d: "M3 3h9v6H3zM3 11h9v6H3z" },
    { d: "M17 11a3 3 0 1 1-.9-2.1M17 7.5v2h-2", stroke: BLUE },
  ],
  highlightFields: [
    { d: "M3 4h14M3 16h14" },
    { d: "M3 8h14v4H3z", stroke: BLUE, fill: "#d9d9d9" },
    { d: "M6 10h8" },
  ],
  previewResults: [
    { d: "M3.5 3h13v14h-13z" },
    {
      d: "M10 7.5c2.5 0 4 2.5 4 2.5s-1.5 2.5-4 2.5S6 10 6 10s1.5-2.5 4-2.5zM10 9.2v1.6",
      stroke: BLUE,
    },
  ],
  firstRecord: "M6 5v10M15 5l-6 5 6 5",
  previousRecord: "M13 5l-6 5 6 5",
  nextRecord: "M7 5l6 5-6 5",
  lastRecord: "M14 5v10M5 5l6 5-6 5",
  findRecipient: [
    { d: "M6 4a2.3 2.3 0 1 1 0 4.6A2.3 2.3 0 0 1 6 4zM2 14.5c0-2.4 1.8-4 4-4" },
    { d: "M12.5 9a3 3 0 1 1 0 6 3 3 0 0 1 0-6zM14.7 14.7 17.5 17.5", stroke: BLUE },
  ],
  checkErrors: [
    { d: "M4 2.5h8l3.5 3.5v11.5H4zM12 2.5V6h3.5" },
    { d: "M6.5 11.5l2 2 4-4.5", stroke: BLUE },
  ],
  finishMerge: [
    { d: "M2.5 4.5h7v11h-7zM10.5 4.5h7v11h-7z" },
    { d: "M5 8h2.5M5 10.5h2.5M13 8h2.5M13 10.5h2.5M8.5 2.5l3 0", stroke: BLUE },
  ],
} as const satisfies IconSet;
