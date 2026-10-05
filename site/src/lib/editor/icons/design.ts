import type { IconSet } from "./types.js";

// The accent colours Word's Design-tab icons use (theme swatches, page colour).
const BLUE = "#2F6FBA";
const ORANGE = "#E97132";
const GREEN = "#4EA72E";
const GOLD = "#E0A82E";

export default {
  themes: [
    { d: "M3 3.5h14v13H3z" },
    { d: "M5 13h2.5v1.5H5z", fill: BLUE, stroke: BLUE },
    { d: "M8.75 13h2.5v1.5h-2.5z", fill: ORANGE, stroke: ORANGE },
    { d: "M12.5 13H15v1.5h-2.5z", fill: GREEN, stroke: GREEN },
    { d: "M5.5 10.5 8 5l2.5 5.5M6.4 8.6h3.2M12 7.5h2.5v3H12z" },
  ],
  themeColors: [
    { d: "M3 3h6v6H3z", fill: BLUE, stroke: BLUE },
    { d: "M11 3h6v6h-6z", fill: ORANGE, stroke: ORANGE },
    { d: "M3 11h6v6H3z", fill: GREEN, stroke: GREEN },
    { d: "M11 11h6v6h-6z", fill: GOLD, stroke: GOLD },
  ],
  themeFonts: [
    { d: "M2.5 15.5 7 4l4.5 11.5M4.2 11.5h5.6" },
    { d: "M17 10.5v5M17 12.5a2.5 2.5 0 1 0 0 1", stroke: BLUE },
  ],
  paragraphSpacing: [
    { d: "M8 4h9M8 8h9M8 12h9M8 16h9" },
    { d: "M4 3v14M2 5l2-2 2 2M2 15l2 2 2-2", stroke: BLUE },
  ],
  themeEffects: [{ d: "M4 5.5h9v9H4z" }, { d: "M6.5 3h11v11", stroke: BLUE }],
  setDefault: [{ d: "M4 2.5h8l3.5 3.5v11.5H4z" }, { d: "M7 11l2 2 4-4", stroke: GREEN }],
  watermark: [
    { d: "M4.5 2.5h8l3 3v12h-11z" },
    { d: "M6.5 13.5l7-7M6.5 9.5l3-3M10 13.5l3.5-3.5", stroke: "#A0A0A0" },
  ],
  pageColor: [
    { d: "M4.5 2.5h8l3 3v12h-11z" },
    {
      d: "M7.5 7.5 10 5l3 3-2.5 2.5zM13.5 10.5c.7 1 1 1.6 1 2a1 1 0 0 1-2 0c0-.4.3-1 1-2z",
      stroke: BLUE,
    },
  ],
  pageBorders: [{ d: "M5 2.5h10v15H5z" }, { d: "M3 1v18M17 1v18M3 1h14M3 19h14", stroke: BLUE }],
} as const satisfies IconSet;
