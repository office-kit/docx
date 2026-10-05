import type { IconSet } from "./types.js";

// Word's accent colours on drawing icons.
const BLUE = "#2B78C5";
const LIGHT_BLUE = "#9CC3E9";
const RED = "#D13438";
const ORANGE = "#E97A3B";
const PURPLE = "#B263C8";
const GREEN = "#3A9E4A";
const YELLOW = "#F8D648";

// Shapes, text boxes, WordArt, SmartArt (Insert, Shape Format, SmartArt Design) and the Draw tab.
export default {
  shapes: [
    { d: "M2.5 3h8v7h-8z", fill: LIGHT_BLUE, stroke: BLUE },
    { d: "M12 6.5a4.5 4.5 0 1 1 0 9 4.5 4.5 0 0 1 0-9z", fill: "#fff" },
  ],
  textBox: [{ d: "M3 3.5h14v13H3z" }, { d: "M6.5 7h7M10 7v7", stroke: BLUE }],
  wordArt: [{ d: "M4 17 10 3l6 14M6.5 11.5h7", stroke: BLUE }, { d: "M3 17h3M14 17h3" }],
  smartArt: [
    { d: "M7.5 2.5h5v4h-5z", fill: GREEN, stroke: GREEN },
    { d: "M2.5 13h5v4h-5zM12.5 13h5v4h-5z", fill: LIGHT_BLUE, stroke: BLUE },
    { d: "M10 6.5v3M5 13V9.5h10V13" },
  ],
  editShape: [
    { d: "M3 15 8 4l5 8 4-3", stroke: BLUE },
    { d: "M2 14h2v2H2zM7 3h2v2H7zM12 11h2v2h-2zM16 8h2v2h-2z", fill: "#fff" },
  ],
  shapeFill: [
    { d: "M4 9 9.5 3.5 15 9l-5.5 5.5zM9.5 3.5 7.5 1.5" },
    { d: "M15.5 11s1.5 2 1.5 3a1.5 1.5 0 0 1-3 0c0-1 1.5-3 1.5-3z", fill: BLUE, stroke: BLUE },
    { d: "M2.5 17h15", stroke: ORANGE },
  ],
  shapeOutline: [{ d: "M4 13 13 4l2.5 2.5-9 9H4z" }, { d: "M2.5 17.5h15", stroke: BLUE }],
  shapeEffects: [
    { d: "M4 4h9v9H4z", fill: "#fff" },
    { d: "M13 7h3v9H7v-3", fill: "#999", stroke: "#999" },
  ],
  textFill: [{ d: "M4.5 15 10 3l5.5 12M6.5 11h7" }, { d: "M3 17.5h14", stroke: BLUE }],
  textOutline: [{ d: "M4.5 15 10 3l5.5 12M6.5 11h7", stroke: BLUE }, { d: "M3 17.5h14" }],
  textEffects: [
    { d: "M4.5 15 10 3l5.5 12M6.5 11h7", stroke: BLUE },
    { d: "M6 17h11", stroke: "#999" },
  ],
  textDirection: "M4 3v14M2 15l2 2 2-2M9 4h8M13 4v11M10 15h6",
  alignText: "M3 3.5h14M3 16.5h14M6 7.5h8M6 10h8M6 12.5h8",
  createLink: [
    { d: "M2.5 4h6v5h-6zM11.5 11h6v5h-6z" },
    { d: "M8.5 6.5h2a1.5 1.5 0 0 1 1.5 1.5v3", stroke: BLUE },
  ],
  altText: [
    { d: "M2.5 4h15v10.5h-15z" },
    { d: "M5 12l3-3 2 2 2-3 3 4", stroke: BLUE },
    { d: "M5 17h10" },
  ],
  formatPane: [{ d: "M3 3h14v14H3zM11 3v14" }, { d: "M13 6h2.5M13 9h2.5M13 12h2.5", stroke: BLUE }],
  addShape: [
    { d: "M2.5 6h8v8h-8z", fill: LIGHT_BLUE, stroke: BLUE },
    { d: "M15 7v6M12 10h6", stroke: GREEN },
  ],
  promote: "M17 5H8M17 10H8M17 15H8M6 7.5 3 10l3 2.5",
  demote: "M17 5H8M17 10H8M17 15H8M3 7.5 6 10l-3 2.5",
  textPane: [
    { d: "M3 3.5h14v13H3z" },
    { d: "M5.5 7h.5M8 7h7M7.5 10h.5M10 10h5M5.5 13h.5M8 13h7", stroke: BLUE },
  ],
  moveUp: [{ d: "M3 9h14v7H3z" }, { d: "M10 7V2M7.5 4.5 10 2l2.5 2.5", stroke: BLUE }],
  moveDown: [{ d: "M3 4h14v7H3z" }, { d: "M10 13v5M7.5 15.5 10 18l2.5-2.5", stroke: BLUE }],
  changeColors: [
    { d: "M3 3h6v6H3z", fill: BLUE, stroke: BLUE },
    { d: "M11 3h6v6h-6z", fill: ORANGE, stroke: ORANGE },
    { d: "M3 11h6v6H3z", fill: GREEN, stroke: GREEN },
    { d: "M11 11h6v6h-6z", fill: YELLOW, stroke: YELLOW },
  ],
  resetGraphic: [{ d: "M4 10a6 6 0 1 0 2-4.5M4 3v3.5h3.5", stroke: BLUE }],
  drawSelect: [{ d: "M5 3v12l3.5-3 2.5 5.5 2-1-2.5-5.5H15z", fill: "#fff" }],
  lasso: [
    { d: "M10 4c4.5 0 7.5 2 7.5 4.5S14.5 13 10 13 2.5 11 2.5 8.5 5.5 4 10 4z", stroke: BLUE },
    { d: "M6 12.5c-1 1.5 0 4 3 4" },
  ],
  eraser: [
    { d: "M3 13 10.5 5.5l5 5L10 16H6z", fill: PURPLE, stroke: PURPLE },
    { d: "M6.5 9.5l5 5M3 17.5h14" },
  ],
  addPen: [{ d: "M10 3v14M3 10h14", stroke: GREEN }],
  inkToShape: [
    { d: "M3 14c2-6 4-8 6-3s3 4 4 0", stroke: RED },
    { d: "M12 3h5.5v5.5H12z", fill: LIGHT_BLUE, stroke: BLUE },
  ],
  drawingCanvas: [
    { d: "M2.5 4h15v12h-15z", fill: "#fff" },
    { d: "M5 13c2-4 4-5 6-2s3 1 4-2", stroke: RED },
  ],
} as const satisfies IconSet;
