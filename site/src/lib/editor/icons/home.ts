import type { IconSet } from "./types.js";

// Home tab icons after Word for Mac's: grey outlines, Word's accents (the
// painter's yellow bristles, the shading bucket's fill, the borders' frame).
export default {
  formatPainter: [
    { d: "M4 3.5h10v4H4z" },
    { d: "M14 5.5h2v4H9.5v2.5" },
    { d: "M8.5 12h2v5.5h-2z", fill: "#f2c94c", stroke: "#b8860b" },
  ],
  changeCase:
    "M2.5 15 6 5l3.5 10M3.7 11.5h4.6M14 9.5a2.2 2.2 0 1 0 0 4.4 2.2 2.2 0 0 0 0-4.4zM16.2 9v5.5",
  clearFormatting: [
    { d: "M3 15.5 7 5l4 10.5M4.4 12h5.2" },
    { d: "M12 14.5l3-3 2.5 2.5-3 3h-2.5z", stroke: "#d0517e" },
  ],
  phoneticGuide: [
    { d: "M3 5.5h2.5M8 5.5h2.5M13 5.5h3", stroke: "#8a8a8a" },
    {
      d: "M3.5 17l2.5-7 2.5 7M4.3 15h3.4M10.5 11h3a1.5 1.5 0 0 1 0 3h-3v-3zm0 3h3.5a1.5 1.5 0 0 1 0 3h-3.5z",
    },
  ],
  characterBorder: "M3 3.5h14v13H3zM6.5 14 10 6l3.5 8M7.7 11.3h4.6",
  textEffects: [{ d: "M4 16.5 10 3.5l6 13M6.3 11.5h7.4", stroke: "#2b6cd9" }],
  characterShading: [
    { d: "M3 3.5h14v13H3z", fill: "#bfbfbf", stroke: "#bfbfbf" },
    { d: "M6.5 14 10 6l3.5 8M7.7 11.3h4.6" },
  ],
  encloseCharacters: "M10 3a7 7 0 1 1 0 14 7 7 0 0 1 0-14zM8 7.5h4M10 7.5v6M8 13.5h4",
  multilevel: "M7 4.5h10.5M9.5 10h8M12 15.5h5.5M3 4.5h1M5.5 10h1M8 15.5h1",
  asianLayout: "M3.5 16.5 9 3.5l5.5 13M5.3 12.5h7.4M15.5 4v7M13.5 9l2 2 2-2",
  sortText: [
    { d: "M5 3.5v13M2.5 14l2.5 2.5L7.5 14" },
    { d: "M10.5 9 12.5 3.5 14.5 9M11.1 7.3h2.8M10.5 11.5h4l-4 5h4", stroke: "#2b6cd9" },
  ],
  alignDistributed: "M3 4.5h14M3 8.5h14M3 12.5h14M3 16.5h14M2 15l1 1.5L2 18M18 15l-1 1.5 1 1.5",
  paragraphShading: [
    { d: "M5 9.5 10 4.5l5 5-5 5z" },
    { d: "M16 12c.9 1.3 1.3 2.2 1.3 2.8a1.3 1.3 0 0 1-2.6 0c0-.6.4-1.5 1.3-2.8z", fill: "#f2c94c" },
    { d: "M3 17.5h14", stroke: "#f2c94c" },
  ],
  paragraphBorders: [
    { d: "M3 3.5h14v13H3z", stroke: "#a6a6a6" },
    { d: "M3 10h14M10 3.5v13", stroke: "#a6a6a6" },
    { d: "M3 16.5h14", stroke: "#333333" },
  ],
  stylesPane: [
    { d: "M4 2.5h9l3 3v12H4z" },
    { d: "M7 14l6.5-6.5 1.5 1.5L8.5 15.5H7z", stroke: "#2b6cd9" },
  ],
  findNext: "M5 8l5 5 5-5",
  findPrevious: "M5 12l5-5 5 5",
  findOptions: "M4 10h.01M10 10h.01M16 10h.01",
} as const satisfies IconSet;
