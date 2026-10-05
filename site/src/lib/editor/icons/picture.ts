import type { IconSet } from "./types.js";

// Picture Format, Chart Design, the Arrange group and Insert ▸ Illustrations.
// Grey outlines with Word's accents: blue for the picture's sky, green for
// its hills, orange for selection / emphasis.
const SKY = "#5B9BD5";
const HILL = "#70AD47";
const SUN = "#FFC000";
const ACCENT = "#ED7D31";
const BLUE = "#4472C4";

export default {
  illPictures: [
    { d: "M2.5 4h15v12h-15z" },
    { d: "M3 15l4.5-5 3.5 3.5 2-2 4 3.5", stroke: HILL },
    { d: "M13.5 7.5a1.5 1.5 0 1 0 0 .01", stroke: SUN, fill: SUN },
  ],
  illIcons: [
    { d: "M10 17s-6-3.6-6-8a3.2 3.2 0 0 1 6-1.6A3.2 3.2 0 0 1 16 9c0 4.4-6 8-6 8z", stroke: BLUE },
    { d: "M13 2.5l1 2 2 .3-1.5 1.4.4 2-1.9-1-1.9 1 .4-2L10 4.8l2-.3z", stroke: SUN },
  ],
  illChart: [
    { d: "M3 17.5h14.5" },
    { d: "M4.5 10h3v7.5h-3z", stroke: BLUE, fill: BLUE },
    { d: "M9 6h3v11.5H9z", stroke: ACCENT, fill: ACCENT },
    { d: "M13.5 12h3v5.5h-3z", stroke: "#A5A5A5", fill: "#A5A5A5" },
  ],
  illScreenshot: [
    { d: "M2.5 5.5h11v9h-11z", stroke: "#8c8c8c" },
    { d: "M7 8.5h10.5v8H7z" },
    { d: "M9.5 8.5l1-1.5h3l1 1.5M12 14a2 2 0 1 0 0-.01" },
    { d: "M16 3v4M14 5h4", stroke: HILL },
  ],
  removeBackground: [
    { d: "M3 4h14v12H3z", stroke: "#c43e96" },
    { d: "M6.5 16v-3.5a3.5 3.5 0 0 1 7 0V16M10 6.5a2 2 0 1 0 0 .01" },
  ],
  corrections: [
    { d: "M10 6.5a3.5 3.5 0 1 1 0 7 3.5 3.5 0 0 1 0-7z", stroke: SUN, fill: "#FFE699" },
    {
      d: "M10 2v2M10 16v2M2 10h2M16 10h2M4.3 4.3l1.4 1.4M14.3 14.3l1.4 1.4M4.3 15.7l1.4-1.4M14.3 5.7l1.4-1.4",
      stroke: SUN,
    },
  ],
  pictureColor: [
    { d: "M2.5 4h15v12h-15z" },
    { d: "M3 4h5v12H3z", stroke: SKY, fill: "#DEEBF7" },
    { d: "M8 4h5v12H8z", stroke: HILL, fill: "#E2F0D9" },
    { d: "M13 4h4.5v12H13z", stroke: ACCENT, fill: "#FBE5D6" },
  ],
  artisticEffects: [
    { d: "M2.5 4h15v12h-15z" },
    { d: "M4 14l3-4 2.5 2.5L12 9l4 5", stroke: HILL },
    { d: "M5 6.5h2M9 6.5h2M13 6.5h2", stroke: SKY },
  ],
  transparency: [
    { d: "M2.5 4h15v12h-15z" },
    { d: "M2.5 4h7.5v6H2.5zM10 10h7.5v6H10z", stroke: "#bfbfbf", fill: "#e6e6e6" },
    { d: "M4 15l4-4 3 3", stroke: HILL },
  ],
  compress: [
    { d: "M5 5h10v10H5z" },
    { d: "M2.5 2.5L6 6M17.5 2.5L14 6M2.5 17.5L6 14M17.5 17.5L14 14", stroke: BLUE },
  ],
  changePicture: [
    { d: "M2.5 4h10v8h-10z" },
    { d: "M3 11l3-3 2.5 2.5 1.5-1.5 2 2", stroke: HILL },
    { d: "M14 9.5a3.5 3.5 0 1 1-1 2.5M14 9.5v2.2h2.2", stroke: HILL },
  ],
  resetPicture: [
    { d: "M2.5 4h10v8h-10z" },
    { d: "M3 11l3-3 2.5 2.5 1.5-1.5 2 2", stroke: HILL },
    { d: "M17.5 14.5a3.5 3.5 0 1 1-1-2.5M17.5 11.5v2.6h-2.6", stroke: ACCENT },
  ],
  pictureBorder: [
    { d: "M3 3.5h11v9H3z" },
    { d: "M2 16.5h16", stroke: BLUE },
    { d: "M12 12.5l3.5-6 1.5 1-3.5 6z", stroke: ACCENT },
  ],
  pictureEffects: [
    { d: "M6 6h11v9H6z", stroke: "#bfbfbf", fill: "#e6e6e6" },
    { d: "M3 3.5h11v9H3z", fill: "#fff" },
    { d: "M4 11.5l3-3 2.5 2 2-2 2 2.5", stroke: HILL },
  ],
  altText: [
    { d: "M2.5 3.5h11v8h-11z" },
    { d: "M3 10.5l3-3 2.5 2.5 1.5-1.5 2.5 2.5", stroke: HILL },
    { d: "M9 13.5h8.5M9 16h8.5", stroke: BLUE },
  ],
  arrPosition: [
    { d: "M4 2.5h12v15H4z" },
    { d: "M6 5h4.5v4H6z", stroke: BLUE, fill: "#DAE3F3" },
    { d: "M12 5.5h2M12 8h2M6 11h8M6 13.5h8" },
  ],
  arrWrapText: [
    { d: "M2.5 4h15M2.5 16h15M2.5 8h3.5M14 8h3.5M2.5 12h3.5M14 12h3.5" },
    { d: "M7.5 6.5h5v7h-5z", stroke: BLUE, fill: "#DAE3F3" },
  ],
  arrBringForward: [
    { d: "M8 8h9.5v9.5H8z", stroke: "#bfbfbf" },
    { d: "M2.5 2.5h9.5V12H2.5z", stroke: BLUE, fill: "#DAE3F3" },
  ],
  arrSendBackward: [
    { d: "M2.5 2.5h9.5V12H2.5z", stroke: BLUE, fill: "#DAE3F3" },
    { d: "M8 8h9.5v9.5H8z", fill: "#fff" },
  ],
  arrSelectionPane: [
    { d: "M3 3.5h14v13H3zM11 3.5v13" },
    { d: "M12.5 6h3M12.5 8.5h3M12.5 11h3", stroke: BLUE },
    { d: "M5 7l4 1.5-1.8.7L9 11l-.8.6-1.8-1.9-.9 1.5z", stroke: ACCENT },
  ],
  arrAlign: [
    { d: "M3 2.5v15" },
    { d: "M4.5 5h9v3.5h-9zM4.5 11.5h6V15h-6z", stroke: BLUE, fill: "#DAE3F3" },
  ],
  arrGroup: [
    { d: "M2.5 2.5h15v15h-15z", stroke: "#bfbfbf" },
    { d: "M4.5 4.5h6v6h-6zM9.5 9.5h6v6h-6z", stroke: BLUE },
  ],
  arrRotate: [
    { d: "M4 16.5h12.5L4 4z", stroke: BLUE },
    { d: "M15 4.5a6 6 0 0 1 2.5 5M17.5 9.5l-2-.3.6-2" },
  ],
  picCrop: [{ d: "M5.5 2v12.5H18M2 5.5h12.5V18" }, { d: "M4 16l12-12", stroke: "#bfbfbf" }],
  picHeight: [
    { d: "M10 2.5v15M7 5l3-2.5L13 5M7 15l3 2.5 3-2.5" },
    { d: "M3 2.5h4M3 17.5h4", stroke: BLUE },
  ],
  picWidth: [
    { d: "M2.5 10h15M5 7l-2.5 3L5 13M15 7l2.5 3-2.5 3" },
    { d: "M2.5 3v4M17.5 3v4", stroke: BLUE },
  ],
  lockAspect: "M6.5 9V6.5a3.5 3.5 0 0 1 7 0V9M5 9h10v8H5z",
  formatPane: [
    { d: "M3 3.5h14v13H3zM11 3.5v13" },
    { d: "M12.5 6h3M12.5 8.5h3M12.5 11h3M12.5 13.5h3", stroke: BLUE },
  ],
  addChartElement: [
    { d: "M3 17.5h11" },
    { d: "M4.5 11h2.5v6.5H4.5zM8.5 7h2.5v10.5H8.5z", stroke: BLUE, fill: BLUE },
    { d: "M15.5 3v6M12.5 6h6", stroke: HILL },
  ],
  changeColors: [
    {
      d: "M10 3a7 7 0 1 0 0 14c1 0 1.5-.8 1-1.6-.5-.9 0-1.9 1-1.9h2a3 3 0 0 0 3-3A7.5 7.5 0 0 0 10 3z",
    },
    { d: "M6.5 9a1 1 0 1 0 0 .01", stroke: ACCENT, fill: ACCENT },
    { d: "M9 6a1 1 0 1 0 0 .01", stroke: SUN, fill: SUN },
    { d: "M13 7a1 1 0 1 0 0 .01", stroke: BLUE, fill: BLUE },
  ],
  switchRowCol: [
    { d: "M3 3h14v14H3zM3 8h14M8 3v14" },
    { d: "M11 11.5h4M13.5 10l1.5 1.5-1.5 1.5", stroke: BLUE },
  ],
  editData: [
    { d: "M2.5 3.5h15v13h-15zM2.5 7.5h15M2.5 11.5h15M7.5 3.5v13M12.5 3.5v13" },
    { d: "M2.5 3.5h15v4h-15z", stroke: HILL },
  ],
  changeChartType: [
    { d: "M2.5 17.5h15" },
    { d: "M4 10h3v7.5H4z", stroke: BLUE },
    { d: "M9.5 4a5 5 0 1 1 0 10V9z", stroke: ACCENT },
  ],
} as const satisfies IconSet;
