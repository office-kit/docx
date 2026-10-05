/**
 * Insert ▸ Icons: a small built-in library of filled glyphs on a 24 × 24 grid.
 * Inserted icons are rasterized to PNG because SVG pictures need the
 * `asvg:svgBlip` Office extension.
 */

export interface LibraryIcon {
  readonly id: string;
  readonly category: "people" | "objects" | "symbols" | "nature";
  readonly path: string;
}

export const ICON_LIBRARY: readonly LibraryIcon[] = [
  {
    id: "person",
    category: "people",
    path: "M12 12a5 5 0 1 0 0-10 5 5 0 0 0 0 10zm0 2c-5 0-9 2.5-9 6v2h18v-2c0-3.5-4-6-9-6z",
  },
  {
    id: "group",
    category: "people",
    path: "M8 11a4 4 0 1 0 0-8 4 4 0 0 0 0 8zm8 0a3.5 3.5 0 1 0 0-7 3.5 3.5 0 0 0 0 7zM8 13c-4 0-7 2-7 5v3h14v-3c0-3-3-5-7-5zm8 0c-.7 0-1.3.1-1.9.2A6 6 0 0 1 17 18v3h6v-3c0-3-3-5-7-5z",
  },
  {
    id: "handshake",
    category: "people",
    path: "M2 8l5-3 4 2 2-1 9 5-3 3-6-3-2 1 5 4-2 2-5-3-2 2-6-5z",
  },
  { id: "house", category: "objects", path: "M12 3l10 9h-3v9h-5v-6h-4v6H5v-9H2z" },
  {
    id: "building",
    category: "objects",
    path: "M4 22V3h11v5h5v14h-6v-4h-4v4zM7 6v2h2V6zm4 0v2h2V6zM7 10v2h2v-2zm4 0v2h2v-2zm6 2v2h2v-2zm0 4v2h2v-2zM7 14v2h2v-2zm4 0v2h2v-2z",
  },
  {
    id: "phone",
    category: "objects",
    path: "M7 2h10a2 2 0 0 1 2 2v16a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2zm0 3v13h10V5zm5 14.5a1 1 0 1 0 0 .01z",
  },
  { id: "laptop", category: "objects", path: "M4 5h16v10H4zm2 2v6h12V7zM1 17h22l-2 3H3z" },
  {
    id: "envelope",
    category: "objects",
    path: "M2 5h20v14H2zm2 2v.5l8 5.5 8-5.5V7zm0 3v7h16v-7l-8 5.5z",
  },
  {
    id: "calendar",
    category: "objects",
    path: "M3 4h3V2h2v2h8V2h2v2h3v18H3zm2 6v10h14V10zm2 2h3v3H7z",
  },
  {
    id: "lightbulb",
    category: "objects",
    path: "M12 2a7 7 0 0 0-4 12.7V18h8v-3.3A7 7 0 0 0 12 2zM9 19h6v1.5a1.5 1.5 0 0 1-1.5 1.5h-3A1.5 1.5 0 0 1 9 20.5z",
  },
  {
    id: "chart",
    category: "objects",
    path: "M3 3h2v16h16v2H3zm4 9h3v6H7zm5-5h3v11h-3zm5 3h3v8h-3z",
  },
  {
    id: "gear",
    category: "objects",
    path: "M10 2h4l.6 2.6 2.2 1 2.4-1.3 2.8 2.8-1.3 2.4 1 2.2L24 12v0l-2.6.6-1 2.2 1.3 2.4-2.8 2.8-2.4-1.3-2.2 1L14 22h-4l-.6-2.6-2.2-1-2.4 1.3-2.8-2.8 1.3-2.4-1-2.2L0 12l2.6-.6 1-2.2-1.3-2.4 2.8-2.8 2.4 1.3 2.2-1zm2 6a4 4 0 1 0 0 8 4 4 0 0 0 0-8z",
  },
  {
    id: "check",
    category: "symbols",
    path: "M12 1a11 11 0 1 0 0 22 11 11 0 0 0 0-22zm-1.5 15.5L5 11l1.5-1.5 4 4 7-7L19 8z",
  },
  {
    id: "cross",
    category: "symbols",
    path: "M12 1a11 11 0 1 0 0 22 11 11 0 0 0 0-22zm4.5 14l-1.5 1.5-3-3-3 3L7.5 15l3-3-3-3L9 7.5l3 3 3-3L16.5 9l-3 3z",
  },
  { id: "warning", category: "symbols", path: "M12 2l11 20H1zm-1 7v6h2V9zm0 8v2h2v-2z" },
  {
    id: "info",
    category: "symbols",
    path: "M12 1a11 11 0 1 0 0 22 11 11 0 0 0 0-22zm-1 5h2v2h-2zm0 4h2v8h-2z",
  },
  {
    id: "star",
    category: "symbols",
    path: "M12 2l3 6.5 7 .8-5.2 4.8 1.4 7L12 17.6 5.8 21l1.4-7L2 9.3l7-.8z",
  },
  {
    id: "heart",
    category: "symbols",
    path: "M12 21S2 14.5 2 8a5 5 0 0 1 10-1.5A5 5 0 0 1 22 8c0 6.5-10 13-10 13z",
  },
  { id: "arrowRight", category: "symbols", path: "M2 10h13V5l7 7-7 7v-5H2z" },
  {
    id: "location",
    category: "symbols",
    path: "M12 1a8 8 0 0 0-8 8c0 6 8 14 8 14s8-8 8-14a8 8 0 0 0-8-8zm0 11a3 3 0 1 1 0-6 3 3 0 0 1 0 6z",
  },
  {
    id: "sun",
    category: "nature",
    path: "M12 7a5 5 0 1 0 0 10 5 5 0 0 0 0-10zM11 0h2v4h-2zm0 20h2v4h-2zM0 11h4v2H0zm20 0h4v2h-4zM3.5 2.1l2.8 2.8-1.4 1.4-2.8-2.8zm14.2 14.2l2.8 2.8-1.4 1.4-2.8-2.8zM2.1 20.5l2.8-2.8 1.4 1.4-2.8 2.8zM16.3 6.3l2.8-2.8 1.4 1.4-2.8 2.8z",
  },
  {
    id: "cloud",
    category: "nature",
    path: "M6 19a5 5 0 0 1-.6-10A7 7 0 0 1 19 8a5.5 5.5 0 0 1-.5 11z",
  },
  { id: "tree", category: "nature", path: "M12 1l7 10h-3l5 7h-8v5h-2v-5H3l5-7H5z" },
  {
    id: "leaf",
    category: "nature",
    path: "M21 3C10 3 3 8 3 16c0 2 .5 3.5 1 4.5L2 22l1.5 1L5 21c1 .5 2.5 1 4 1 8 0 12-8 12-19zM7 18c2-5 6-8 10-10-3 3-6 6-8 11z",
  },
];

/** The icon as a standalone SVG document, black on transparent. */
export function iconSvg(icon: LibraryIcon, size: number): string {
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" width="${size}" height="${size}"><path d="${icon.path}" fill="#000"/></svg>`;
}
