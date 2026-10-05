/**
 * SVG outlines for the DrawingML preset shapes (ECMA-376 Part 1, §20.1.9.18,
 * ST_ShapeType) a picture can be cropped to, at their default adjust values.
 * Shapes outside this set draw as their bounding rectangle.
 */

/** The shapes Picture Format ▸ Crop ▸ Crop to Shape offers, in gallery order. */
export const CROP_SHAPES = [
  "rect",
  "roundRect",
  "snip1Rect",
  "snip2SameRect",
  "snip2DiagRect",
  "snipRoundRect",
  "round1Rect",
  "round2SameRect",
  "round2DiagRect",
  "ellipse",
  "triangle",
  "rtTriangle",
  "parallelogram",
  "trapezoid",
  "diamond",
  "pentagon",
  "hexagon",
  "heptagon",
  "octagon",
  "decagon",
  "dodecagon",
  "teardrop",
  "plaque",
  "heart",
  "star4",
  "star5",
  "star6",
  "star8",
  "star10",
  "star12",
  "rightArrow",
  "leftArrow",
  "upArrow",
  "downArrow",
  "leftRightArrow",
  "homePlate",
  "chevron",
] as const;

// Shape guides use a 100000 scale; most corner adjustments default to 16667.
const DEFAULT_CORNER = 0.16667;

type Pt = readonly [number, number];

function poly(points: readonly Pt[]): string {
  return `M${points.map(([x, y]) => `${round(x)},${round(y)}`).join("L")}Z`;
}

function round(n: number): number {
  return Math.round(n * 100) / 100;
}

/** A regular polygon (or star when `inner` is given) inscribed in the box. */
function regular(w: number, h: number, sides: number, inner?: number): string {
  const pts: Pt[] = [];
  const count = inner === undefined ? sides : sides * 2;
  for (let i = 0; i < count; i++) {
    const r = inner !== undefined && i % 2 === 1 ? inner : 1;
    const a = -Math.PI / 2 + (i * 2 * Math.PI) / count;
    pts.push([w / 2 + (w / 2) * r * Math.cos(a), h / 2 + (h / 2) * r * Math.sin(a)]);
  }
  return poly(pts);
}

/** Rounded (`r`) or snipped (`s`) corners, clockwise from top-left; 0 leaves a corner square. */
function corners(w: number, h: number, kinds: string, size: number): string {
  const [tl, tr, br, bl] = [...kinds];
  const d = Math.min(w, h) * size;
  const seg = (
    kind: string | undefined,
    x: number,
    y: number,
    dx1: number,
    dy1: number,
    dx2: number,
    dy2: number,
  ): string => {
    if (kind === "r")
      return `L${round(x + dx1 * d)},${round(y + dy1 * d)}Q${round(x)},${round(y)} ${round(x + dx2 * d)},${round(y + dy2 * d)}`;
    if (kind === "s")
      return `L${round(x + dx1 * d)},${round(y + dy1 * d)}L${round(x + dx2 * d)},${round(y + dy2 * d)}`;
    return `L${round(x)},${round(y)}`;
  };
  return (
    `M${round(tl === "0" ? 0 : d)},0` +
    seg(tr, w, 0, -1, 0, 0, 1) +
    seg(br, w, h, 0, -1, -1, 0) +
    seg(bl, 0, h, 1, 0, 0, -1) +
    seg(tl, 0, 0, 0, 1, 1, 0) +
    "Z"
  );
}

function arrow(w: number, h: number, dir: "r" | "l" | "u" | "d"): string {
  // Default adjusts: shaft 50 % of the cross size, head 50 % of the cross size long.
  const horizontal = dir === "r" || dir === "l";
  const len = horizontal ? w : h;
  const cross = horizontal ? h : w;
  const head = Math.min(len, cross * 0.5);
  const shaft = [cross * 0.25, cross * 0.75] as const;
  const pts: Pt[] = [
    [0, shaft[0]],
    [len - head, shaft[0]],
    [len - head, 0],
    [len, cross / 2],
    [len - head, cross],
    [len - head, shaft[1]],
    [0, shaft[1]],
  ];
  const map = ([a, b]: Pt): Pt => {
    switch (dir) {
      case "r":
        return [a, b];
      case "l":
        return [w - a, b];
      case "d":
        return [b, a];
      default:
        return [b, h - a];
    }
  };
  return poly(pts.map(map));
}

/** The SVG path of preset `prst` in a `w` × `h` box. */
export function presetPath(prst: string, w: number, h: number): string {
  switch (prst) {
    case "roundRect":
      return corners(w, h, "rrrr", DEFAULT_CORNER);
    case "snip1Rect":
      return corners(w, h, "0s00", DEFAULT_CORNER);
    case "snip2SameRect":
      return corners(w, h, "ss00", DEFAULT_CORNER);
    case "snip2DiagRect":
      return corners(w, h, "s0s0", DEFAULT_CORNER);
    case "snipRoundRect":
      return corners(w, h, "rs00", DEFAULT_CORNER);
    case "round1Rect":
      return corners(w, h, "0r00", DEFAULT_CORNER);
    case "round2SameRect":
      return corners(w, h, "rr00", DEFAULT_CORNER);
    case "round2DiagRect":
      return corners(w, h, "r0r0", DEFAULT_CORNER);
    case "ellipse":
      return `M0,${h / 2}A${w / 2},${h / 2} 0 1 0 ${w},${h / 2}A${w / 2},${h / 2} 0 1 0 0,${h / 2}Z`;
    case "triangle":
      return poly([
        [w / 2, 0],
        [w, h],
        [0, h],
      ]);
    case "rtTriangle":
      return poly([
        [0, 0],
        [w, h],
        [0, h],
      ]);
    case "parallelogram": {
      const o = Math.min(w, h) * 0.25;
      return poly([
        [o, 0],
        [w, 0],
        [w - o, h],
        [0, h],
      ]);
    }
    case "trapezoid": {
      const o = Math.min(w, h) * 0.25;
      return poly([
        [o, 0],
        [w - o, 0],
        [w, h],
        [0, h],
      ]);
    }
    case "diamond":
      return poly([
        [w / 2, 0],
        [w, h / 2],
        [w / 2, h],
        [0, h / 2],
      ]);
    case "pentagon":
      return regular(w, h, 5);
    case "hexagon":
      return poly([
        [w * 0.25, 0],
        [w * 0.75, 0],
        [w, h / 2],
        [w * 0.75, h],
        [w * 0.25, h],
        [0, h / 2],
      ]);
    case "heptagon":
      return regular(w, h, 7);
    case "octagon": {
      const o = Math.min(w, h) * 0.29289;
      return poly([
        [o, 0],
        [w - o, 0],
        [w, o],
        [w, h - o],
        [w - o, h],
        [o, h],
        [0, h - o],
        [0, o],
      ]);
    }
    case "decagon":
      return regular(w, h, 10);
    case "dodecagon":
      return regular(w, h, 12);
    case "teardrop":
      return `M0,${h / 2}A${w / 2},${h / 2} 0 0 1 ${w / 2},0L${w},0L${w},${h / 2}A${w / 2},${h / 2} 0 1 1 0,${h / 2}Z`;
    case "plaque": {
      const r = Math.min(w, h) * DEFAULT_CORNER;
      return `M${r},0L${w - r},0A${r},${r} 0 0 0 ${w},${r}L${w},${h - r}A${r},${r} 0 0 0 ${w - r},${h}L${r},${h}A${r},${r} 0 0 0 0,${h - r}L0,${r}A${r},${r} 0 0 0 ${r},0Z`;
    }
    case "heart":
      return `M${w / 2},${h * 0.25}C${w / 2},0 0,0 0,${h * 0.3}C0,${h * 0.6} ${w / 2},${h * 0.8} ${w / 2},${h}C${w / 2},${h * 0.8} ${w},${h * 0.6} ${w},${h * 0.3}C${w},0 ${w / 2},0 ${w / 2},${h * 0.25}Z`;
    case "star4":
      return regular(w, h, 4, 0.25);
    case "star5":
      return regular(w, h, 5, 0.382);
    case "star6":
      return regular(w, h, 6, 0.5);
    case "star8":
      return regular(w, h, 8, 0.75);
    case "star10":
      return regular(w, h, 10, 0.8);
    case "star12":
      return regular(w, h, 12, 0.75);
    case "rightArrow":
      return arrow(w, h, "r");
    case "leftArrow":
      return arrow(w, h, "l");
    case "upArrow":
      return arrow(w, h, "u");
    case "downArrow":
      return arrow(w, h, "d");
    case "leftRightArrow": {
      const head = Math.min(w / 2, h * 0.5);
      return poly([
        [0, h / 2],
        [head, 0],
        [head, h * 0.25],
        [w - head, h * 0.25],
        [w - head, 0],
        [w, h / 2],
        [w - head, h],
        [w - head, h * 0.75],
        [head, h * 0.75],
        [head, h],
      ]);
    }
    case "homePlate": {
      const o = Math.min(w, h * 0.5);
      return poly([
        [0, 0],
        [w - o, 0],
        [w, h / 2],
        [w - o, h],
        [0, h],
      ]);
    }
    case "chevron": {
      const o = Math.min(w, h * 0.5);
      return poly([
        [0, 0],
        [w - o, 0],
        [w, h / 2],
        [w - o, h],
        [0, h],
        [o, h / 2],
      ]);
    }
    default:
      return `M0,0H${w}V${h}H0Z`;
  }
}
