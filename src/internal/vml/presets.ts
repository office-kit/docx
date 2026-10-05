/**
 * Word's Insert ▸ Shapes gallery, expressed in VML (ECMA-376 Part 4, §19.1).
 *
 * Each preset is a path in VML's path language on a 21600 × 21600 coordinate
 * space (§19.1.2.14 `path`), written with the shape's default adjust values
 * baked in. Shapes that exist among VML's legacy shape types also carry their
 * `o:spt` number (§19.2.2.x, `ST_...` shape type) so Word identifies them as
 * that preset; shapes Word added with DrawingML only (the equation shapes,
 * for example) have no shape type number and are written as custom geometry,
 * which is the only way VML can express them.
 *
 * VML's straight-line primitive `v:line` is not used: its geometry lives in
 * `from` / `to` coordinates rather than in the box every other shape is
 * edited through.
 */

export const COORD_SIZE = 21600;
const C = COORD_SIZE;
const H = C / 2;
// Control-point distance that makes four cubic Béziers approximate a circle.
const KAPPA = 0.5523;

type Pt = readonly [number, number];

const r = (n: number): number => Math.round(n);
const pt = ([x, y]: Pt): string => `${r(x)},${r(y)}`;

/** A closed polygon subpath. */
function poly(points: readonly Pt[]): string {
  const [first, ...rest] = points;
  if (!first) return "";
  return `m${pt(first)}l${rest.map(pt).join(",")}x`;
}

/** An open polyline subpath (not closed, not filled). */
function line(points: readonly Pt[]): string {
  const [first, ...rest] = points;
  if (!first) return "";
  return `m${pt(first)}l${rest.map(pt).join(",")}`;
}

/** A closed ellipse subpath from four cubic Béziers. */
function ellipse(cx: number, cy: number, rx: number, ry: number): string {
  const kx = rx * KAPPA;
  const ky = ry * KAPPA;
  const c = (...ps: Pt[]): string => `c${ps.map(pt).join(",")}`;
  return (
    `m${pt([cx + rx, cy])}` +
    c([cx + rx, cy + ky], [cx + kx, cy + ry], [cx, cy + ry]) +
    c([cx - kx, cy + ry], [cx - rx, cy + ky], [cx - rx, cy]) +
    c([cx - rx, cy - ky], [cx - kx, cy - ry], [cx, cy - ry]) +
    c([cx + kx, cy - ry], [cx + rx, cy - ky], [cx + rx, cy]) +
    "x"
  );
}

/** A star with `n` points whose inner vertices sit at `inner` × the outer radius. */
function star(n: number, inner: number): string {
  const points: Pt[] = [];
  for (let i = 0; i < n * 2; i++) {
    const radius = i % 2 === 0 ? 1 : inner;
    const angle = -Math.PI / 2 + (i * Math.PI) / n;
    points.push([H + H * radius * Math.cos(angle), H + H * radius * Math.sin(angle)]);
  }
  return poly(points);
}

/** A rectangle with quarter-ellipse corners of radius `rad`. */
function roundRect(rad: number): string {
  return (
    `m${rad},0l${C - rad},0qx${C},${rad}l${C},${C - rad}qy${C - rad},${C}` +
    `l${rad},${C}qx0,${C - rad}l0,${rad}qy${rad},0x`
  );
}

/** Transpose a polygon (swap x/y) — a right arrow becomes a down arrow. */
const transpose = (points: readonly Pt[]): Pt[] => points.map(([x, y]) => [y, x]);
const mirrorX = (points: readonly Pt[]): Pt[] => points.map(([x, y]) => [C - x, y]);
const mirrorY = (points: readonly Pt[]): Pt[] => points.map(([x, y]) => [x, C - y]);

const RIGHT_ARROW: Pt[] = [
  [0, 5400],
  [16200, 5400],
  [16200, 0],
  [C, H],
  [16200, C],
  [16200, 16200],
  [0, 16200],
];
const RIGHT_ARROW_CALLOUT: Pt[] = [
  [0, 0],
  [14400, 0],
  [14400, 8100],
  [18000, 8100],
  [18000, 5400],
  [C, H],
  [18000, 16200],
  [18000, 13500],
  [14400, 13500],
  [14400, C],
  [0, C],
];
const LEFT_RIGHT_ARROW: Pt[] = [
  [0, H],
  [4320, 0],
  [4320, 5400],
  [17280, 5400],
  [17280, 0],
  [C, H],
  [17280, C],
  [17280, 16200],
  [4320, 16200],
  [4320, C],
];

function quadArrow(): Pt[] {
  const a = 6480; // shaft half-width edge
  const b = 8640; // head base
  const head = 2160;
  const out: Pt[] = [];
  // Built clockwise from the top point, one arm per quadrant.
  const arm = (rot: (p: Pt) => Pt): Pt[] =>
    (
      [
        [H, 0],
        [b + 4320, head * 2],
        [H + 1080, head * 2],
        [H + 1080, a + 2160],
      ] as Pt[]
    ).map(rot);
  const rots: Array<(p: Pt) => Pt> = [
    (p) => p,
    ([x, y]) => [C - y, x],
    ([x, y]) => [C - x, C - y],
    ([x, y]) => [y, C - x],
  ];
  for (const rot of rots) out.push(...arm(rot));
  return out;
}

function sun(): string {
  const rays: string[] = [];
  for (let i = 0; i < 8; i++) {
    const a = (i * Math.PI) / 4;
    const spread = Math.PI / 16;
    const tip: Pt = [H + H * Math.cos(a), H + H * Math.sin(a)];
    const left: Pt = [H + 7400 * Math.cos(a - spread), H + 7400 * Math.sin(a - spread)];
    const right: Pt = [H + 7400 * Math.cos(a + spread), H + 7400 * Math.sin(a + spread)];
    rays.push(poly([tip, left, right]));
  }
  return rays.join("") + ellipse(H, H, 5400, 5400);
}

function cloud(): string {
  // A ring of overlapping lobes; their union reads as Word's cloud.
  const lobes: string[] = [];
  const count = 9;
  for (let i = 0; i < count; i++) {
    const a = (i * 2 * Math.PI) / count;
    lobes.push(ellipse(H + 7000 * Math.cos(a), H + 6200 * Math.sin(a), 4300, 4300));
  }
  return lobes.join("") + ellipse(H, H, 8000, 7000);
}

function wedgeEllipse(): string {
  // Ellipse body plus a tail toward the default adjust point (1350, 25920).
  return (
    ellipse(H, H, H, H) +
    poly([
      [6300, 19400],
      [1350, 25920],
      [3400, 17200],
    ])
  );
}

/** Gallery categories, in the order of Word's Shapes menu. */
export const SHAPE_CATEGORIES = [
  "lines",
  "rectangles",
  "basicShapes",
  "blockArrows",
  "equationShapes",
  "flowchart",
  "starsAndBanners",
  "callouts",
] as const;
export type ShapeCategory = (typeof SHAPE_CATEGORIES)[number];

/**
 * How a preset is written. `element` names the VML primitive (§19.1.2): most
 * presets are a `v:shape` with a `path`, but Word writes rectangles, ellipses
 * and straight lines as `v:rect` / `v:oval` / `v:line`, and so do we.
 */
export interface ShapePresetDef {
  readonly category: ShapeCategory;
  readonly element: "shape" | "rect" | "roundrect" | "oval";
  /** VML shape type number (`o:spt`), when the preset is a legacy VML shape type. */
  readonly spt?: number;
  /** Path on the 21600 × 21600 coordinate space (`element: "shape"` only). */
  readonly path?: string;
  /** Lines and brackets are outlines only. */
  readonly filled?: boolean;
  /** One-dimensional shapes (lines, connectors) are sized by their end points (`o:oned`). */
  readonly oneD?: boolean;
  readonly startArrow?: boolean;
  readonly endArrow?: boolean;
}

const shape = (
  category: ShapeCategory,
  path: string,
  spt?: number,
  extra: Partial<ShapePresetDef> = {},
): ShapePresetDef => ({
  category,
  element: "shape",
  path: `${path}e`,
  ...(spt === undefined ? {} : { spt }),
  ...extra,
});

const BRACKET = 1800;
const BRACE = 1800;

function buildPresets() {
  return {
    // --- Lines ---------------------------------------------------------------
    // Word writes straight lines as its "straight connector" shape type (32), a
    // one-dimensional shape whose direction comes from the box's flip.
    line: shape(
      "lines",
      line([
        [0, 0],
        [C, C],
      ]),
      32,
      { filled: false, oneD: true },
    ),
    arrow: shape(
      "lines",
      line([
        [0, 0],
        [C, C],
      ]),
      32,
      { filled: false, oneD: true, endArrow: true },
    ),
    doubleArrow: shape(
      "lines",
      line([
        [0, 0],
        [C, C],
      ]),
      32,
      {
        filled: false,
        oneD: true,
        startArrow: true,
        endArrow: true,
      },
    ),
    elbowConnector: shape(
      "lines",
      line([
        [0, 0],
        [H, 0],
        [H, C],
        [C, C],
      ]),
      34,
      {
        filled: false,
        oneD: true,
      },
    ),
    curvedConnector: shape("lines", `m0,0c${H},0,${H},${C},${C},${C}`, 38, {
      filled: false,
      oneD: true,
    }),
    // Freeform and scribble are drawn point by point; the path is replaced by
    // the drawn points. The gallery icon uses this sample stroke.
    freeform: shape(
      "lines",
      line([
        [0, C],
        [5400, 4000],
        [12000, 15000],
        [C, 0],
      ]),
      undefined,
      {
        filled: false,
      },
    ),
    scribble: shape(
      "lines",
      "m0,16000c3000,0,4000,-6000,9000,2000c5000,8000,9000,-14000,12600,-4000",
      undefined,
      {
        filled: false,
      },
    ),

    // --- Rectangles ----------------------------------------------------------
    rect: { category: "rectangles", element: "rect", spt: 1 },
    roundRect: { category: "rectangles", element: "roundrect", spt: 2 },

    // --- Basic shapes --------------------------------------------------------
    textBox: shape(
      "basicShapes",
      poly([
        [0, 0],
        [0, C],
        [C, C],
        [C, 0],
      ]),
      202,
    ),
    ellipse: { category: "basicShapes", element: "oval", spt: 3 },
    triangle: shape(
      "basicShapes",
      poly([
        [H, 0],
        [0, C],
        [C, C],
      ]),
      5,
    ),
    rightTriangle: shape(
      "basicShapes",
      poly([
        [0, 0],
        [0, C],
        [C, C],
      ]),
      6,
    ),
    parallelogram: shape(
      "basicShapes",
      poly([
        [5400, 0],
        [C, 0],
        [16200, C],
        [0, C],
      ]),
      7,
    ),
    trapezoid: shape(
      "basicShapes",
      poly([
        [5400, 0],
        [16200, 0],
        [C, C],
        [0, C],
      ]),
    ),
    diamond: shape(
      "basicShapes",
      poly([
        [H, 0],
        [C, H],
        [H, C],
        [0, H],
      ]),
      4,
    ),
    pentagon: shape(
      "basicShapes",
      poly([
        [H, 0],
        [C, 8250],
        [17550, C],
        [4050, C],
        [0, 8250],
      ]),
      56,
    ),
    hexagon: shape(
      "basicShapes",
      poly([
        [5400, 0],
        [16200, 0],
        [C, H],
        [16200, C],
        [5400, C],
        [0, H],
      ]),
      9,
    ),
    octagon: shape(
      "basicShapes",
      poly([
        [6326, 0],
        [C - 6326, 0],
        [C, 6326],
        [C, C - 6326],
        [C - 6326, C],
        [6326, C],
        [0, C - 6326],
        [0, 6326],
      ]),
      10,
    ),
    cross: shape(
      "basicShapes",
      poly([
        [5400, 0],
        [16200, 0],
        [16200, 5400],
        [C, 5400],
        [C, 16200],
        [16200, 16200],
        [16200, C],
        [5400, C],
        [5400, 16200],
        [0, 16200],
        [0, 5400],
        [5400, 5400],
      ]),
      11,
    ),
    frame: shape(
      "basicShapes",
      poly([
        [0, 0],
        [C, 0],
        [C, C],
        [0, C],
      ]) +
        poly([
          [2700, 2700],
          [2700, 18900],
          [18900, 18900],
          [18900, 2700],
        ]),
    ),
    donut: shape("basicShapes", ellipse(H, H, H, H) + ellipse(H, H, 5400, 5400), 23),
    foldedCorner: shape(
      "basicShapes",
      poly([
        [0, 0],
        [C, 0],
        [C, 18000],
        [18000, C],
        [0, C],
      ]) +
        poly([
          [18000, C],
          [18720, 18720],
          [C, 18000],
        ]),
      65,
    ),
    plaque: shape(
      "basicShapes",
      `m3600,0l18000,0qy${C},3600l${C},18000qx18000,${C}l3600,${C}qy0,18000l0,3600qx3600,0x`,
      21,
    ),
    can: shape(
      "basicShapes",
      `m0,2700l0,18900c0,${C + 900},${C},${C + 900},${C},18900l${C},2700c${C},-900,0,-900,0,2700x` +
        `m0,2700c0,6300,${C},6300,${C},2700`,
      22,
    ),
    cube: shape(
      "basicShapes",
      poly([
        [0, 5400],
        [16200, 5400],
        [16200, C],
        [0, C],
      ]) +
        poly([
          [0, 5400],
          [5400, 0],
          [C, 0],
          [16200, 5400],
        ]) +
        poly([
          [16200, 5400],
          [C, 0],
          [C, 16200],
          [16200, C],
        ]),
      16,
    ),
    bevel: shape(
      "basicShapes",
      poly([
        [0, 0],
        [C, 0],
        [C, C],
        [0, C],
      ]) +
        poly([
          [2700, 2700],
          [18900, 2700],
          [18900, 18900],
          [2700, 18900],
        ]) +
        line([
          [0, 0],
          [2700, 2700],
        ]) +
        line([
          [C, 0],
          [18900, 2700],
        ]) +
        line([
          [C, C],
          [18900, 18900],
        ]) +
        line([
          [0, C],
          [2700, 18900],
        ]),
      84,
    ),
    smileyFace: shape(
      "basicShapes",
      ellipse(H, H, H, H) +
        ellipse(7305, 7515, 1165, 1165) +
        ellipse(14295, 7515, 1165, 1165) +
        `nfm4870,15550c8000,18300,13600,18300,16730,15550`,
      96,
    ),
    heart: shape(
      "basicShapes",
      `m${H},${C}c3000,17000,0,13000,0,6750c0,1000,7500,-1500,${H},4350` +
        `c13300,-1500,${C},1000,${C},6750c${C},13000,18600,17000,${H},${C}x`,
      74,
    ),
    lightningBolt: shape(
      "basicShapes",
      poly([
        [8458, 0],
        [0, 3923],
        [7564, 8416],
        [4993, 9720],
        [12197, 13904],
        [9987, 14934],
        [C, C],
        [14768, 12911],
        [16558, 12016],
        [11030, 6840],
        [12831, 6120],
      ]),
      73,
    ),
    sun: shape("basicShapes", sun(), 183),
    moon: shape(
      "basicShapes",
      `m${C},0c8000,0,0,4800,0,${H}c0,16800,8000,${C},${C},${C}` +
        `c14000,19000,${H},15000,${H},${H}c${H},6600,14000,2600,${C},0x`,
      184,
    ),
    bracketPair: shape(
      "basicShapes",
      `nfm3600,0qx0,3600l0,18000qy3600,${C}m18000,${C}qx${C},18000l${C},3600qy18000,0`,
      185,
      { filled: false },
    ),
    bracePair: shape(
      "basicShapes",
      `nfm3600,0qx${BRACE},${BRACE}l${BRACE},9000qy0,${H}qx${BRACE},12600l${BRACE},19800qy3600,${C}` +
        `m18000,${C}qx${C - BRACE},19800l${C - BRACE},12600qy${C},${H}qx${C - BRACE},9000l${C - BRACE},${BRACE}qy18000,0`,
      186,
      { filled: false },
    ),
    leftBracket: shape("basicShapes", `nfm${C},0qx0,${BRACKET}l0,${C - BRACKET}qy${C},${C}`, 85, {
      filled: false,
    }),
    rightBracket: shape("basicShapes", `nfm0,0qx${C},${BRACKET}l${C},${C - BRACKET}qy0,${C}`, 86, {
      filled: false,
    }),
    leftBrace: shape(
      "basicShapes",
      `nfm${C},0qx${H},${BRACE}l${H},9000qy0,${H}qx${H},12600l${H},19800qy${C},${C}`,
      87,
      { filled: false },
    ),
    rightBrace: shape(
      "basicShapes",
      `nfm0,0qx${H},${BRACE}l${H},9000qy${C},${H}qx${H},12600l${H},19800qy0,${C}`,
      88,
      { filled: false },
    ),

    // --- Block arrows --------------------------------------------------------
    rightArrow: shape("blockArrows", poly(RIGHT_ARROW), 13),
    leftArrow: shape("blockArrows", poly(mirrorX(RIGHT_ARROW)), 66),
    upArrow: shape("blockArrows", poly(mirrorY(transpose(RIGHT_ARROW))), 68),
    downArrow: shape("blockArrows", poly(transpose(RIGHT_ARROW)), 67),
    leftRightArrow: shape("blockArrows", poly(LEFT_RIGHT_ARROW), 69),
    upDownArrow: shape("blockArrows", poly(transpose(LEFT_RIGHT_ARROW)), 70),
    quadArrow: shape("blockArrows", poly(quadArrow()), 76),
    bentUpArrow: shape(
      "blockArrows",
      poly([
        [0, 16200],
        [13500, 16200],
        [13500, 5400],
        [10800, 5400],
        [16200, 0],
        [C, 5400],
        [18900, 5400],
        [18900, C],
        [0, C],
      ]),
      90,
    ),
    leftUpArrow: shape(
      "blockArrows",
      poly([
        [0, 16200],
        [5400, 10800],
        [5400, 14850],
        [14850, 14850],
        [14850, 5400],
        [10800, 5400],
        [16200, 0],
        [C, 5400],
        [17550, 5400],
        [17550, 17550],
        [5400, 17550],
        [5400, C],
      ]),
      89,
    ),
    stripedRightArrow: shape(
      "blockArrows",
      poly([
        [0, 5400],
        [1350, 5400],
        [1350, 16200],
        [0, 16200],
      ]) +
        poly([
          [2700, 5400],
          [4050, 5400],
          [4050, 16200],
          [2700, 16200],
        ]) +
        poly([
          [5400, 5400],
          [16200, 5400],
          [16200, 0],
          [C, H],
          [16200, C],
          [16200, 16200],
          [5400, 16200],
        ]),
      93,
    ),
    notchedRightArrow: shape("blockArrows", poly([...RIGHT_ARROW, [2700, H]]), 94),
    homePlate: shape(
      "blockArrows",
      poly([
        [0, 0],
        [16200, 0],
        [C, H],
        [16200, C],
        [0, C],
      ]),
      15,
    ),
    chevron: shape(
      "blockArrows",
      poly([
        [0, 0],
        [16200, 0],
        [C, H],
        [16200, C],
        [0, C],
        [5400, H],
      ]),
      55,
    ),
    rightArrowCallout: shape("blockArrows", poly(RIGHT_ARROW_CALLOUT), 78),
    leftArrowCallout: shape("blockArrows", poly(mirrorX(RIGHT_ARROW_CALLOUT)), 77),
    upArrowCallout: shape("blockArrows", poly(mirrorY(transpose(RIGHT_ARROW_CALLOUT))), 79),
    downArrowCallout: shape("blockArrows", poly(transpose(RIGHT_ARROW_CALLOUT)), 80),

    // --- Equation shapes (DrawingML-only presets, written as custom geometry) ---
    mathPlus: shape(
      "equationShapes",
      poly([
        [8100, 2700],
        [13500, 2700],
        [13500, 8100],
        [18900, 8100],
        [18900, 13500],
        [13500, 13500],
        [13500, 18900],
        [8100, 18900],
        [8100, 13500],
        [2700, 13500],
        [2700, 8100],
        [8100, 8100],
      ]),
    ),
    mathMinus: shape(
      "equationShapes",
      poly([
        [2700, 8100],
        [18900, 8100],
        [18900, 13500],
        [2700, 13500],
      ]),
    ),
    mathMultiply: shape(
      "equationShapes",
      poly([
        [4500, 2700],
        [H, 6900],
        [17100, 2700],
        [18900, 4500],
        [14700, H],
        [18900, 17100],
        [17100, 18900],
        [H, 14700],
        [4500, 18900],
        [2700, 17100],
        [6900, H],
        [2700, 4500],
      ]),
    ),
    mathDivide: shape(
      "equationShapes",
      poly([
        [2700, 9000],
        [18900, 9000],
        [18900, 12600],
        [2700, 12600],
      ]) +
        ellipse(H, 4900, 1900, 1900) +
        ellipse(H, 16700, 1900, 1900),
    ),
    mathEqual: shape(
      "equationShapes",
      poly([
        [2700, 5400],
        [18900, 5400],
        [18900, 9450],
        [2700, 9450],
      ]) +
        poly([
          [2700, 12150],
          [18900, 12150],
          [18900, 16200],
          [2700, 16200],
        ]),
    ),
    mathNotEqual: shape(
      "equationShapes",
      poly([
        [2700, 5400],
        [18900, 5400],
        [18900, 9450],
        [2700, 9450],
      ]) +
        poly([
          [2700, 12150],
          [18900, 12150],
          [18900, 16200],
          [2700, 16200],
        ]) +
        poly([
          [12600, 1800],
          [15300, 1800],
          [9000, 19800],
          [6300, 19800],
        ]),
    ),

    // --- Flowchart -----------------------------------------------------------
    flowChartProcess: shape(
      "flowchart",
      poly([
        [0, 0],
        [C, 0],
        [C, C],
        [0, C],
      ]),
      109,
    ),
    flowChartAlternateProcess: shape("flowchart", roundRect(3600), 176),
    flowChartDecision: shape(
      "flowchart",
      poly([
        [H, 0],
        [C, H],
        [H, C],
        [0, H],
      ]),
      110,
    ),
    flowChartInputOutput: shape(
      "flowchart",
      poly([
        [4320, 0],
        [C, 0],
        [17280, C],
        [0, C],
      ]),
      111,
    ),
    flowChartPredefinedProcess: shape(
      "flowchart",
      poly([
        [0, 0],
        [C, 0],
        [C, C],
        [0, C],
      ]) +
        line([
          [2540, 0],
          [2540, C],
        ]) +
        line([
          [19060, 0],
          [19060, C],
        ]),
      112,
    ),
    flowChartInternalStorage: shape(
      "flowchart",
      poly([
        [0, 0],
        [C, 0],
        [C, C],
        [0, C],
      ]) +
        line([
          [4236, 0],
          [4236, C],
        ]) +
        line([
          [0, 4236],
          [C, 4236],
        ]),
      113,
    ),
    flowChartDocument: shape(
      "flowchart",
      `m0,0l${C},0,${C},17322c${H},17322,${H},23922,0,20172x`,
      114,
    ),
    flowChartMultidocument: shape(
      "flowchart",
      `m0,3600l18000,3600,18000,18000c9000,18000,9000,23400,0,20400x` +
        line([
          [1800, 3600],
          [1800, 1800],
          [19800, 1800],
          [19800, 16200],
          [18000, 16200],
        ]) +
        line([
          [3600, 1800],
          [3600, 0],
          [C, 0],
          [C, 14400],
          [19800, 14400],
        ]),
      115,
    ),
    flowChartTerminator: shape(
      "flowchart",
      `m3475,0qx0,${H},3475,${C}l18125,${C}qx${C},${H},18125,0x`,
      116,
    ),
    flowChartPreparation: shape(
      "flowchart",
      poly([
        [4353, 0],
        [17214, 0],
        [C, H],
        [17214, C],
        [4353, C],
        [0, H],
      ]),
      117,
    ),
    flowChartManualInput: shape(
      "flowchart",
      poly([
        [0, 4292],
        [C, 0],
        [C, C],
        [0, C],
      ]),
      118,
    ),
    flowChartManualOperation: shape(
      "flowchart",
      poly([
        [0, 0],
        [C, 0],
        [17240, C],
        [4360, C],
      ]),
      119,
    ),
    flowChartConnector: { category: "flowchart", element: "oval", spt: 120 },
    flowChartOffpageConnector: shape(
      "flowchart",
      poly([
        [0, 0],
        [C, 0],
        [C, 17255],
        [H, C],
        [0, 17255],
      ]),
      177,
    ),
    flowChartPunchedCard: shape(
      "flowchart",
      poly([
        [4321, 0],
        [C, 0],
        [C, C],
        [0, C],
        [0, 4321],
      ]),
      121,
    ),
    flowChartPunchedTape: shape(
      "flowchart",
      `m0,2230c3600,6200,7200,6200,${H},2230c14400,-1740,18000,-1740,${C},2230l${C},19370` +
        `c18000,15400,14400,15400,${H},19370c7200,23340,3600,23340,0,19370x`,
      122,
    ),
    flowChartSummingJunction: shape(
      "flowchart",
      ellipse(H, H, H, H) +
        line([
          [3163, 3163],
          [18437, 18437],
        ]) +
        line([
          [3163, 18437],
          [18437, 3163],
        ]),
      123,
    ),
    flowChartOr: shape(
      "flowchart",
      ellipse(H, H, H, H) +
        line([
          [H, 0],
          [H, C],
        ]) +
        line([
          [0, H],
          [C, H],
        ]),
      124,
    ),
    flowChartCollate: shape(
      "flowchart",
      poly([
        [0, 0],
        [C, 0],
        [0, C],
        [C, C],
      ]),
      125,
    ),
    flowChartSort: shape(
      "flowchart",
      poly([
        [H, 0],
        [C, H],
        [H, C],
        [0, H],
      ]) +
        line([
          [0, H],
          [C, H],
        ]),
      126,
    ),
    flowChartExtract: shape(
      "flowchart",
      poly([
        [H, 0],
        [C, C],
        [0, C],
      ]),
      127,
    ),
    flowChartMerge: shape(
      "flowchart",
      poly([
        [0, 0],
        [C, 0],
        [H, C],
      ]),
      128,
    ),
    flowChartOnlineStorage: shape(
      "flowchart",
      `m3600,0l${C},0qx18000,${H},${C},${C}l3600,${C}qx0,${H},3600,0x`,
      130,
    ),
    flowChartDelay: shape("flowchart", `m0,0l${H},0qx${C},${H},${H},${C}l0,${C}x`, 135),
    flowChartMagneticTape: shape(
      "flowchart",
      `m${H},${C}c4835,${C},0,16765,0,${H}c0,4835,4835,0,${H},0c16765,0,${C},4835,${C},${H}` +
        `c${C},14000,19800,16900,17500,18900l${C},18900,${C},${C}x`,
      131,
    ),
    flowChartMagneticDisk: shape(
      "flowchart",
      `m0,3400l0,18200c0,22700,${C},22700,${C},18200l${C},3400c${C},-1100,0,-1100,0,3400x` +
        `m0,3400c0,7900,${C},7900,${C},3400`,
      132,
    ),
    flowChartMagneticDrum: shape(
      "flowchart",
      `m3400,0l18200,0c22700,0,22700,${C},18200,${C}l3400,${C}c-1100,${C},-1100,0,3400,0x` +
        `m18200,0c13700,0,13700,${C},18200,${C}`,
      133,
    ),
    flowChartDisplay: shape(
      "flowchart",
      `m3600,0l17800,0qx${C},${H},17800,${C}l3600,${C},0,${H}x`,
      134,
    ),

    // --- Stars and banners ---------------------------------------------------
    irregularSeal1: shape(
      "starsAndBanners",
      poly([
        [10901, 5905],
        [8458, 2399],
        [7417, 6425],
        [476, 2399],
        [4732, 7722],
        [106, 8718],
        [3828, 11880],
        [243, 14689],
        [5772, 14041],
        [4868, 17719],
        [7819, 15730],
        [8590, C],
        [10637, 15038],
        [13349, 19840],
        [14125, 14561],
        [18248, 18195],
        [16938, 13044],
        [C, 13393],
        [17710, 10579],
        [21198, 8242],
        [16806, 7417],
        [18239, 4560],
        [14068, 5651],
        [13801, 1220],
      ]),
      71,
    ),
    irregularSeal2: shape(
      "starsAndBanners",
      poly([
        [11464, 4340],
        [9722, 1887],
        [8548, 6383],
        [4503, 3626],
        [5373, 7816],
        [1174, 8270],
        [3934, 11592],
        [0, 12875],
        [3329, 15372],
        [1283, 17824],
        [4804, 18239],
        [4918, C],
        [7525, 18125],
        [8698, 19712],
        [9871, 17371],
        [11614, 18844],
        [12178, 15937],
        [14943, 17371],
        [14640, 14348],
        [18878, 15632],
        [16382, 12311],
        [18270, 11292],
        [16986, 9404],
        [21600, 6646],
        [16382, 6533],
        [18005, 3172],
        [14524, 5778],
        [14789, 0],
      ]),
      72,
    ),
    star4: shape("starsAndBanners", star(4, 0.25), 187),
    star5: shape("starsAndBanners", star(5, 0.382), 12),
    star8: shape("starsAndBanners", star(8, 0.765), 58),
    star16: shape("starsAndBanners", star(16, 0.75), 59),
    star24: shape("starsAndBanners", star(24, 0.75), 92),
    star32: shape("starsAndBanners", star(32, 0.75), 60),
    ribbon: shape(
      "starsAndBanners",
      poly([
        [0, 2700],
        [5400, 2700],
        [5400, 0],
        [16200, 0],
        [16200, 2700],
        [C, 2700],
        [18900, 10800],
        [C, 18900],
        [16200, 18900],
        [16200, C],
        [5400, C],
        [5400, 18900],
        [0, 18900],
        [2700, 10800],
      ]),
    ),
    wave: shape(
      "starsAndBanners",
      `m0,2700c7200,-2700,14400,8100,${C},2700l${C},18900c14400,24300,7200,13500,0,18900x`,
      64,
    ),
    doubleWave: shape(
      "starsAndBanners",
      `m0,1800c3600,-1800,7200,5400,${H},1800c14400,-1800,18000,5400,${C},1800l${C},19800` +
        `c18000,23400,14400,16200,${H},19800c7200,23400,3600,16200,0,19800x`,
      188,
    ),
    horizontalScroll: shape(
      "starsAndBanners",
      `m0,4050qy1350,2700l18900,2700,18900,1350qx20250,0qy${C},1350l${C},17550qx20250,18900l2700,18900,2700,20250qx1350,${C}qy0,20250x` +
        `m2700,2700l2700,18900`,
      98,
    ),
    verticalScroll: shape(
      "starsAndBanners",
      `m4050,0qx2700,1350l2700,18900,1350,18900qy0,20250qx1350,${C}l17550,${C}qy18900,20250l18900,2700,20250,2700qx${C},1350qy20250,0x` +
        `m2700,2700l18900,2700`,
      97,
    ),

    // --- Callouts ------------------------------------------------------------
    wedgeRectCallout: shape(
      "callouts",
      poly([
        [0, 0],
        [C, 0],
        [C, C],
        [8400, C],
        [1350, 25920],
        [3600, C],
        [0, C],
      ]),
      61,
    ),
    wedgeRoundRectCallout: shape(
      "callouts",
      `m3600,0l18000,0qx${C},3600l${C},18000qy18000,${C}l8400,${C},1350,25920,3600,${C}qx0,18000l0,3600qy3600,0x`,
      62,
    ),
    wedgeEllipseCallout: shape("callouts", wedgeEllipse(), 63),
    cloudCallout: shape(
      "callouts",
      cloud() + ellipse(4000, 20500, 1300, 1100) + ellipse(1800, 23400, 800, 700),
      106,
    ),
    borderCallout1: shape(
      "callouts",
      poly([
        [0, 0],
        [C, 0],
        [C, C],
        [0, C],
      ]) + `nfm0,${H}l-8280,24300`,
      47,
    ),
    accentCallout1: shape(
      "callouts",
      poly([
        [0, 0],
        [C, 0],
        [C, C],
        [0, C],
      ]) + `nfm-1800,0l-1800,${C}m-1800,${H}l-8280,24300`,
      44,
    ),
    callout1: shape(
      "callouts",
      poly([
        [0, 0],
        [C, 0],
        [C, C],
        [0, C],
      ]) + `nfm0,${H}l-8280,24300`,
      41,
      {},
    ),
  } as const satisfies Readonly<Record<string, ShapePresetDef>>;
}

// Built in a function so bundlers can drop the whole table (and the geometry
// helpers it calls) from bundles that never draw a shape.
export const SHAPE_PRESETS = /* @__PURE__ */ buildPresets();

export type ShapePreset = keyof typeof SHAPE_PRESETS;

/** Look up a preset by name (for names that arrive as plain strings). */
export function isShapePreset(name: string): name is ShapePreset {
  return Object.hasOwn(SHAPE_PRESETS, name);
}
