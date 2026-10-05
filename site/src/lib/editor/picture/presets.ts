/**
 * Word's Picture Format galleries, expressed with ECMA-376 properties only.
 *
 * Picture Styles are Word's presets rebuilt from `a:ln`, `a:effectLst`,
 * `a:scene3d` / `a:sp3d` and `a:prstGeom`; the styles that need anything
 * else (gradient frames, metal textures) are left out. Corrections, Color and
 * Transparency presets are blip effects; Word's Sharpen/Soften, Color Tone and
 * Artistic Effects are `a14:` extensions and have no ECMA-376 form.
 */

import type {
  BevelPreset,
  LineDash,
  PictureColorAdjustments,
  PictureEffects,
  PictureOutline,
  PictureShadow,
} from "@office-kit/docx";

const PT = 12700;

export interface PictureStyle {
  readonly id: string;
  readonly name: string;
  readonly outline: PictureOutline | undefined;
  readonly effects: PictureEffects;
  readonly geometry: string;
}

const frame = (
  color: string,
  pt: number,
  compound: PictureOutline["compound"] = "sng",
): PictureOutline => ({
  color,
  widthEmu: pt * PT,
  dash: "solid",
  compound,
});

const dropShadow: PictureShadow = {
  kind: "outer",
  color: "000000",
  opacity: 40,
  blurEmu: 6 * PT,
  distanceEmu: 3 * PT,
  directionDeg: 45,
  align: "tl",
};

/** Picture Styles gallery, in Word's order where the style is expressible. */
export const PICTURE_STYLES: readonly PictureStyle[] = [
  {
    id: "simpleWhite",
    name: "Simple Frame, White",
    outline: frame("FFFFFF", 7),
    effects: {},
    geometry: "rect",
  },
  {
    id: "bevelMatteWhite",
    name: "Beveled Matte, White",
    outline: frame("FFFFFF", 7),
    effects: {
      shadow: { ...dropShadow, blurEmu: 4 * PT, opacity: 35 },
      bevel: { preset: "softRound", widthEmu: 4 * PT, heightEmu: 4 * PT },
    },
    geometry: "rect",
  },
  {
    id: "dropShadow",
    name: "Drop Shadow Rectangle",
    outline: undefined,
    effects: { shadow: dropShadow },
    geometry: "rect",
  },
  {
    id: "roundedDiagonalWhite",
    name: "Rounded Diagonal Corner, White",
    outline: frame("FFFFFF", 7),
    effects: { shadow: { ...dropShadow, opacity: 30 } },
    geometry: "round2DiagRect",
  },
  {
    id: "snipDiagonalWhite",
    name: "Snip Diagonal Corner, White",
    outline: frame("FFFFFF", 7),
    effects: { shadow: { ...dropShadow, opacity: 30 } },
    geometry: "snip2DiagRect",
  },
  {
    id: "moderateFrameWhite",
    name: "Moderate Frame, White",
    outline: frame("FFFFFF", 5),
    effects: { shadow: { ...dropShadow, blurEmu: 4 * PT, distanceEmu: 2 * PT, opacity: 35 } },
    geometry: "rect",
  },
  {
    id: "centerShadow",
    name: "Center Shadow Rectangle",
    outline: undefined,
    effects: {
      shadow: {
        kind: "outer",
        color: "000000",
        opacity: 40,
        blurEmu: 20 * PT,
        distanceEmu: 0,
        directionDeg: 0,
        align: "ctr",
      },
    },
    geometry: "rect",
  },
  {
    id: "softEdgeRect",
    name: "Soft Edge Rectangle",
    outline: undefined,
    effects: { softEdgeEmu: 9 * PT },
    geometry: "rect",
  },
  {
    id: "doubleFrameBlack",
    name: "Double Frame, Black",
    outline: frame("000000", 6, "dbl"),
    effects: {},
    geometry: "rect",
  },
  {
    id: "thickMatteBlack",
    name: "Thick Matte, Black",
    outline: frame("000000", 15),
    effects: {},
    geometry: "rect",
  },
  {
    id: "simpleBlack",
    name: "Simple Frame, Black",
    outline: frame("000000", 7),
    effects: {},
    geometry: "rect",
  },
  {
    id: "reflectedRounded",
    name: "Reflected Rounded Rectangle",
    outline: undefined,
    effects: {
      reflection: {
        blurEmu: 0.5 * PT,
        startOpacity: 50,
        endOpacity: 0.3,
        endPosition: 35,
        distanceEmu: 0,
      },
    },
    geometry: "roundRect",
  },
  {
    id: "softEdgeOval",
    name: "Soft Edge Oval",
    outline: undefined,
    effects: { softEdgeEmu: 9 * PT },
    geometry: "ellipse",
  },
  {
    id: "compoundBlack",
    name: "Compound Frame, Black",
    outline: frame("000000", 6, "thickThin"),
    effects: {},
    geometry: "rect",
  },
  {
    id: "moderateFrameBlack",
    name: "Moderate Frame, Black",
    outline: frame("000000", 5),
    effects: { shadow: { ...dropShadow, opacity: 35 } },
    geometry: "rect",
  },
  {
    id: "beveledOvalBlack",
    name: "Beveled Oval, Black",
    outline: frame("000000", 5),
    effects: { bevel: { preset: "circle", widthEmu: 4 * PT, heightEmu: 4 * PT } },
    geometry: "ellipse",
  },
  {
    id: "bevelRectangle",
    name: "Bevel Rectangle",
    outline: undefined,
    effects: { bevel: { preset: "circle", widthEmu: 10 * PT, heightEmu: 4 * PT } },
    geometry: "rect",
  },
  {
    id: "relaxedPerspectiveWhite",
    name: "Relaxed Perspective, White",
    outline: frame("FFFFFF", 7),
    effects: {
      shadow: { ...dropShadow, opacity: 30 },
      rotation3d: { camera: "perspectiveRelaxed" },
    },
    geometry: "rect",
  },
  {
    id: "reflectedBevelBlack",
    name: "Reflected Bevel, Black",
    outline: frame("000000", 7),
    effects: {
      bevel: { preset: "angle", widthEmu: 4 * PT, heightEmu: 4 * PT },
      reflection: {
        blurEmu: 0.5 * PT,
        startOpacity: 40,
        endOpacity: 0.3,
        endPosition: 30,
        distanceEmu: 0,
      },
    },
    geometry: "rect",
  },
  {
    id: "perspectiveShadowWhite",
    name: "Perspective Shadow, White",
    outline: frame("FFFFFF", 7),
    effects: {
      shadow: {
        ...dropShadow,
        scaleY: 50,
        skewX: -30,
        align: "bl",
        directionDeg: 270,
        distanceEmu: 0,
      },
    },
    geometry: "rect",
  },
];

/** Corrections ▸ Brightness / Contrast: Word's 5 × 5 grid (contrast rows, brightness columns). */
export const CORRECTION_STEPS = [-40, -20, 0, 20, 40] as const;

/** Color ▸ Color Saturation (`a:hsl sat`). */
export const SATURATIONS = [0, 33, 66, 100, 133, 166, 200, 300, 400] as const;

export interface Recolor {
  readonly id: string;
  readonly name: string;
  readonly adjustments: PictureColorAdjustments;
}

// The Office theme's dark/light variants of each accent, for Recolor's accent rows.
const ACCENTS = [
  ["4472C4", "203864", "DAE3F3"],
  ["ED7D31", "843C0C", "FBE5D6"],
  ["A5A5A5", "525252", "EDEDED"],
  ["FFC000", "7F6000", "FFF2CC"],
  ["5B9BD5", "1F4E79", "DEEBF7"],
  ["70AD47", "385723", "E2F0D9"],
] as const;

/** Color ▸ Recolor. */
export const RECOLORS: readonly Recolor[] = [
  { id: "none", name: "No Recolor", adjustments: {} },
  { id: "grayscale", name: "Grayscale", adjustments: { grayscale: true } },
  { id: "sepia", name: "Sepia", adjustments: { duotone: ["000000", "D9C3A5"] } },
  { id: "washout", name: "Washout", adjustments: { brightness: 70, contrast: -70 } },
  { id: "bw25", name: "Black and White: 25%", adjustments: { biLevelThreshold: 25 } },
  { id: "bw50", name: "Black and White: 50%", adjustments: { biLevelThreshold: 50 } },
  { id: "bw75", name: "Black and White: 75%", adjustments: { biLevelThreshold: 75 } },
  ...ACCENTS.map(([accent, dark], i) => ({
    id: `dark${i + 1}`,
    name: `Accent ${i + 1} Dark`,
    adjustments: { duotone: [dark, accent] as const },
  })),
  ...ACCENTS.map(([accent, , light], i) => ({
    id: `light${i + 1}`,
    name: `Accent ${i + 1} Light`,
    adjustments: { duotone: [accent, light] as const },
  })),
];

/** Transparency presets, percent. */
export const TRANSPARENCIES = [0, 15, 30, 50, 65, 80, 95] as const;

/** Picture Border ▸ Weight, points. */
export const BORDER_WEIGHTS = [0.25, 0.5, 0.75, 1, 1.5, 2.25, 3, 4.5, 6] as const;

/** Picture Border ▸ Dashes. */
export const BORDER_DASHES: readonly LineDash[] = [
  "solid",
  "sysDot",
  "sysDash",
  "dash",
  "dashDot",
  "lgDash",
  "lgDashDot",
  "lgDashDotDot",
];

// Word's shadow gallery directions: the shadow falls away from the light.
const DIRECTIONS: ReadonlyArray<readonly [string, number]> = [
  ["Offset: Bottom Right", 45],
  ["Offset: Bottom", 90],
  ["Offset: Bottom Left", 135],
  ["Offset: Right", 0],
  ["Offset: Center", 0],
  ["Offset: Left", 180],
  ["Offset: Top Right", 315],
  ["Offset: Top", 270],
  ["Offset: Top Left", 225],
];

/** Picture Effects ▸ Shadow. */
export const SHADOWS: ReadonlyArray<{ name: string; shadow: PictureShadow }> = [
  ...DIRECTIONS.map(([name, dir]) => ({
    name,
    shadow: {
      kind: "outer" as const,
      color: "000000",
      opacity: 40,
      blurEmu: 4 * PT,
      distanceEmu: name === "Offset: Center" ? 0 : 3 * PT,
      directionDeg: dir,
      align: "ctr" as const,
    },
  })),
  ...DIRECTIONS.map(([name, dir]) => ({
    name: name.replace("Offset", "Inside"),
    shadow: {
      kind: "inner" as const,
      color: "000000",
      opacity: 50,
      blurEmu: 5 * PT,
      distanceEmu: name === "Offset: Center" ? 0 : 3 * PT,
      directionDeg: dir,
    },
  })),
];

/** Picture Effects ▸ Reflection: tight / half / full × touching / 4 pt / 8 pt offset. */
const reflection = (name: string, endPosition: number, distanceEmu: number) => ({
  name,
  endPosition,
  distanceEmu,
  startOpacity: 50,
  endOpacity: 0.3,
  blurEmu: 0.5 * PT,
});

export const REFLECTIONS = [
  reflection("Tight Reflection: Touching", 35, 0),
  reflection("Half Reflection: Touching", 55, 0),
  reflection("Full Reflection: Touching", 90, 0),
  reflection("Tight Reflection: 4 pt offset", 35, 4 * PT),
  reflection("Half Reflection: 4 pt offset", 55, 4 * PT),
  reflection("Full Reflection: 4 pt offset", 90, 4 * PT),
  reflection("Tight Reflection: 8 pt offset", 35, 8 * PT),
  reflection("Half Reflection: 8 pt offset", 55, 8 * PT),
  reflection("Full Reflection: 8 pt offset", 90, 8 * PT),
];

/** Picture Effects ▸ Glow: 4 sizes × the six accents, at 40 % transparency. */
export const GLOW_SIZES = [5, 8, 11, 18] as const;
export const GLOW_COLORS = ACCENTS.map(([accent]) => accent);
export const GLOW_OPACITY = 60;

/** Picture Effects ▸ Soft Edges, points. */
export const SOFT_EDGES = [1, 2.5, 5, 10, 25, 50] as const;

/** Picture Effects ▸ Bevel. */
export const BEVELS: readonly BevelPreset[] = [
  "circle",
  "relaxedInset",
  "cross",
  "coolSlant",
  "angle",
  "softRound",
  "convex",
  "slope",
  "divot",
  "riblet",
  "hardEdge",
  "artDeco",
];

/** Picture Effects ▸ 3-D Rotation (ST_PresetCameraType). */
export const ROTATIONS_3D = [
  "isometricOffAxis1Left",
  "isometricOffAxis1Right",
  "isometricOffAxis2Left",
  "isometricOffAxis2Right",
  "isometricOffAxis1Top",
  "isometricOffAxis2Top",
  "perspectiveFront",
  "perspectiveLeft",
  "perspectiveRight",
  "perspectiveBelow",
  "perspectiveAbove",
  "perspectiveRelaxed",
  "perspectiveRelaxedModerately",
  "obliqueTopLeft",
  "obliqueTopRight",
  "obliqueBottomLeft",
  "obliqueBottomRight",
] as const;

/** Crop ▸ Aspect Ratio. */
export const ASPECT_RATIOS = {
  square: [[1, 1]],
  portrait: [
    [2, 3],
    [3, 4],
    [3, 5],
    [4, 5],
  ],
  landscape: [
    [3, 2],
    [4, 3],
    [5, 3],
    [5, 4],
    [16, 9],
    [16, 10],
  ],
} as const;

/** Chart Design ▸ Change Colors: Colorful palettes and Monochromatic ramps. */
export const CHART_PALETTES: ReadonlyArray<{ id: string; colors: readonly string[] }> = [
  { id: "colorful1", colors: ["4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47"] },
  { id: "colorful2", colors: ["4472C4", "A5A5A5", "5B9BD5", "264478", "636363", "255E91"] },
  { id: "colorful3", colors: ["ED7D31", "FFC000", "70AD47", "9E480E", "997300", "43682B"] },
  { id: "colorful4", colors: ["70AD47", "4472C4", "FFC000", "43682B", "264478", "997300"] },
  { id: "mono1", colors: ["4472C4", "8FAADC", "2F5597", "B4C7E7", "203864", "DAE3F3"] },
  { id: "mono2", colors: ["ED7D31", "F4B183", "C55A11", "F8CBAD", "843C0C", "FBE5D6"] },
  { id: "mono3", colors: ["A5A5A5", "C9C9C9", "7B7B7B", "DBDBDB", "525252", "EDEDED"] },
  { id: "mono4", colors: ["70AD47", "A9D18E", "548235", "C5E0B4", "385723", "E2F0D9"] },
];
