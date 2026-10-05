/**
 * Plain-data views of a `<w:drawing>` (ECMA-376 Part 1, §20.4 DrawingML -
 * WordprocessingML Drawing, §20.1 DrawingML main, §19.3 pictures). Lengths are
 * EMU (914400 per inch), angles degrees, ratios percent.
 */

/** Word's Wrap Text choices; `behindText` / `inFrontOfText` are `wp:wrapNone` with `behindDoc`. */
export type WrapStyle =
  | "inline"
  | "square"
  | "tight"
  | "through"
  | "topAndBottom"
  | "behindText"
  | "inFrontOfText";

/** ST_WrapText: which sides of a wrapped object text may flow on. */
export type WrapSide = "bothSides" | "left" | "right" | "largest";

/** ST_RelFromH. */
export type HorizontalRelativeTo =
  | "character"
  | "column"
  | "insideMargin"
  | "leftMargin"
  | "margin"
  | "outsideMargin"
  | "page"
  | "rightMargin";

/** ST_RelFromV. */
export type VerticalRelativeTo =
  | "bottomMargin"
  | "insideMargin"
  | "line"
  | "margin"
  | "outsideMargin"
  | "page"
  | "paragraph"
  | "topMargin";

/** ST_AlignH. */
export type HorizontalAlign = "left" | "center" | "right" | "inside" | "outside";
/** ST_AlignV. */
export type VerticalAlign = "top" | "center" | "bottom" | "inside" | "outside";

/** `wp:positionH`: an alignment or an absolute offset, relative to a frame. */
export type HorizontalPosition =
  | { readonly relativeTo: HorizontalRelativeTo; readonly align: HorizontalAlign }
  | { readonly relativeTo: HorizontalRelativeTo; readonly offsetEmu: number };

/** `wp:positionV`. */
export type VerticalPosition =
  | { readonly relativeTo: VerticalRelativeTo; readonly align: VerticalAlign }
  | { readonly relativeTo: VerticalRelativeTo; readonly offsetEmu: number };

/** A point of a `wp:wrapPolygon`, in the 21600-unit square Word maps onto the object. */
export interface WrapPoint {
  readonly x: number;
  readonly y: number;
}

/** Distance from text (`distT` / `distB` / `distL` / `distR`). */
export interface WrapDistance {
  readonly top: number;
  readonly bottom: number;
  readonly left: number;
  readonly right: number;
}

/** A floating (`wp:anchor`) object's placement and wrapping. */
export interface DrawingAnchor {
  readonly horizontal: HorizontalPosition;
  readonly vertical: VerticalPosition;
  readonly wrapSide: WrapSide;
  readonly wrapPolygon?: readonly WrapPoint[];
  readonly distance: WrapDistance;
  readonly allowOverlap: boolean;
  /** Lock anchor: the anchor paragraph stays put while the object moves. */
  readonly locked: boolean;
  readonly layoutInCell: boolean;
  /** Z-order among floating objects; higher draws on top. */
  readonly relativeHeight: number;
}

/** `a:srcRect`: percent trimmed from each edge (negative pads, as Crop ▸ Fit does). */
export interface PictureCrop {
  readonly left: number;
  readonly top: number;
  readonly right: number;
  readonly bottom: number;
}

/** ST_PresetLineDashVal. */
export type LineDash =
  | "solid"
  | "dot"
  | "dash"
  | "lgDash"
  | "dashDot"
  | "lgDashDot"
  | "lgDashDotDot"
  | "sysDash"
  | "sysDot"
  | "sysDashDot"
  | "sysDashDotDot";

/** ST_CompoundLine. */
export type CompoundLine = "sng" | "dbl" | "thickThin" | "thinThick" | "tri";

/** `a:ln` with a solid fill: Picture Border. */
export interface PictureOutline {
  readonly color: string;
  readonly widthEmu: number;
  readonly dash: LineDash;
  readonly compound: CompoundLine;
}

/** ST_RectAlignment. */
export type RectAlignment = "tl" | "t" | "tr" | "l" | "ctr" | "r" | "bl" | "b" | "br";

/** `a:outerShdw` / `a:innerShdw`. */
export interface PictureShadow {
  readonly kind: "outer" | "inner";
  readonly color: string;
  /** Opacity in percent (100 = opaque). */
  readonly opacity: number;
  readonly blurEmu: number;
  readonly distanceEmu: number;
  readonly directionDeg: number;
  /** Outer shadows only: size in percent, skew in degrees, and the anchor of that transform. */
  readonly scaleX?: number;
  readonly scaleY?: number;
  readonly skewX?: number;
  readonly align?: RectAlignment;
}

/** `a:reflection`. */
export interface PictureReflection {
  readonly blurEmu: number;
  readonly startOpacity: number;
  readonly endOpacity: number;
  /** How far down the reflection the fade ends, in percent of its height. */
  readonly endPosition: number;
  readonly distanceEmu: number;
}

/** `a:glow`. */
export interface PictureGlow {
  readonly radiusEmu: number;
  readonly color: string;
  readonly opacity: number;
}

/** ST_BevelPresetType. */
export type BevelPreset =
  | "relaxedInset"
  | "circle"
  | "slope"
  | "cross"
  | "angle"
  | "softRound"
  | "convex"
  | "coolSlant"
  | "divot"
  | "riblet"
  | "hardEdge"
  | "artDeco";

/** `a:sp3d/a:bevelT`. */
export interface PictureBevel {
  readonly preset: BevelPreset;
  readonly widthEmu: number;
  readonly heightEmu: number;
}

/** `a:scene3d/a:camera`: an ST_PresetCameraType and an optional explicit rotation (degrees). */
export interface Picture3dRotation {
  readonly camera: string;
  readonly latitude?: number;
  readonly longitude?: number;
  readonly revolution?: number;
}

/** Picture Effects: the `a:effectLst`, `a:scene3d` and `a:sp3d` of the picture's shape properties. */
export interface PictureEffects {
  readonly shadow?: PictureShadow;
  readonly reflection?: PictureReflection;
  readonly glow?: PictureGlow;
  readonly softEdgeEmu?: number;
  readonly bevel?: PictureBevel;
  readonly rotation3d?: Picture3dRotation;
}

/**
 * Corrections / Color / Transparency: the blip effects of `a:blip`
 * (§20.1.8). Brightness and contrast are percent in [-100, 100]; saturation is
 * percent where 100 leaves the picture unchanged.
 */
export interface PictureColorAdjustments {
  readonly brightness?: number;
  readonly contrast?: number;
  readonly saturation?: number;
  readonly grayscale?: boolean;
  /** Black and White: `a:biLevel` threshold in percent. */
  readonly biLevelThreshold?: number;
  /** Recolor: `a:duotone`, dark then light color. */
  readonly duotone?: readonly [string, string];
  /** Set Transparent Color: `a:clrChange` from this color to transparent. */
  readonly transparentColor?: string;
  /** Picture transparency in percent (`a:alphaModFix`). */
  readonly transparency?: number;
}

export interface PictureImage {
  readonly partName: string;
  readonly contentType: string;
  readonly data: Uint8Array;
}

/** The picture-specific part of a drawing (`pic:pic`). */
export interface PictureInfo {
  readonly image?: PictureImage;
  readonly crop: PictureCrop;
  /** ST_ShapeType of `a:prstGeom` (Crop to Shape). */
  readonly geometry: string;
  readonly outline?: PictureOutline;
  readonly effects: PictureEffects;
  readonly adjustments: PictureColorAdjustments;
}

export type DrawingKind = "picture" | "chart" | "other";

/** Everything the editor reads from one `<w:drawing>`. */
export interface DrawingInfo {
  readonly kind: DrawingKind;
  /** `wp:docPr` id / name / descr / title / hidden. */
  readonly id: number;
  readonly name: string;
  readonly description: string;
  readonly title: string;
  readonly hidden: boolean;
  readonly widthEmu: number;
  readonly heightEmu: number;
  readonly lockAspect: boolean;
  /** `a:hlinkClick` target URL. */
  readonly hyperlink?: string;
  readonly wrap: WrapStyle;
  readonly anchor?: DrawingAnchor;
  readonly rotation: number;
  readonly flipH: boolean;
  readonly flipV: boolean;
  readonly picture?: PictureInfo;
  /** For charts: the chart part the graphic frame references. */
  readonly chartPartName?: string;
}
