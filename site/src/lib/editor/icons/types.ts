/**
 * One stroke or fill of a ribbon icon on the 20×20 grid. Word's ribbon icons
 * are mostly grey outlines with a colored accent (a blue header band, a red
 * eraser…), so a part may carry its own stroke / fill color.
 */
export interface IconPart {
  readonly d: string;
  /** Stroke color; `currentColor` (the ribbon's icon grey) when omitted. */
  readonly stroke?: string;
  /** Fill color; unfilled when omitted. */
  readonly fill?: string;
}

/** A plain path string is a single grey outline. */
export type IconDef = string | readonly IconPart[];

export type IconSet = Readonly<Record<string, IconDef>>;
