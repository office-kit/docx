/**
 * Ribbon icons, one module per feature area. Line icons are drawn on a 20×20
 * grid in the style of Word's ribbon (thin outlines, a colored accent only where
 * Word has one), as SVG rather than font glyphs so every platform shows the
 * same shape.
 */

import core from "./core.js";
import design from "./design.js";
import draw from "./draw.js";
import home from "./home.js";
import insert from "./insert.js";
import layout from "./layout.js";
import mailings from "./mailings.js";
import picture from "./picture.js";
import references from "./references.js";
import review from "./review.js";
import table from "./table.js";
import type { IconDef, IconPart } from "./types.js";
import view from "./view.js";

export type { IconDef, IconPart };

export const ICONS = {
  ...core,
  ...home,
  ...insert,
  ...draw,
  ...design,
  ...layout,
  ...references,
  ...mailings,
  ...review,
  ...view,
  ...table,
  ...picture,
};

export type IconName = keyof typeof ICONS;

/** The icon's parts, with a plain path string read as one grey outline. */
export function iconParts(name: IconName): readonly IconPart[] {
  const def: IconDef = ICONS[name];
  return typeof def === "string" ? [{ d: def }] : def;
}
