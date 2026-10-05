/**
 * Previews for the Styles gallery and the Styles pane: each style's own
 * formatting as inline CSS. Word's built-in styles the document does not
 * define yet are previewed from Word's definitions in a scratch document, so
 * Title shows big before it is ever applied.
 */

import { builtinStyles, createDocx, ensureBuiltinStyle, type Docx } from "@office-kit/docx";
import {
  commands,
  createStyleResolver,
  type EditorModel,
  type ResolvedRunFormat,
  type StyleResolver,
} from "@office-kit/docx-editor";

// The gallery squeezes sizes into its tile: Word's previews top out near 20 pt.
const MAX_PREVIEW_PT = 20;
const MIN_PREVIEW_PT = 9;
const DEFAULT_PT = 11;

let builtinResolver: StyleResolver | undefined;

function builtins(): StyleResolver {
  if (!builtinResolver) {
    const scratch: Docx = createDocx({ paragraphs: [] });
    for (const s of builtinStyles()) ensureBuiltinStyle(scratch, s.styleId);
    builtinResolver = createStyleResolver(scratch);
  }
  return builtinResolver;
}

function css(fmt: ResolvedRunFormat): string {
  const pt = Math.min(
    MAX_PREVIEW_PT,
    Math.max(MIN_PREVIEW_PT, (fmt.sizeHalfPoints ?? DEFAULT_PT * 2) / 2),
  );
  const parts = [`font-size:${pt}px`];
  if (fmt.font) parts.push(`font-family:"${fmt.font.replaceAll('"', "")}",sans-serif`);
  if (fmt.bold) parts.push("font-weight:700");
  if (fmt.italic) parts.push("font-style:italic");
  if (fmt.color && fmt.color !== "auto") parts.push(`color:#${fmt.color}`);
  if (fmt.toggles.has("caps")) parts.push("text-transform:uppercase");
  if (fmt.toggles.has("smallCaps")) parts.push("font-variant-caps:small-caps");
  if (fmt.underline && fmt.underline !== "none") parts.push("text-decoration:underline");
  return parts.join(";");
}

export interface StylePreviews {
  css(entry: commands.StyleEntry): string;
}

/** A preview lookup for the current document. */
export function stylePreviews(model: EditorModel): StylePreviews {
  const own = createStyleResolver(model.doc);
  return {
    css: (entry) => css((entry.inDocument ? own : builtins()).style(entry.styleId).run),
  };
}

/** The styles of the gallery: the quick styles, in Word's order. */
export function galleryStyles(model: EditorModel): commands.StyleEntry[] {
  return commands.listStyles(model).filter((s) => s.quick);
}
