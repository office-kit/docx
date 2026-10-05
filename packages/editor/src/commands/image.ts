/**
 * Picture commands: insert at the caret, and everything on Picture Format
 * that changes the picture itself (Adjust, Picture Styles, Crop, Size, Alt
 * Text). Pictures are addressed by `index` in `imageDrawings` order; sizes
 * are EMU (914400 per inch).
 */

import {
  addImageRun,
  type AddImageOptions,
  changePicture,
  imageDrawings,
  imagePixelSize,
  images,
  isolateParagraphRunRange,
  type PictureColorAdjustments,
  type PictureCrop,
  type PictureEffects,
  type PictureOutline,
  readDrawing,
  replaceImage,
  resetPicture,
  setDrawingTransform,
  setImageAltText,
  setImageSizeEmu,
  setPictureColorAdjustments,
  setPictureCrop,
  setPictureEffects,
  setPictureGeometry,
  setPictureOutline,
  type WmlParagraph,
  type WmlRun,
} from "@office-kit/docx";
import { absoluteOffset } from "../char-offset.js";
import { paragraphAt } from "../doc-access.js";
import type { EditorModel } from "../model.js";
import { caretBlockIndex, moveLastBlockAfter } from "./insert-util.js";
import type { Command } from "./types.js";

/**
 * Put `run` at the caret: the caret's run is split there and the new run goes
 * between the halves, then the caret moves just past it. Without a caret
 * paragraph (a table or raw block is selected) the run gets a paragraph of its
 * own after the caret block.
 */
export function insertRunAtCaret(model: EditorModel, run: WmlRun): void {
  const pos = model.selection?.focus;
  const para = pos ? paragraphAt(model.doc, pos) : undefined;
  if (!pos || !para) {
    const at = caretBlockIndex(model.doc, pos?.block);
    const paragraph: WmlParagraph = { kind: "paragraph", children: [run], extras: [] };
    model.doc.document.body.blocks.push(paragraph);
    moveLastBlockAfter(model.doc, at);
    model.doc.dirty = true;
    return;
  }
  const before = new Set<WmlRun>(isolateParagraphRunRange(para, 0, absoluteOffset(para, pos)));
  let index = 0;
  for (const [i, child] of para.children.entries()) {
    if (child.kind === "run" && before.has(child)) index = i + 1;
  }
  para.children.splice(index, 0, run);
  // The caret goes to the start of whatever follows; an empty text run after
  // the object gives it a home when nothing does.
  const runs = para.children.filter((c): c is WmlRun => c.kind === "run");
  let after = runs.indexOf(run) + 1;
  if (!runs[after]) {
    para.children.splice(index + 1, 0, { kind: "run", pieces: [], extras: [] });
    after = runs.indexOf(run) + 1;
  }
  model.setSelection({
    anchor: { ...pos, inline: after, offset: 0 },
    focus: { ...pos, inline: after, offset: 0 },
  });
  model.doc.dirty = true;
}

export const insertImageCommand: Command<{ bytes: Uint8Array; options: AddImageOptions }> = {
  id: "image.insert",
  group: "image",
  label: "Insert picture",
  run(model, { bytes, options }) {
    insertRunAtCaret(model, addImageRun(model.doc, bytes, options));
  },
};

export const replaceImageCommand: Command<{ index: number; bytes: Uint8Array }> = {
  id: "image.replace",
  group: "image",
  label: "Replace picture",
  run(model, { index, bytes }) {
    const refs = images(model.doc);
    const ref = refs[index];
    if (!ref) return;
    replaceImage(model.doc, ref.partName, bytes);
  },
  isEnabled: (model) => images(model.doc).length > 0,
};

const hasDrawings = (model: EditorModel): boolean => imageDrawings(model.doc).length > 0;

/** Resize a picture or chart (EMU: 914400 per inch). */
export const resizeImageCommand: Command<{ index: number; cxEmu: number; cyEmu: number }> = {
  id: "image.resize",
  group: "image",
  label: "Resize picture",
  run(model, { index, cxEmu, cyEmu }) {
    setImageSizeEmu(model.doc, index, cxEmu, cyEmu);
  },
  isEnabled: hasDrawings,
};

/** Set alt text (description / title) on an existing image. */
export const setImageAltCommand: Command<{ index: number; descr?: string; title?: string }> = {
  id: "image.altText",
  group: "image",
  label: "Picture alt text",
  run(model, { index, descr, title }) {
    setImageAltText(model.doc, index, {
      ...(descr !== undefined ? { descr } : {}),
      ...(title !== undefined ? { title } : {}),
    });
  },
  isEnabled: hasDrawings,
};

/** Change Picture: new image bytes, same size and formatting; the crop is cleared as Word does. */
export const changePictureCommand: Command<{ index: number; bytes: Uint8Array }> = {
  id: "image.change",
  group: "image",
  label: "Change Picture",
  run(model, { index, bytes }) {
    changePicture(model.doc, index, bytes);
    setPictureCrop(model.doc, index, { left: 0, top: 0, right: 0, bottom: 0 });
  },
  isEnabled: hasDrawings,
};

/** Compress Pictures: re-encoded bytes, optionally with the cropped areas already cut away. */
export const compressPictureCommand: Command<{
  index: number;
  bytes: Uint8Array;
  croppedAway: boolean;
}> = {
  id: "image.compress",
  group: "image",
  label: "Compress Pictures",
  run(model, { index, bytes, croppedAway }) {
    changePicture(model.doc, index, bytes);
    if (croppedAway) setPictureCrop(model.doc, index, { left: 0, top: 0, right: 0, bottom: 0 });
  },
  isEnabled: hasDrawings,
};

export const resetPictureCommand: Command<{ index: number; size: boolean }> = {
  id: "image.reset",
  group: "image",
  label: "Reset Picture",
  run(model, { index, size }) {
    resetPicture(model.doc, index, { size });
  },
  isEnabled: hasDrawings,
};

/** Crop; with `size`, the picture's frame changes too (dragging a crop handle shrinks the frame). */
export const cropPictureCommand: Command<{
  index: number;
  crop: PictureCrop;
  size?: { cxEmu: number; cyEmu: number };
}> = {
  id: "image.crop",
  group: "image",
  label: "Crop",
  run(model, { index, crop, size }) {
    setPictureCrop(model.doc, index, crop);
    if (size) setImageSizeEmu(model.doc, index, size.cxEmu, size.cyEmu);
  },
  isEnabled: hasDrawings,
};

/**
 * Crop ▸ Aspect Ratio / Fill / Fit: crop the image to the frame's (or a new)
 * aspect ratio, centered. `ratio` is width / height; `fit` pads instead of
 * cutting (negative crop), so the whole image shows.
 */
export const cropToAspectCommand: Command<{
  index: number;
  mode: "fill" | "fit" | { ratio: number };
}> = {
  id: "image.cropAspect",
  group: "image",
  label: "Aspect Ratio",
  run(model, { index, mode }) {
    const drawing = imageDrawings(model.doc)[index];
    if (!drawing) throw new Error(`No picture at index ${index}.`);
    const info = readDrawing(model.doc, drawing);
    const image = info.picture?.image;
    if (!image) throw new Error("The picture has no image data.");
    const natural = imageAspect(image.data);
    let frameW = info.widthEmu;
    let frameH = info.heightEmu;
    if (typeof mode === "object") {
      if (!(mode.ratio > 0)) throw new Error(`Invalid aspect ratio ${mode.ratio}.`);
      // Keep the area roughly the same while matching the new ratio.
      const area = frameW * frameH;
      frameW = Math.sqrt(area * mode.ratio);
      frameH = frameW / mode.ratio;
    }
    const frame = frameW / frameH;
    const fit = mode === "fit";
    // Fraction of the image to trim (positive) or pad (negative) per side.
    let cropX = 0;
    let cropY = 0;
    if (natural > frame !== fit) cropX = (1 - frame / natural) / 2;
    else cropY = (1 - natural / frame) / 2;
    setPictureCrop(model.doc, index, {
      left: cropPercent(cropX),
      right: cropPercent(cropX),
      top: cropPercent(cropY),
      bottom: cropPercent(cropY),
    });
    setImageSizeEmu(model.doc, index, frameW, frameH);
  },
  isEnabled: hasDrawings,
};

/** Width / height of image bytes; square when the format's header is unknown. */
function imageAspect(bytes: Uint8Array): number {
  const size = imagePixelSize(bytes);
  return size && size.height > 0 ? size.width / size.height : 1;
}

/** A crop fraction as the percent setPictureCrop takes (3 decimals, the a:srcRect resolution). */
function cropPercent(f: number): number {
  return Math.round(f * 100 * 1000) / 1000;
}

export const pictureShapeCommand: Command<{ index: number; preset: string }> = {
  id: "image.geometry",
  group: "image",
  label: "Crop to Shape",
  run(model, { index, preset }) {
    setPictureGeometry(model.doc, index, preset);
  },
  isEnabled: hasDrawings,
};

export const pictureBorderCommand: Command<{ index: number; outline: PictureOutline | undefined }> =
  {
    id: "image.outline",
    group: "image",
    label: "Picture Border",
    run(model, { index, outline }) {
      setPictureOutline(model.doc, index, outline);
    },
    isEnabled: hasDrawings,
  };

export const pictureEffectsCommand: Command<{ index: number; effects: PictureEffects }> = {
  id: "image.effects",
  group: "image",
  label: "Picture Effects",
  run(model, { index, effects }) {
    setPictureEffects(model.doc, index, effects);
  },
  isEnabled: hasDrawings,
};

export const pictureAdjustCommand: Command<{
  index: number;
  adjustments: PictureColorAdjustments;
}> = {
  id: "image.adjust",
  group: "image",
  label: "Picture Corrections",
  run(model, { index, adjustments }) {
    setPictureColorAdjustments(model.doc, index, adjustments);
  },
  isEnabled: hasDrawings,
};

/**
 * A Picture Styles gallery entry: border, effects and shape applied together
 * (Word's styles are composed from exactly these properties).
 */
export const pictureStyleCommand: Command<{
  index: number;
  outline: PictureOutline | undefined;
  effects: PictureEffects;
  geometry: string;
}> = {
  id: "image.style",
  group: "image",
  label: "Picture Styles",
  run(model, { index, outline, effects, geometry }) {
    setPictureOutline(model.doc, index, outline);
    setPictureEffects(model.doc, index, effects);
    setPictureGeometry(model.doc, index, geometry);
  },
  isEnabled: hasDrawings,
};

/** Rotate (absolute degrees) and flip. */
export const pictureTransformCommand: Command<{
  index: number;
  rotation?: number;
  flipH?: boolean;
  flipV?: boolean;
}> = {
  id: "image.transform",
  group: "image",
  label: "Rotate",
  run(model, { index, ...t }) {
    setDrawingTransform(model.doc, index, t);
  },
  isEnabled: hasDrawings,
};

export const imageCommands = [
  insertImageCommand,
  replaceImageCommand,
  resizeImageCommand,
  setImageAltCommand,
  changePictureCommand,
  compressPictureCommand,
  resetPictureCommand,
  cropPictureCommand,
  cropToAspectCommand,
  pictureShapeCommand,
  pictureBorderCommand,
  pictureEffectsCommand,
  pictureAdjustCommand,
  pictureStyleCommand,
  pictureTransformCommand,
];
