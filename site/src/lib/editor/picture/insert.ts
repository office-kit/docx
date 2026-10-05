/**
 * Getting pictures into the document from the browser: files (Insert ▸
 * Pictures, drag and drop, paste), screen captures and the icon library, plus
 * Compress Pictures. Formats a .docx cannot carry as ECMA-376 blips (WebP,
 * AVIF, SVG — SVG blips are an Office extension) are re-encoded as PNG.
 */

import { commands, EMU_PER_PX } from "@office-kit/docx-editor";
import {
  type DrawingInfo,
  imageDrawings,
  imageNaturalSizeEmu,
  readDrawing,
} from "@office-kit/docx";
import type { EditorSession } from "../session.svelte";
import { fitWidth, textWidthEmu } from "./state";

// An icon is drawn into a bitmap this many pixels square, then inserted at 1".
const ICON_RASTER_PX = 512;
const ICON_SIZE_EMU = 914400;
const PX_PER_INCH = 96;
const EMU_PER_INCH = 914400;
const JPEG_QUALITY = 0.9;

async function blobBytes(blob: Blob): Promise<Uint8Array> {
  return new Uint8Array(await blob.arrayBuffer());
}

function canvasBlob(canvas: HTMLCanvasElement, type: string): Promise<Blob> {
  return new Promise((resolve, reject) => {
    canvas.toBlob(
      (b) => (b ? resolve(b) : reject(new Error("The browser could not encode the image."))),
      type,
      JPEG_QUALITY,
    );
  });
}

/** Draw an image source into a canvas of `w` × `h` and encode it. */
async function encode(
  source: CanvasImageSource,
  w: number,
  h: number,
  type: string,
  crop?: { sx: number; sy: number; sw: number; sh: number },
): Promise<Uint8Array> {
  const canvas = document.createElement("canvas");
  canvas.width = Math.max(1, Math.round(w));
  canvas.height = Math.max(1, Math.round(h));
  const ctx = canvas.getContext("2d");
  if (!ctx) throw new Error("Canvas 2D is not available.");
  if (crop)
    ctx.drawImage(source, crop.sx, crop.sy, crop.sw, crop.sh, 0, 0, canvas.width, canvas.height);
  else ctx.drawImage(source, 0, 0, canvas.width, canvas.height);
  return blobBytes(await canvasBlob(canvas, type));
}

/**
 * Bytes a .docx can embed and the picture's natural size. PNG, JPEG, GIF and
 * BMP go in as they are (their headers give the size and DPI); anything else
 * the browser can decode is re-encoded as PNG at 96 DPI.
 */
async function docxImage(
  blob: Blob,
): Promise<{ bytes: Uint8Array; widthEmu: number; heightEmu: number }> {
  const original = await blobBytes(blob);
  const size = imageNaturalSizeEmu(original);
  if (size) return { bytes: original, ...size };
  const bitmap = await createImageBitmap(blob);
  const bytes = await encode(bitmap, bitmap.width, bitmap.height, "image/png");
  const natural = { widthEmu: bitmap.width * EMU_PER_PX, heightEmu: bitmap.height * EMU_PER_PX };
  bitmap.close();
  return { bytes, ...natural };
}

/** Insert image files at the caret, each scaled down to the text column like Word does. */
export async function insertPictureFiles(
  session: EditorSession,
  files: readonly Blob[],
): Promise<void> {
  const images = await Promise.all(files.filter((f) => f.type.startsWith("image/")).map(docxImage));
  const model = session.model;
  if (!model) return;
  const column = textWidthEmu(model.doc);
  for (const { bytes, ...size } of images) {
    session.apply(commands.insertImageCommand, { bytes, options: fitWidth(size, column) });
  }
}

/** Insert ▸ Screenshot: capture a screen, window or tab the user picks. */
export async function insertScreenshot(session: EditorSession): Promise<void> {
  const stream = await navigator.mediaDevices.getDisplayMedia({ video: true, audio: false });
  try {
    const video = document.createElement("video");
    video.muted = true;
    video.srcObject = stream;
    await video.play();
    const bytes = await encode(video, video.videoWidth, video.videoHeight, "image/png");
    const model = session.model;
    if (!model) return;
    const size = fitWidth(
      { widthEmu: video.videoWidth * EMU_PER_PX, heightEmu: video.videoHeight * EMU_PER_PX },
      textWidthEmu(model.doc),
    );
    session.apply(commands.insertImageCommand, { bytes, options: { ...size, name: "Screenshot" } });
  } finally {
    for (const track of stream.getTracks()) track.stop();
  }
}

async function rasterizeSvg(svg: string): Promise<Uint8Array> {
  const url = URL.createObjectURL(new Blob([svg], { type: "image/svg+xml" }));
  try {
    const img = new Image();
    img.src = url;
    await img.decode();
    return await encode(img, ICON_RASTER_PX, ICON_RASTER_PX, "image/png");
  } finally {
    URL.revokeObjectURL(url);
  }
}

/** Insert ▸ Icons: rasterize SVG icons (PNG, since SVG blips are an extension) and insert each at 1". */
export async function insertIcons(
  session: EditorSession,
  icons: ReadonlyArray<{ svg: string; name: string }>,
): Promise<void> {
  const rasters = await Promise.all(icons.map((icon) => rasterizeSvg(icon.svg)));
  icons.forEach(({ name }, i) => {
    const bytes = rasters[i];
    if (!bytes) return;
    session.apply(commands.insertImageCommand, {
      bytes,
      options: { widthEmu: ICON_SIZE_EMU, heightEmu: ICON_SIZE_EMU, name, altText: name },
    });
  });
}

/** Change Picture ▸ From a File: new bytes into the selected picture. */
export async function changePictureFrom(
  session: EditorSession,
  index: number,
  file: Blob,
): Promise<void> {
  const { bytes } = await docxImage(file);
  session.apply(commands.changePictureCommand, { index, bytes });
}

/**
 * Compress Pictures: re-encode each picture at `ppi` for its displayed size
 * (never upscaling), optionally cutting away the cropped areas.
 */
export async function compressPictures(
  session: EditorSession,
  indices: readonly number[],
  ppi: number | undefined,
  deleteCropped: boolean,
): Promise<void> {
  const model = session.model;
  if (!model) return;
  const drawings = imageDrawings(model.doc);
  const jobs = indices.flatMap((index) => {
    const drawing = drawings[index];
    const info: DrawingInfo | undefined = drawing ? readDrawing(model.doc, drawing) : undefined;
    return info?.picture?.image ? [{ index, info, image: info.picture.image }] : [];
  });
  const results = await Promise.all(
    jobs.map(async ({ index, info, image }) => {
      const bitmap = await createImageBitmap(
        new Blob([image.data.slice().buffer], { type: image.contentType }),
      );
      const crop = info.picture?.crop ?? { left: 0, top: 0, right: 0, bottom: 0 };
      const region = deleteCropped
        ? {
            sx: (Math.max(0, crop.left) / 100) * bitmap.width,
            sy: (Math.max(0, crop.top) / 100) * bitmap.height,
            sw: bitmap.width * (1 - (Math.max(0, crop.left) + Math.max(0, crop.right)) / 100),
            sh: bitmap.height * (1 - (Math.max(0, crop.top) + Math.max(0, crop.bottom)) / 100),
          }
        : { sx: 0, sy: 0, sw: bitmap.width, sh: bitmap.height };
      // The visible part's displayed size decides the pixel budget.
      const shownW = (info.widthEmu / EMU_PER_INCH) * (ppi ?? PX_PER_INCH);
      const shownH = (info.heightEmu / EMU_PER_INCH) * (ppi ?? PX_PER_INCH);
      const scale = ppi === undefined ? 1 : Math.min(1, shownW / region.sw, shownH / region.sh);
      const type = image.contentType === "image/jpeg" ? "image/jpeg" : "image/png";
      const bytes = await encode(bitmap, region.sw * scale, region.sh * scale, type, region);
      bitmap.close();
      return { index, bytes };
    }),
  );
  for (const { index, bytes } of results) {
    session.apply(commands.compressPictureCommand, { index, bytes, croppedAway: deleteCropped });
  }
}

/** The color of the image pixel under a point given as fractions of the picture frame. */
export async function pixelColorAt(
  info: DrawingInfo,
  fx: number,
  fy: number,
): Promise<string | undefined> {
  const image = info.picture?.image;
  if (!image) return undefined;
  const crop = info.picture?.crop ?? { left: 0, top: 0, right: 0, bottom: 0 };
  const bitmap = await createImageBitmap(
    new Blob([image.data.slice().buffer], { type: image.contentType }),
  );
  const u = crop.left / 100 + fx * (1 - (crop.left + crop.right) / 100);
  const v = crop.top / 100 + fy * (1 - (crop.top + crop.bottom) / 100);
  const canvas = document.createElement("canvas");
  canvas.width = 1;
  canvas.height = 1;
  const ctx = canvas.getContext("2d");
  if (!ctx) return undefined;
  ctx.drawImage(
    bitmap,
    Math.floor(u * bitmap.width),
    Math.floor(v * bitmap.height),
    1,
    1,
    0,
    0,
    1,
    1,
  );
  bitmap.close();
  const [r = 0, g = 0, b = 0] = ctx.getImageData(0, 0, 1, 1).data;
  return [r, g, b]
    .map((c) => c.toString(16).padStart(2, "0"))
    .join("")
    .toUpperCase();
}
