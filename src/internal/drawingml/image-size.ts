/**
 * Pixel dimensions (and stored resolution) of PNG, JPEG, GIF and BMP images,
 * read from their headers. Word sizes a newly inserted picture from these:
 * pixels at the file's DPI, or 96 DPI when the file does not say.
 */

export interface ImagePixelSize {
  readonly width: number;
  readonly height: number;
  /** Horizontal / vertical resolution in dots per inch. */
  readonly dpiX: number;
  readonly dpiY: number;
}

const DEFAULT_DPI = 96;
const INCHES_PER_METER = 39.3701;
const CM_PER_INCH = 2.54;

function u16be(b: Uint8Array, i: number): number {
  return ((b[i] ?? 0) << 8) | (b[i + 1] ?? 0);
}
function u16le(b: Uint8Array, i: number): number {
  return (b[i] ?? 0) | ((b[i + 1] ?? 0) << 8);
}
function u32be(b: Uint8Array, i: number): number {
  return (
    (((b[i] ?? 0) << 24) >>> 0) + ((b[i + 1] ?? 0) << 16) + ((b[i + 2] ?? 0) << 8) + (b[i + 3] ?? 0)
  );
}
function i32le(b: Uint8Array, i: number): number {
  return (b[i] ?? 0) | ((b[i + 1] ?? 0) << 8) | ((b[i + 2] ?? 0) << 16) | ((b[i + 3] ?? 0) << 24);
}

function png(b: Uint8Array): ImagePixelSize | undefined {
  if (b.length < 24 || b[0] !== 0x89 || b[1] !== 0x50) return undefined;
  const width = u32be(b, 16);
  const height = u32be(b, 20);
  let dpiX = DEFAULT_DPI;
  let dpiY = DEFAULT_DPI;
  // Walk chunks for pHYs (pixels per unit; unit 1 = meter).
  let i = 8;
  while (i + 8 <= b.length) {
    const len = u32be(b, i);
    const type = String.fromCharCode(b[i + 4] ?? 0, b[i + 5] ?? 0, b[i + 6] ?? 0, b[i + 7] ?? 0);
    if (type === "pHYs" && b[i + 16] === 1) {
      dpiX = Math.round(u32be(b, i + 8) / INCHES_PER_METER) || DEFAULT_DPI;
      dpiY = Math.round(u32be(b, i + 12) / INCHES_PER_METER) || DEFAULT_DPI;
      break;
    }
    if (type === "IDAT" || type === "IEND") break;
    i += 12 + len;
  }
  return { width, height, dpiX, dpiY };
}

function jpeg(b: Uint8Array): ImagePixelSize | undefined {
  if (b.length < 4 || b[0] !== 0xff || b[1] !== 0xd8) return undefined;
  let dpiX = DEFAULT_DPI;
  let dpiY = DEFAULT_DPI;
  let i = 2;
  while (i + 9 < b.length) {
    if (b[i] !== 0xff) {
      i++;
      continue;
    }
    const marker = b[i + 1] ?? 0;
    const len = u16be(b, i + 2);
    // APP0 JFIF density: units 1 = dots per inch, 2 = dots per cm.
    if (marker === 0xe0 && b[i + 4] === 0x4a && b[i + 5] === 0x46) {
      const units = b[i + 11];
      const x = u16be(b, i + 12);
      const y = u16be(b, i + 14);
      if (units === 1 && x && y) [dpiX, dpiY] = [x, y];
      if (units === 2 && x && y)
        [dpiX, dpiY] = [Math.round(x * CM_PER_INCH), Math.round(y * CM_PER_INCH)];
    }
    // SOF0..SOF15 except DHT (C4), JPG (C8) and DAC (CC) carry the frame size.
    if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
      return { height: u16be(b, i + 5), width: u16be(b, i + 7), dpiX, dpiY };
    }
    i += 2 + len;
  }
  return undefined;
}

function gif(b: Uint8Array): ImagePixelSize | undefined {
  if (b.length < 10 || b[0] !== 0x47 || b[1] !== 0x49 || b[2] !== 0x46) return undefined;
  return { width: u16le(b, 6), height: u16le(b, 8), dpiX: DEFAULT_DPI, dpiY: DEFAULT_DPI };
}

function bmp(b: Uint8Array): ImagePixelSize | undefined {
  if (b.length < 26 || b[0] !== 0x42 || b[1] !== 0x4d) return undefined;
  return {
    width: Math.abs(i32le(b, 18)),
    height: Math.abs(i32le(b, 22)),
    dpiX: DEFAULT_DPI,
    dpiY: DEFAULT_DPI,
  };
}

/** The image's pixel size, or `undefined` for formats this reader does not know. */
export function readImagePixelSize(bytes: Uint8Array): ImagePixelSize | undefined {
  return png(bytes) ?? jpeg(bytes) ?? gif(bytes) ?? bmp(bytes);
}
