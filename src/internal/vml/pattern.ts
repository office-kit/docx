/**
 * Pattern fills. VML draws a pattern fill from a small two-colour bitmap that
 * the fill references by relationship (`v:fill type="pattern" r:id`), painting
 * its dark pixels in the fill colour and its light pixels in `color2`
 * (ECMA-376 Part 4, §19.1.2.5). Word's pattern gallery is a fixed set of 8 × 8
 * tiles; each is described here as eight rows of bits, most significant bit
 * on the left, 1 = foreground.
 */

export const FILL_PATTERNS = {
  pct5: [0x80, 0x00, 0x00, 0x00, 0x08, 0x00, 0x00, 0x00],
  pct10: [0x80, 0x00, 0x08, 0x00, 0x80, 0x00, 0x08, 0x00],
  pct20: [0x88, 0x00, 0x22, 0x00, 0x88, 0x00, 0x22, 0x00],
  pct25: [0x88, 0x22, 0x88, 0x22, 0x88, 0x22, 0x88, 0x22],
  pct50: [0xaa, 0x55, 0xaa, 0x55, 0xaa, 0x55, 0xaa, 0x55],
  pct75: [0x77, 0xdd, 0x77, 0xdd, 0x77, 0xdd, 0x77, 0xdd],
  ltHorz: [0xff, 0x00, 0x00, 0x00, 0xff, 0x00, 0x00, 0x00],
  ltVert: [0x88, 0x88, 0x88, 0x88, 0x88, 0x88, 0x88, 0x88],
  dkHorz: [0xff, 0xff, 0x00, 0x00, 0xff, 0xff, 0x00, 0x00],
  dkVert: [0xcc, 0xcc, 0xcc, 0xcc, 0xcc, 0xcc, 0xcc, 0xcc],
  ltDnDiag: [0x88, 0x44, 0x22, 0x11, 0x88, 0x44, 0x22, 0x11],
  ltUpDiag: [0x11, 0x22, 0x44, 0x88, 0x11, 0x22, 0x44, 0x88],
  smGrid: [0xff, 0x88, 0x88, 0x88, 0xff, 0x88, 0x88, 0x88],
  smCheck: [0x99, 0x66, 0x66, 0x99, 0x99, 0x66, 0x66, 0x99],
  diagCross: [0x81, 0x42, 0x24, 0x18, 0x18, 0x24, 0x42, 0x81],
  dotGrid: [0xaa, 0x00, 0x80, 0x00, 0x80, 0x00, 0x80, 0x00],
} as const satisfies Readonly<Record<string, readonly number[]>>;

export type FillPattern = keyof typeof FILL_PATTERNS;

const SIZE = 8;
const FILE_HEADER = 14;
const INFO_HEADER = 40;
const PALETTE = 8;
// 1-bit rows are padded to a 4-byte boundary.
const ROW_BYTES = 4;

/**
 * Encode a pattern as a 1-bit BMP (black foreground, white background), the
 * format Word itself stores pattern tiles in.
 */
export function patternBitmap(pattern: FillPattern): Uint8Array {
  const rows = FILL_PATTERNS[pattern];
  const offset = FILE_HEADER + INFO_HEADER + PALETTE;
  const size = offset + ROW_BYTES * SIZE;
  const bytes = new Uint8Array(size);
  const view = new DataView(bytes.buffer);
  bytes[0] = 0x42; // "B"
  bytes[1] = 0x4d; // "M"
  view.setUint32(2, size, true);
  view.setUint32(10, offset, true);
  view.setUint32(14, INFO_HEADER, true);
  view.setInt32(18, SIZE, true);
  view.setInt32(22, SIZE, true);
  view.setUint16(26, 1, true); // planes
  view.setUint16(28, 1, true); // bits per pixel
  view.setUint32(34, ROW_BYTES * SIZE, true);
  // Palette: index 0 = black (foreground), index 1 = white.
  view.setUint32(FILE_HEADER + INFO_HEADER + 4, 0x00ffffff, true);
  for (let y = 0; y < SIZE; y++) {
    // BMP rows run bottom-up; a set bit selects palette index 1 (white), so
    // the pattern's foreground bits are inverted.
    const bits = rows[SIZE - 1 - y] ?? 0;
    bytes[offset + y * ROW_BYTES] = ~bits & 0xff;
  }
  return bytes;
}
