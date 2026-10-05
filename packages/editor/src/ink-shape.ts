/**
 * Draw ▸ Ink to Shape: recognise a hand-drawn stroke as the closest preset
 * shape (line, triangle, rectangle, diamond, pentagon, hexagon or ellipse).
 * The stroke is simplified to its corners (Ramer–Douglas–Peucker) and judged
 * by corner count, closure and roundness.
 */

import type { ShapePreset } from "@office-kit/docx";

export type InkPoint = readonly [number, number];

export interface RecognizedShape {
  readonly preset: ShapePreset;
  /** Box in the stroke's coordinates. */
  readonly left: number;
  readonly top: number;
  readonly width: number;
  readonly height: number;
  readonly flipH: boolean;
  readonly flipV: boolean;
}

// Tolerances relative to the stroke's size.
const CLOSED_GAP = 0.25;
const STRAIGHT_DEVIATION = 0.08;
const SIMPLIFY_EPSILON = 0.1;
const ROUNDNESS = 0.12;
const MIDPOINT_TOLERANCE = 0.2;
const MIN_SIZE = 4;

function distanceToSegment([px, py]: InkPoint, [ax, ay]: InkPoint, [bx, by]: InkPoint): number {
  const dx = bx - ax;
  const dy = by - ay;
  const len2 = dx * dx + dy * dy;
  const t = len2 === 0 ? 0 : Math.max(0, Math.min(1, ((px - ax) * dx + (py - ay) * dy) / len2));
  return Math.hypot(px - (ax + t * dx), py - (ay + t * dy));
}

/** Ramer–Douglas–Peucker simplification. */
export function simplify(points: readonly InkPoint[], epsilon: number): InkPoint[] {
  if (points.length < 3) return [...points];
  const first = points[0];
  const last = points.at(-1);
  if (!first || !last) return [...points];
  let index = -1;
  let max = 0;
  for (let i = 1; i < points.length - 1; i++) {
    const p = points[i];
    if (!p) continue;
    const d = distanceToSegment(p, first, last);
    if (d > max) {
      max = d;
      index = i;
    }
  }
  if (max <= epsilon || index < 0) return [first, last];
  const left = simplify(points.slice(0, index + 1), epsilon);
  const right = simplify(points.slice(index), epsilon);
  return [...left.slice(0, -1), ...right];
}

export function recognizeInkShape(points: readonly InkPoint[]): RecognizedShape | undefined {
  const first = points[0];
  const last = points.at(-1);
  if (!first || !last || points.length < 2) return undefined;
  const xs = points.map((p) => p[0]);
  const ys = points.map((p) => p[1]);
  const left = Math.min(...xs);
  const top = Math.min(...ys);
  const width = Math.max(...xs) - left;
  const height = Math.max(...ys) - top;
  const size = Math.max(width, height);
  if (size < MIN_SIZE) return undefined;
  const box = { left, top, width, height, flipH: false, flipV: false };

  const closed = Math.hypot(last[0] - first[0], last[1] - first[1]) < CLOSED_GAP * size;
  if (!closed) {
    const deviation = Math.max(...points.map((p) => distanceToSegment(p, first, last)));
    if (deviation > STRAIGHT_DEVIATION * size) return undefined;
    return { ...box, preset: "line", flipH: last[0] < first[0] !== last[1] < first[1], flipV: false };
  }

  const cx = left + width / 2;
  const cy = top + height / 2;
  // Roundness: distance to the centre, normalised by the box's half-axes.
  const radii = points.map(([x, y]) => Math.hypot((x - cx) / (width / 2 || 1), (y - cy) / (height / 2 || 1)));
  const mean = radii.reduce((a, b) => a + b, 0) / radii.length;
  const spread = Math.sqrt(radii.reduce((a, r) => a + (r - mean) ** 2, 0) / radii.length) / (mean || 1);

  const corners = simplify(points, SIMPLIFY_EPSILON * size);
  // The closing point duplicates the first corner.
  const vertices = corners.length > 1 ? corners.slice(0, -1) : corners;
  const n = vertices.length;
  if (spread < ROUNDNESS && n >= 5) return { ...box, preset: "ellipse" };
  switch (n) {
    case 3: {
      // An apex below the base is a flipped triangle.
      const apex = vertices.reduce((a, b) => (Math.abs(b[1] - cy) > Math.abs(a[1] - cy) ? b : a));
      return { ...box, preset: "triangle", flipV: apex[1] > cy };
    }
    case 4: {
      const nearMid = vertices.every(
        ([x, y]) =>
          Math.abs(x - cx) < MIDPOINT_TOLERANCE * width || Math.abs(y - cy) < MIDPOINT_TOLERANCE * height,
      );
      return { ...box, preset: nearMid ? "diamond" : "rect" };
    }
    case 5:
      return { ...box, preset: "pentagon" };
    case 6:
      return { ...box, preset: "hexagon" };
    default:
      return { ...box, preset: "ellipse" };
  }
}
