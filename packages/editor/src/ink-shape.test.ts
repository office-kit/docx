import { describe, expect, it } from "vitest";
import { type InkPoint, recognizeInkShape } from "./ink-shape.js";

/** Points along a polygon's edges, as a hand would draw them. */
function trace(corners: readonly InkPoint[], steps = 10): InkPoint[] {
  const out: InkPoint[] = [];
  corners.forEach((a, i) => {
    const b = corners[(i + 1) % corners.length] ?? a;
    for (let s = 0; s < steps; s++)
      out.push([a[0] + ((b[0] - a[0]) * s) / steps, a[1] + ((b[1] - a[1]) * s) / steps]);
  });
  const start = corners[0];
  if (start) out.push(start);
  return out;
}

describe("recognizeInkShape", () => {
  it("recognises a straight stroke as a line", () => {
    const shape = recognizeInkShape([
      [0, 0],
      [50, 51],
      [100, 100],
    ]);
    expect(shape).toMatchObject({ preset: "line", width: 100, height: 100 });
  });

  it("leaves a squiggle alone", () => {
    expect(
      recognizeInkShape([
        [0, 0],
        [30, 60],
        [60, 0],
        [100, 70],
      ]),
    ).toBeUndefined();
  });

  it("recognises rectangles, diamonds and triangles", () => {
    expect(
      recognizeInkShape(
        trace([
          [0, 0],
          [100, 0],
          [100, 60],
          [0, 60],
        ]),
      )?.preset,
    ).toBe("rect");
    expect(
      recognizeInkShape(
        trace([
          [50, 0],
          [100, 50],
          [50, 100],
          [0, 50],
        ]),
      )?.preset,
    ).toBe("diamond");
    expect(
      recognizeInkShape(
        trace([
          [50, 0],
          [100, 90],
          [0, 90],
        ]),
      ),
    ).toMatchObject({ preset: "triangle", flipV: false });
  });

  it("recognises a circle as an ellipse", () => {
    const points: InkPoint[] = [];
    for (let i = 0; i <= 40; i++) {
      const a = (i / 40) * 2 * Math.PI;
      points.push([50 + 50 * Math.cos(a), 30 + 30 * Math.sin(a)]);
    }
    expect(recognizeInkShape(points)?.preset).toBe("ellipse");
  });
});
