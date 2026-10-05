/**
 * VML path language (ECMA-376 Part 4, §19.1.2.14 `path`, with the shape
 * type formulas of §19.1.2.6 `v:f`) → SVG path data.
 *
 * VML coordinates live in the shape's coordinate space (`coordorigin` /
 * `coordsize`); they are mapped onto the shape's box in points, so stroke
 * widths and arrowheads stay in real units. Every arc form is converted to
 * SVG elliptical arcs or Béziers.
 */

/** One drawn subpath group: VML can mark subpaths as not filled (`nf`) or not stroked (`ns`). */
export interface SvgSubpath {
  readonly d: string;
  readonly fill: boolean;
  readonly stroke: boolean;
}

export interface VmlGeometry {
  /** `coordsize` (defaults 1000,1000) and `coordorigin` (0,0). */
  readonly coordSize: readonly [number, number];
  readonly coordOrigin?: readonly [number, number];
  /** Box size in points. */
  readonly width: number;
  readonly height: number;
  /** Adjust values (`adj`), referenced as `#n`. */
  readonly adj?: readonly number[];
  /** `v:f` equations, referenced as `@n`. */
  readonly formulas?: readonly string[];
}

// Angles in VML formulas and arc commands are in 1/65536ths of a degree.
const FD = 65536;
const DEG = Math.PI / 180;
// Control-point distance for a quarter ellipse as one cubic Bézier.
const KAPPA = 0.5523;

/** Evaluate shape type formulas (§19.1.2.6) in order; later ones may use earlier ones. */
export function evaluateFormulas(geometry: VmlGeometry): number[] {
  const [cw, ch] = geometry.coordSize;
  const [ox, oy] = geometry.coordOrigin ?? [0, 0];
  const values: number[] = [];
  const named: Readonly<Record<string, number>> = {
    width: cw,
    height: ch,
    xcenter: ox + cw / 2,
    ycenter: oy + ch / 2,
    xlimo: 0,
    ylimo: 0,
    hasstroke: 1,
    hasfill: 1,
    linedrawn: 1,
    pixellinewidth: 1,
    pixelwidth: geometry.width,
    pixelheight: geometry.height,
    emuwidth: geometry.width * 12700,
    emuheight: geometry.height * 12700,
    emuwidth2: geometry.width * 6350,
    emuheight2: geometry.height * 6350,
  };
  const arg = (token: string | undefined): number => {
    if (token === undefined) return 0;
    if (token.startsWith("#")) return geometry.adj?.[Number(token.slice(1))] ?? 0;
    if (token.startsWith("@")) return values[Number(token.slice(1))] ?? 0;
    const n = Number(token);
    return Number.isFinite(n) ? n : (named[token.toLowerCase()] ?? 0);
  };
  for (const eqn of geometry.formulas ?? []) {
    const [op = "val", ...rest] = eqn.trim().split(/\s+/);
    const [a = 0, b = 0, c = 0] = rest.map(arg);
    values.push(applyOp(op.toLowerCase(), a, b, c));
  }
  return values;
}

function applyOp(op: string, a: number, b: number, c: number): number {
  switch (op) {
    case "val":
      return a;
    case "sum":
      return a + b - c;
    case "prod":
      return c === 0 ? 0 : (a * b) / c;
    case "mid":
      return (a + b) / 2;
    case "abs":
      return Math.abs(a);
    case "min":
      return Math.min(a, b);
    case "max":
      return Math.max(a, b);
    case "if":
      return a > 0 ? b : c;
    case "mod":
      return Math.sqrt(a * a + b * b + c * c);
    case "atan2":
      return (Math.atan2(b, a) / DEG) * FD;
    case "sin":
      return a * Math.sin((b / FD) * DEG);
    case "cos":
      return a * Math.cos((b / FD) * DEG);
    case "tan":
      return a * Math.tan((b / FD) * DEG);
    case "cosatan2":
      return a * Math.cos(Math.atan2(c, b));
    case "sinatan2":
      return a * Math.sin(Math.atan2(c, b));
    case "sqrt":
      return Math.sqrt(Math.max(0, a));
    case "sumangle":
      return a + b * FD - c * FD;
    case "ellipse":
      return b === 0 ? 0 : c * Math.sqrt(Math.max(0, 1 - (a / b) ** 2));
    default:
      return 0;
  }
}

const COMMANDS = /nf|ns|ae|al|at|ar|wa|wr|qx|qy|qb|h[a-i]|[mlcxetrv]/y;

type Token = { cmd: string } | { num: number };

function tokenize(path: string, formulas: readonly number[], adj: readonly number[]): Token[] {
  const tokens: Token[] = [];
  const value = (text: string): number => {
    if (text.startsWith("@")) return formulas[Number(text.slice(1))] ?? 0;
    if (text.startsWith("#")) return adj[Number(text.slice(1))] ?? 0;
    const n = Number(text);
    return Number.isFinite(n) ? n : 0;
  };
  // Parameters between two commands are comma-separated, and an empty slot is
  // a zero: "m@0,l,21600r21600,xe" is m(@0,0) l(0,21600) r(21600,0) x e.
  const pushParams = (params: string): void => {
    if (params.trim() === "") return;
    for (const slot of params.split(",")) {
      const numbers = slot.match(/[@#]?-?[0-9]*\.?[0-9]+/g);
      if (!numbers) tokens.push({ num: 0 });
      else for (const n of numbers) tokens.push({ num: value(n) });
    }
  };
  let i = 0;
  let params = "";
  while (i < path.length) {
    COMMANDS.lastIndex = i;
    const cmd = COMMANDS.exec(path);
    if (cmd) {
      pushParams(params);
      params = "";
      tokens.push({ cmd: cmd[0] });
      i = COMMANDS.lastIndex;
    } else {
      params += path[i] ?? "";
      i++;
    }
  }
  pushParams(params);
  return tokens;
}

const ARITY: Readonly<Record<string, number>> = {
  m: 2,
  l: 2,
  t: 2,
  r: 2,
  c: 6,
  v: 6,
  ae: 6,
  al: 6,
  at: 8,
  ar: 8,
  wa: 8,
  wr: 8,
  qx: 2,
  qy: 2,
  qb: 4,
};

const fmt = (n: number): string => String(Math.round(n * 100) / 100);

/** Convert a VML path to SVG subpaths in points on the shape's box. */
export function vmlPathToSvg(path: string, geometry: VmlGeometry): SvgSubpath[] {
  const formulas = evaluateFormulas(geometry);
  const tokens = tokenize(path, formulas, geometry.adj ?? []);
  const [cw, ch] = geometry.coordSize;
  const [ox, oy] = geometry.coordOrigin ?? [0, 0];
  const sx = geometry.width / (cw || 1);
  const sy = geometry.height / (ch || 1);
  const X = (x: number): number => (x - ox) * sx;
  const Y = (y: number): number => (y - oy) * sy;

  const out: SvgSubpath[] = [];
  let d = "";
  let fill = true;
  let stroke = true;
  let cx = 0;
  let cy = 0;
  let startX = 0;
  let startY = 0;
  const flush = (): void => {
    if (d) out.push({ d: d.trim(), fill, stroke });
    d = "";
  };
  const moveTo = (x: number, y: number): void => {
    d += `M${fmt(X(x))} ${fmt(Y(y))} `;
    cx = x;
    cy = y;
    startX = x;
    startY = y;
  };
  const lineTo = (x: number, y: number): void => {
    d += `L${fmt(X(x))} ${fmt(Y(y))} `;
    cx = x;
    cy = y;
  };
  const curveTo = (x1: number, y1: number, x2: number, y2: number, x: number, y: number): void => {
    d += `C${fmt(X(x1))} ${fmt(Y(y1))} ${fmt(X(x2))} ${fmt(Y(y2))} ${fmt(X(x))} ${fmt(Y(y))} `;
    cx = x;
    cy = y;
  };
  /** Elliptical arc on the ellipse centred (ecx, ecy) with radii (rx, ry), between two angles (radians). */
  const arc = (
    ecx: number,
    ecy: number,
    rx: number,
    ry: number,
    from: number,
    sweep: number,
    connect: boolean,
  ): void => {
    const px = (a: number): number => ecx + rx * Math.cos(a);
    const py = (a: number): number => ecy + ry * Math.sin(a);
    if (connect && d) lineTo(px(from), py(from));
    else moveTo(px(from), py(from));
    // Split into ≤180° pieces so each SVG arc's flags are unambiguous; a full
    // circle needs two.
    const pieces = Math.max(1, Math.ceil(Math.abs(sweep) / Math.PI));
    for (let k = 1; k <= pieces; k++) {
      const a = from + (sweep * k) / pieces;
      const x = px(a);
      const y = py(a);
      d += `A${fmt(Math.abs(rx * sx))} ${fmt(Math.abs(ry * sy))} 0 0 ${sweep > 0 ? 1 : 0} ${fmt(X(x))} ${fmt(Y(y))} `;
      cx = x;
      cy = y;
    }
  };
  /** at / ar / wa / wr: arc on the ellipse inscribed in a box, from the angle of one point to another's. */
  const boxArc = (n: number[], clockwise: boolean, connect: boolean): void => {
    const [l = 0, t = 0, r = 0, b = 0, x1 = 0, y1 = 0, x2 = 0, y2 = 0] = n;
    const ecx = (l + r) / 2;
    const ecy = (t + b) / 2;
    const rx = (r - l) / 2;
    const ry = (b - t) / 2;
    const angle = (x: number, y: number): number =>
      Math.atan2((y - ecy) / (ry || 1), (x - ecx) / (rx || 1));
    const a1 = angle(x1, y1);
    let a2 = angle(x2, y2);
    // y grows downward, so a positive angle step is clockwise on screen.
    if (clockwise) while (a2 <= a1) a2 += 2 * Math.PI;
    else while (a2 >= a1) a2 -= 2 * Math.PI;
    arc(ecx, ecy, rx, ry, a1, a2 - a1, connect);
  };
  /** qx / qy: a quarter ellipse whose first tangent is horizontal (qx) or vertical (qy). */
  const quadrant = (x: number, y: number, horizontalFirst: boolean): void => {
    if (horizontalFirst) curveTo(cx + (x - cx) * KAPPA, cy, x, y - (y - cy) * KAPPA, x, y);
    else curveTo(cx, cy + (y - cy) * KAPPA, x - (x - cx) * KAPPA, y, x, y);
  };

  let cmd = "m";
  let qToggle = true;
  let k = 0;
  while (k < tokens.length) {
    const token = tokens[k];
    if (!token) break;
    if ("cmd" in token) {
      k++;
      cmd = token.cmd;
      qToggle = cmd === "qx";
      switch (cmd) {
        case "x":
          d += "Z ";
          cx = startX;
          cy = startY;
          continue;
        case "e":
          flush();
          fill = true;
          stroke = true;
          continue;
        case "nf":
          flush();
          fill = false;
          continue;
        case "ns":
          flush();
          stroke = false;
          continue;
        default:
          // The h* commands are rendering hints; they draw nothing.
          continue;
      }
    }
    const arity = ARITY[cmd];
    if (!arity) {
      k++;
      continue;
    }
    const n: number[] = [];
    for (let j = 0; j < arity; j++) {
      const t = tokens[k + j];
      if (!t || !("num" in t)) break;
      n.push(t.num);
    }
    if (n.length < arity) {
      k += n.length || 1;
      continue;
    }
    k += arity;
    const [a = 0, b = 0, c = 0, dd = 0, e = 0, f = 0] = n;
    switch (cmd) {
      case "m":
        moveTo(a, b);
        // Further pairs after a moveto are line segments.
        cmd = "l";
        break;
      case "t":
        moveTo(cx + a, cy + b);
        cmd = "r";
        break;
      case "l":
        lineTo(a, b);
        break;
      case "r":
        lineTo(cx + a, cy + b);
        break;
      case "c":
        curveTo(a, b, c, dd, e, f);
        break;
      case "v":
        curveTo(cx + a, cy + b, cx + c, cy + dd, cx + e, cy + f);
        break;
      case "ae":
      case "al":
        // Angles count counter-clockwise from the x axis (y up), in 1/65536°.
        arc(a, b, c, dd, (-e / FD) * DEG, (-f / FD) * DEG, cmd === "ae");
        break;
      case "at":
      case "ar":
        boxArc(n, false, cmd === "at");
        break;
      case "wa":
      case "wr":
        boxArc(n, true, cmd === "wa");
        break;
      case "qx":
      case "qy":
        // Successive pairs alternate between x-first and y-first quadrants.
        quadrant(a, b, qToggle);
        qToggle = !qToggle;
        break;
      case "qb":
        // Quadratic Bézier: control point, then end point.
        d += `Q${fmt(X(a))} ${fmt(Y(b))} ${fmt(X(c))} ${fmt(Y(dd))} `;
        cx = c;
        cy = dd;
        break;
    }
  }
  flush();
  return out;
}
