/**
 * SmartArt on the canvas. Word lays a diagram out from its data model and
 * layout definition with a general constraint engine; the canvas instead
 * implements the layouts the editor inserts (Basic Block List, Vertical
 * Bullet List, Basic Process, Basic Cycle, Hierarchy) directly, drawing other
 * layouts as a block list, and colours shapes from the diagram's colour and
 * style definitions over the document theme.
 */

import {
  type Docx,
  getRawPartRoot,
  getSmartArt,
  runSmartArt,
  SMARTART_COLORS,
  SMARTART_STYLES,
  type SmartArtColors,
  type SmartArtInfo,
  type SmartArtNode,
  type SmartArtStyle,
  type WmlRun,
  type XmlElement,
  xmlPartNames,
} from "@office-kit/docx";
import type { DocPosition } from "./selection.js";

const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";

/** Office theme colours, used when the document has no theme part. */
const DEFAULT_THEME: Readonly<Record<string, string>> = {
  dk1: "000000",
  lt1: "FFFFFF",
  dk2: "44546A",
  lt2: "E7E6E6",
  accent1: "4472C4",
  accent2: "ED7D31",
  accent3: "A5A5A5",
  accent4: "FFC000",
  accent5: "5B9BD5",
  accent6: "70AD47",
};

const themeCache = new WeakMap<Docx, Readonly<Record<string, string>>>();

/** The theme's colour scheme (`a:clrScheme`), by scheme colour name. */
export function themeColors(doc: Docx): Readonly<Record<string, string>> {
  const cached = themeCache.get(doc);
  if (cached) return cached;
  const out: Record<string, string> = { ...DEFAULT_THEME };
  const name = xmlPartNames(doc).find((n) => /\/theme\/theme\d*\.xml$/.test(n));
  const root = name ? getRawPartRoot(doc, name) : undefined;
  const scheme = root && findA(root, "clrScheme");
  for (const c of scheme?.children ?? []) {
    if (c.kind !== "element") continue;
    const color = c.children.find((x): x is XmlElement => x.kind === "element");
    const val = color?.attrs.find(
      (a) => a.name.local === (color.name.local === "sysClr" ? "lastClr" : "val"),
    )?.value;
    if (val && /^[0-9A-Fa-f]{6}$/.test(val)) out[c.name.local] = val.toUpperCase();
  }
  themeCache.set(doc, out);
  return out;
}

function findA(el: XmlElement, local: string): XmlElement | undefined {
  if (el.name.uri === A_NS && el.name.local === local) return el;
  for (const c of el.children) {
    if (c.kind !== "element") continue;
    const found = findA(c, local);
    if (found) return found;
  }
  return undefined;
}

interface Box {
  readonly x: number;
  readonly y: number;
  readonly w: number;
  readonly h: number;
}

interface Painter {
  readonly fill: (index: number) => string;
  readonly line: string;
  readonly lineWidth: number;
  readonly text: string;
  readonly shadow: string;
  readonly accent: string;
}

const escapeHtml = (s: string): string =>
  s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
const n2 = (n: number): string => String(Math.round(n * 100) / 100);

function painter(
  doc: Docx,
  colors: SmartArtColors | undefined,
  style: SmartArtStyle | undefined,
): Painter {
  const theme = themeColors(doc);
  const fills = SMARTART_COLORS[colors ?? "accent1_2"].fill.map((c) => `#${theme[c] ?? "4472C4"}`);
  const s = SMARTART_STYLES[style ?? "simple1"];
  const shadows = [
    "",
    "drop-shadow(0 1pt 1.5pt rgba(0,0,0,.25))",
    "drop-shadow(0 2pt 3pt rgba(0,0,0,.35))",
    "drop-shadow(0 3pt 5pt rgba(0,0,0,.45))",
  ];
  return {
    fill: (i) => fills[i % fills.length] ?? "#4472C4",
    line: `#${theme.lt1 ?? "FFFFFF"}`,
    lineWidth: s.line >= 3 ? 2.25 : 1,
    text: `#${theme.lt1 ?? "FFFFFF"}`,
    shadow: shadows[s.effect] ?? "",
    accent: fills[0] ?? "#4472C4",
  };
}

/** Font size that fits `lines` into a box, capped like Word's 65 pt primary font. */
function fitFont(lines: readonly string[], box: Box): number {
  const longest = Math.max(1, ...lines.map((l) => l.length));
  const byWidth = (box.w * 0.85) / (longest * 0.55);
  const byHeight = (box.h * 0.8) / (lines.length * 1.2);
  return Math.max(5, Math.min(65, byWidth, byHeight));
}

function nodeLines(node: SmartArtNode): string[] {
  const out = [node.text];
  for (const child of node.children ?? []) out.push(`• ${child.text}`);
  return out;
}

function textHtml(
  lines: readonly string[],
  box: Box,
  color: string,
  size: number,
  align = "center",
): string {
  // Spans only: the diagram renders inside a body <p>.
  const body = lines
    .map((l) => `<span style="display:block">${escapeHtml(l) || "​"}</span>`)
    .join("");
  return (
    `<span class="wk-smartart-text" style="position:absolute;left:${n2(box.x)}pt;top:${n2(box.y)}pt;width:${n2(box.w)}pt;` +
    `height:${n2(box.h)}pt;display:flex;flex-direction:column;justify-content:center;text-align:${align};` +
    `font-size:${n2(size)}pt;line-height:1.15;color:${color};overflow:hidden;padding:0 4pt;box-sizing:border-box">${body}</span>`
  );
}

function rect(box: Box, fill: string, p: Painter, radius = 0): string {
  return (
    `<rect x="${n2(box.x)}" y="${n2(box.y)}" width="${n2(box.w)}" height="${n2(box.h)}" rx="${n2(radius)}" ` +
    `fill="${fill}" stroke="${p.line}" stroke-width="${p.lineWidth}"/>`
  );
}

function treeDepth(list: readonly SmartArtNode[]): number {
  return list.length ? 1 + Math.max(...list.map((c) => treeDepth(c.children ?? []))) : 0;
}

function leafCount(node: SmartArtNode): number {
  return node.children?.length ? node.children.reduce((s, c) => s + leafCount(c), 0) : 1;
}

/** Lay out and draw a diagram in a `width` × `height` point frame. */
function drawDiagram(info: SmartArtInfo, p: Painter): { svg: string; text: string } {
  const { width: W, height: H, nodes } = info;
  let svg = "";
  let text = "";
  const n = Math.max(nodes.length, 1);
  switch (info.layout) {
    case "basicProcess": {
      // n boxes and n-1 arrows: box h = 0.6 w, arrow 0.2 w wide with a gap each side.
      const unit = n + (n - 1) * 0.5;
      const w = Math.min(W / unit, H / 0.6);
      const h = w * 0.6;
      const y = (H - h) / 2;
      const x0 = (W - w * unit) / 2;
      nodes.forEach((node, i) => {
        const box = { x: x0 + i * w * 1.5, y, w, h };
        svg += rect(box, p.fill(i), p, w * 0.1);
        const lines = nodeLines(node);
        text += textHtml(lines, box, p.text, fitFont(lines, box));
        if (i < nodes.length - 1) {
          const ax = box.x + w * 1.08;
          const aw = w * 0.34;
          const ah = h * 0.35;
          const ay = H / 2;
          svg +=
            `<path d="M${n2(ax)} ${n2(ay - ah * 0.3)} H${n2(ax + aw * 0.6)} V${n2(ay - ah / 2)} L${n2(ax + aw)} ${n2(ay)} ` +
            `L${n2(ax + aw * 0.6)} ${n2(ay + ah / 2)} V${n2(ay + ah * 0.3)} H${n2(ax)} Z" fill="${p.accent}" fill-opacity="0.6"/>`;
        }
      });
      break;
    }
    case "verticalBulletList": {
      const rows = nodes.map((node) => 1 + (node.children?.length ?? 0) * 0.6);
      const total = rows.reduce((a, b) => a + b, 0) + (n - 1) * 0.1;
      const unit = H / Math.max(total, 1);
      let y = 0;
      nodes.forEach((node, i) => {
        const box = { x: 0, y, w: W, h: unit };
        svg += rect(box, p.fill(i), p, unit * 0.17);
        text += textHtml([node.text], box, p.text, fitFont([node.text], box), "left");
        y += unit;
        const kids = (node.children ?? []).map((c) => `• ${c.text}`);
        if (kids.length) {
          const kidBox = { x: W * 0.05, y, w: W * 0.95, h: unit * 0.6 * kids.length };
          text += textHtml(kids, kidBox, "#000", fitFont(kids, kidBox) * 0.8, "left");
          y += kidBox.h;
        }
        y += unit * 0.1;
      });
      break;
    }
    case "basicCycle": {
      const cx = W / 2;
      const cy = H / 2;
      const ring = Math.min(W, H) / 2;
      const r = Math.min(ring * 0.42, (Math.PI * ring) / (n + 1) / 1.1);
      const orbit = ring - r;
      nodes.forEach((node, i) => {
        const a = -Math.PI / 2 + (i * 2 * Math.PI) / n;
        const x = cx + orbit * Math.cos(a);
        const y = cy + orbit * Math.sin(a);
        svg += `<circle cx="${n2(x)}" cy="${n2(y)}" r="${n2(r)}" fill="${p.fill(i)}" stroke="${p.line}" stroke-width="${p.lineWidth}"/>`;
        const box = { x: x - r * 0.75, y: y - r * 0.75, w: r * 1.5, h: r * 1.5 };
        const lines = nodeLines(node);
        text += textHtml(lines, box, p.text, fitFont(lines, box));
        if (n > 1) {
          // A small arrowhead halfway to the next node, pointing along the orbit.
          const mid = a + Math.PI / n;
          const mx = cx + orbit * Math.cos(mid);
          const my = cy + orbit * Math.sin(mid);
          const s = r * 0.3;
          const deg = (mid * 180) / Math.PI + 90;
          svg +=
            `<path d="M${n2(-s)} ${n2(-s * 0.6)} L${n2(s)} 0 L${n2(-s)} ${n2(s * 0.6)} Z" fill="${p.accent}" fill-opacity="0.6" ` +
            `transform="translate(${n2(mx)} ${n2(my)}) rotate(${n2(deg)})"/>`;
        }
      });
      break;
    }
    case "hierarchy": {
      const levels = Math.max(treeDepth(nodes), 1);
      const totalLeaves = Math.max(
        nodes.reduce((s, c) => s + leafCount(c), 0),
        1,
      );
      const slot = W / totalLeaves;
      const rowH = H / levels;
      const bw = slot * 0.8;
      const bh = Math.min(rowH * 0.7, bw * 0.635);
      let index = 0;
      const place = (node: SmartArtNode, level: number, left: number): { x: number; y: number } => {
        const span = leafCount(node) * slot;
        const x = left + span / 2;
        const y = level * rowH + (rowH - bh) / 2;
        const box = { x: x - bw / 2, y, w: bw, h: bh };
        svg += rect(box, p.fill(index++), p, bw * 0.08);
        text += textHtml([node.text], box, p.text, fitFont([node.text], box));
        let childLeft = left;
        for (const child of node.children ?? []) {
          const c = place(child, level + 1, childLeft);
          childLeft += leafCount(child) * slot;
          const midY = y + bh + (c.y - y - bh) / 2;
          svg +=
            `<path d="M${n2(x)} ${n2(y + bh)} V${n2(midY)} H${n2(c.x)} V${n2(c.y)}" fill="none" ` +
            `stroke="${p.accent}" stroke-width="1"/>`;
        }
        return { x, y };
      };
      let left = 0;
      for (const node of nodes) {
        place(node, 0, left);
        left += leafCount(node) * slot;
      }
      break;
    }
    default: {
      // Basic Block List: a grid of 0.6-aspect blocks with 10 % gaps, rows centred.
      let best = { cols: 1, w: 0 };
      for (let cols = 1; cols <= n; cols++) {
        const rows = Math.ceil(n / cols);
        const w = Math.min(W / (cols + (cols - 1) * 0.1), H / (rows * 0.6 + (rows - 1) * 0.1));
        if (w > best.w) best = { cols, w };
      }
      const { cols, w } = best;
      const h = w * 0.6;
      const rows = Math.ceil(n / cols);
      const top = (H - (rows * h + (rows - 1) * w * 0.1)) / 2;
      nodes.forEach((node, i) => {
        const row = Math.floor(i / cols);
        const inRow = Math.min(cols, n - row * cols);
        const rowLeft = (W - (inRow * w + (inRow - 1) * w * 0.1)) / 2;
        const box = { x: rowLeft + (i % cols) * w * 1.1, y: top + row * (h + w * 0.1), w, h };
        svg += rect(box, p.fill(i), p);
        const lines = nodeLines(node);
        text += textHtml(lines, box, p.text, fitFont(lines, box));
      });
    }
  }
  return { svg, text };
}

/** HTML for the SmartArt graphics a run holds (none when it holds no diagram). */
export function renderRunSmartArt(run: WmlRun, doc: Docx, at: DocPosition): string {
  const ref = runSmartArt(doc, run);
  if (!ref) return "";
  const info = getSmartArt(doc, ref);
  const p = painter(doc, info.colors, info.style);
  const { svg, text } = drawDiagram(info, p);
  return (
    `<span class="wk-shape wk-smartart" contenteditable="false" data-wk-object="smartArt" ` +
    `data-wk-at="${escapeHtml(JSON.stringify(at))}" style="display:inline-block;position:relative;vertical-align:bottom;` +
    `width:${n2(info.width)}pt;height:${n2(info.height)}pt">` +
    `<svg width="100%" height="100%" viewBox="0 0 ${n2(info.width)} ${n2(info.height)}" overflow="visible" ` +
    `style="position:absolute;inset:0${p.shadow ? `;filter:${p.shadow}` : ""}">${svg}</svg>${text}</span>`
  );
}
