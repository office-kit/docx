<script lang="ts">
  /**
   * Pointer and keyboard handling for shapes on the canvas, kept out of
   * EditorCanvas: selecting VML shapes, text boxes, WordArt, ink and SmartArt
   * (Shift+click adds to the selection); moving, resizing and rotating them
   * with Word's handles; drawing a new shape by dragging after picking it in
   * a Shapes gallery; Draw tab pens, the stroke eraser and lasso select; and
   * typing inside text boxes. Commands run through the session as usual.
   *
   * The handles are drawn in a fixed layer in viewport coordinates, outside
   * the zoomed page, so they keep their size at every zoom.
   */
  import { tick } from 'svelte';
  import {
    getShapePoints,
    getShapeLayout,
    getShapeWrap,
    getSmartArt,
    shapeKind,
    type ShapeLayout,
    type ShapeStroke,
  } from '@office-kit/docx';
  import { commands, pageGeometry, recognizeInkShape, runCommand, type DocPosition, type InkPoint } from '@office-kit/docx-editor';
  import { getSession, type SelectedObject } from '../session.svelte';
  import { HIGHLIGHTER_OPACITY, PENCIL_OPACITY, samePosition, shapeTools } from './tools.svelte';
  import WordArtTextDialog from './WordArtTextDialog.svelte';
  import './shapes.css';

  type Props = { canvas: HTMLDivElement | null };
  const { canvas }: Props = $props();
  const session = getSession();

  const MY_KINDS: ReadonlySet<string> = new Set(['shape', 'textBox', 'ink', 'smartArt']);
  const POINTS_PER_PX = 0.75;
  const TWIPS_PER_POINT = 20;
  // A click without a drag draws Word's default 1" shape.
  const DEFAULT_SIZE = 72;
  const MIN_DRAG = 3;
  const NUDGE = 1;
  const HANDLES = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'] as const;
  type Handle = (typeof HANDLES)[number];

  /** Points per client pixel at the current zoom. */
  const k = (): number => POINTS_PER_PX / session.zoom;

  // ---- Finding objects --------------------------------------------------

  function parseAt(el: Element): DocPosition | undefined {
    const raw = el.getAttribute('data-wk-at');
    if (!raw) return undefined;
    const parsed: unknown = JSON.parse(raw);
    return isPosition(parsed) ? parsed : undefined;
  }

  function isPosition(v: unknown): v is DocPosition {
    return typeof v === 'object' && v !== null && typeof (v as { block?: unknown }).block === 'number';
  }

  function objectElement(at: DocPosition): HTMLElement | undefined {
    if (!canvas) return undefined;
    for (const el of canvas.querySelectorAll<HTMLElement>('[data-wk-object]')) {
      if (!MY_KINDS.has(el.dataset.wkObject ?? '')) continue;
      const pos = parseAt(el);
      if (pos && samePosition(pos, at)) return el;
    }
    return undefined;
  }

  function mine(obj: SelectedObject | null): obj is SelectedObject {
    return !!obj && MY_KINDS.has(obj.kind);
  }

  function layoutAt(at: DocPosition): ShapeLayout | undefined {
    const model = session.model;
    if (!model) return undefined;
    try {
      return getShapeLayout(commands.shapeAt(model.doc, at));
    } catch {
      // The object was SmartArt (not VML) or is gone after an undo.
      return undefined;
    }
  }

  // ---- Selection frame --------------------------------------------------

  let frames = $state<Array<{ rect: DOMRect; primary: boolean }>>([]);

  function measure(): void {
    const out: Array<{ rect: DOMRect; primary: boolean }> = [];
    const sel = session.selectedObject;
    if (mine(sel)) {
      const el = objectElement(sel.at);
      if (el) out.push({ rect: el.getBoundingClientRect(), primary: true });
    }
    for (const at of shapeTools.multi) {
      if (mine(sel) && samePosition(at, sel.at)) continue;
      const el = objectElement(at);
      if (el) out.push({ rect: el.getBoundingClientRect(), primary: false });
    }
    frames = out;
  }

  $effect(() => {
    // Re-measure after every render and selection change.
    if (session.version < 0 || session.tick < 0 || session.zoom <= 0) return;
    void session.selectedObject;
    void shapeTools.multi.length;
    let raf = requestAnimationFrame(measure);
    const onMove = (): void => {
      cancelAnimationFrame(raf);
      raf = requestAnimationFrame(measure);
    };
    window.addEventListener('scroll', onMove, true);
    window.addEventListener('resize', onMove);
    return () => {
      cancelAnimationFrame(raf);
      window.removeEventListener('scroll', onMove, true);
      window.removeEventListener('resize', onMove);
    };
  });

  // The drawing tool shows as a crosshair over the page.
  $effect(() => {
    canvas?.classList.toggle('wk-drawing', shapeTools.tool.kind !== 'none');
    canvas?.classList.toggle('wk-erasing', shapeTools.tool.kind === 'eraser');
  });

  function select(kind: SelectedObject['kind'], at: DocPosition, add: boolean): void {
    if (add) {
      const exists = shapeTools.multi.some((p) => samePosition(p, at));
      const base = mine(session.selectedObject) && !shapeTools.multi.length ? [session.selectedObject.at] : shapeTools.multi;
      shapeTools.multi = exists ? base.filter((p) => !samePosition(p, at)) : [...base, at];
    } else {
      shapeTools.multi = [];
    }
    session.selectedObject = { kind, at };
    session.tick++;
  }

  function deselect(): void {
    if (mine(session.selectedObject)) session.selectedObject = null;
    shapeTools.multi = [];
    shapeTools.editPoints = false;
  }

  // ---- Anchoring a drawn object ------------------------------------------

  /** The paragraph under a client point (or the closest one above it), as a position. */
  function paragraphAtPoint(x: number, y: number): { pos: DocPosition; rect: DOMRect; page: DOMRect } | undefined {
    if (!canvas) return undefined;
    let best: HTMLElement | undefined;
    for (const p of canvas.querySelectorAll<HTMLElement>('.wk-p[data-wk-block]:not([data-wk-clone])')) {
      const r = p.getBoundingClientRect();
      if (r.top <= y) best = p;
      if (r.top <= y && r.bottom >= y && r.left <= x && r.right >= x) {
        best = p;
        break;
      }
    }
    best ??= canvas.querySelector<HTMLElement>('.wk-p[data-wk-block]') ?? undefined;
    if (!best) return undefined;
    const block = Number(best.dataset.wkBlock);
    const cellAttr = best.dataset.wkCell;
    const [row = 0, col = 0] = (cellAttr ?? '').split(',').map(Number);
    const pos: DocPosition = cellAttr ? { block, cell: { row, col }, para: Number(best.dataset.wkPara ?? 0) } : { block };
    const page = (best.closest('.wk-pagebox') ?? canvas).getBoundingClientRect();
    return { pos, rect: best.getBoundingClientRect(), page };
  }

  /**
   * Offsets for a box drawn at client (x, y): horizontal from the column's
   * left edge, vertical from the anchor paragraph's top, as Word anchors a
   * drawn shape.
   */
  function anchorFor(x: number, y: number): { at: DocPosition; left: number; top: number } | undefined {
    const model = session.model;
    const hit = paragraphAtPoint(x, y);
    if (!model || !hit) return undefined;
    const marginLeft = pageGeometry(model.doc).left / TWIPS_PER_POINT;
    return {
      at: hit.pos,
      left: (x - hit.page.left) * k() - marginLeft,
      top: (y - hit.rect.top) * k(),
    };
  }

  // ---- Gestures ---------------------------------------------------------

  type Gesture =
    | { kind: 'draw'; x0: number; y0: number; x: number; y: number; points: InkPoint[] }
    | { kind: 'ink'; points: InkPoint[] }
    | { kind: 'erase' }
    | { kind: 'lasso'; x0: number; y0: number; x: number; y: number }
    | { kind: 'move'; at: DocPosition; el: HTMLElement; x0: number; y0: number; dx: number; dy: number; layout: ShapeLayout }
    | { kind: 'resize'; at: DocPosition; el: HTMLElement; handle: Handle; x0: number; y0: number; dx: number; dy: number; layout?: ShapeLayout; smartArt?: { width: number; height: number } }
    | { kind: 'rotate'; at: DocPosition; el: HTMLElement; cx: number; cy: number; angle: number; layout: ShapeLayout }
    | { kind: 'point'; at: DocPosition; index: number; points: Array<[number, number]>; layout: ShapeLayout; x0: number; y0: number };

  let gesture = $state<Gesture | null>(null);
  let erased = new Set<Element>();

  function onPointerDown(e: PointerEvent): void {
    if (e.button !== 0 || !canvas) return;
    const target = e.target as Element;
    const tool = shapeTools.tool;
    if (tool.kind === 'shape') {
      e.preventDefault();
      gesture = { kind: 'draw', x0: e.clientX, y0: e.clientY, x: e.clientX, y: e.clientY, points: [[e.clientX, e.clientY]] };
    } else if (tool.kind === 'pen') {
      e.preventDefault();
      gesture = { kind: 'ink', points: [[e.clientX, e.clientY]] };
    } else if (tool.kind === 'eraser') {
      e.preventDefault();
      erased = new Set();
      gesture = { kind: 'erase' };
      eraseAt(e.clientX, e.clientY);
    } else if (tool.kind === 'lasso') {
      e.preventDefault();
      gesture = { kind: 'lasso', x0: e.clientX, y0: e.clientY, x: e.clientX, y: e.clientY };
    } else if (tool.kind === 'link') {
      e.preventDefault();
      const el = target.closest<HTMLElement>('[data-wk-object="textBox"]');
      const to = el && parseAt(el);
      if (to) session.apply(commands.linkTextBoxCommand, { at: tool.from, to });
      shapeTools.setTool({ kind: 'none' });
    } else {
      const el = target.closest<HTMLElement>('[data-wk-object]');
      const kind = el?.dataset.wkObject;
      const at = el && parseAt(el);
      if (!el || !kind || !MY_KINDS.has(kind) || !at) {
        if (!target.closest('[data-wk-object]')) deselect();
        return;
      }
      select(kind as SelectedObject['kind'], at, e.shiftKey);
      // Clicking into a text box's text places the caret there.
      if (target.closest('.wk-txbx')) return;
      e.preventDefault();
      const layout = layoutAt(at);
      if (layout && !layout.inline && !e.shiftKey) {
        gesture = { kind: 'move', at, el, x0: e.clientX, y0: e.clientY, dx: 0, dy: 0, layout };
      }
    }
    if (gesture) (e.currentTarget as Element | null)?.setPointerCapture?.(e.pointerId);
  }

  function onPointerMove(e: PointerEvent): void {
    const g = gesture;
    if (!g) return;
    switch (g.kind) {
      case 'draw':
        gesture = { ...g, x: e.clientX, y: e.clientY, points: [...g.points, [e.clientX, e.clientY]] };
        break;
      case 'ink':
        gesture = { ...g, points: [...g.points, [e.clientX, e.clientY]] };
        break;
      case 'erase':
        eraseAt(e.clientX, e.clientY);
        break;
      case 'lasso':
        gesture = { ...g, x: e.clientX, y: e.clientY };
        break;
      case 'move': {
        const dx = e.clientX - g.x0;
        const dy = e.clientY - g.y0;
        g.el.style.translate = `${dx / session.zoom}px ${dy / session.zoom}px`;
        gesture = { ...g, dx, dy };
        break;
      }
      case 'resize': {
        const dx = e.clientX - g.x0;
        const dy = e.clientY - g.y0;
        const box = resized(g, dx, dy);
        g.el.style.width = `${box.width}pt`;
        g.el.style.height = `${box.height}pt`;
        gesture = { ...g, dx, dy };
        break;
      }
      case 'point': {
        const base = editablePoints?.points[g.index];
        if (!base) break;
        const x = Math.max(0, Math.min(g.layout.width, base[0] + (e.clientX - g.x0) * k()));
        const y = Math.max(0, Math.min(g.layout.height, base[1] + (e.clientY - g.y0) * k()));
        gesture = { ...g, points: g.points.map((p, i): [number, number] => (i === g.index ? [x, y] : p)) };
        break;
      }
      case 'rotate': {
        const angle = (Math.atan2(e.clientY - g.cy, e.clientX - g.cx) * 180) / Math.PI + 90;
        const snapped = e.shiftKey ? Math.round(angle / 15) * 15 : Math.round(angle);
        g.el.style.rotate = `${snapped - g.layout.rotation}deg`;
        gesture = { ...g, angle: snapped };
        break;
      }
    }
  }

  function resized(
    g: Extract<Gesture, { kind: 'resize' }>,
    dx: number,
    dy: number,
  ): { left: number; top: number; width: number; height: number } {
    const base = g.layout ?? { left: 0, top: 0, width: g.smartArt?.width ?? 0, height: g.smartArt?.height ?? 0 };
    let { left, top, width, height } = base;
    const px = dx * k();
    const py = dy * k();
    if (g.handle.includes('e')) width += px;
    if (g.handle.includes('s')) height += py;
    if (g.handle.includes('w')) {
      width -= px;
      left += px;
    }
    if (g.handle.startsWith('n')) {
      height -= py;
      top += py;
    }
    return { left, top, width: Math.max(1, width), height: Math.max(1, height) };
  }

  function onPointerUp(e: PointerEvent): void {
    const g = gesture;
    gesture = null;
    if (!g) return;
    switch (g.kind) {
      case 'draw':
        finishDraw(g, e.shiftKey);
        break;
      case 'ink':
        finishInk(g.points);
        break;
      case 'lasso':
        finishLasso(g);
        break;
      case 'move':
        g.el.style.translate = '';
        if (Math.abs(g.dx) + Math.abs(g.dy) >= MIN_DRAG) {
          session.apply(commands.shapeLayoutCommand, {
            at: g.at,
            layout: { left: g.layout.left + g.dx * k(), top: g.layout.top + g.dy * k() },
          });
        }
        break;
      case 'resize': {
        const box = resized(g, g.dx, g.dy);
        if (g.smartArt) session.apply(commands.smartArtSizeCommand, { at: g.at, width: box.width, height: box.height });
        else session.apply(commands.shapeLayoutCommand, { at: g.at, layout: g.layout?.inline ? { width: box.width, height: box.height } : box });
        break;
      }
      case 'rotate':
        g.el.style.rotate = '';
        session.apply(commands.shapeLayoutCommand, { at: g.at, layout: { rotation: g.angle } });
        break;
      case 'point':
        session.apply(commands.shapePointsCommand, { at: g.at, points: g.points });
        break;
      case 'erase':
        break;
    }
  }

  function finishDraw(g: Extract<Gesture, { kind: 'draw' }>, square: boolean): void {
    const tool = shapeTools.tool;
    if (tool.kind !== 'shape') return;
    const dragged = Math.abs(g.x - g.x0) >= MIN_DRAG || Math.abs(g.y - g.y0) >= MIN_DRAG;
    const freehand = tool.preset === 'freeform' || tool.preset === 'scribble';
    const xs = freehand ? g.points.map((p) => p[0]) : [g.x0, g.x];
    const ys = freehand ? g.points.map((p) => p[1]) : [g.y0, g.y];
    const x1 = Math.min(...xs);
    const y1 = Math.min(...ys);
    let width = dragged ? (Math.max(...xs) - x1) * k() : DEFAULT_SIZE;
    let height = dragged ? (Math.max(...ys) - y1) * k() : DEFAULT_SIZE;
    if (square && !freehand) width = height = Math.max(width, height);
    const anchor = anchorFor(x1, y1);
    if (!anchor) return;
    const oneD = tool.preset === 'line' || tool.preset === 'arrow' || tool.preset === 'doubleArrow' || tool.preset === 'elbowConnector' || tool.preset === 'curvedConnector';
    const at = session.apply(commands.insertShapeCommand, {
      at: anchor.at,
      options: {
        preset: tool.preset,
        left: anchor.left,
        top: anchor.top,
        width,
        height,
        ...(oneD ? { flipH: g.x < g.x0 !== g.y < g.y0 } : {}),
        ...(freehand && dragged ? { points: g.points.map(([x, y]): [number, number] => [(x - x1) * k(), (y - y1) * k()]) } : {}),
        ...(tool.textLayout ? { textLayout: tool.textLayout } : {}),
      },
    });
    shapeTools.setTool({ kind: 'none' });
    if (at) {
      const kind = tool.preset === 'textBox' ? 'textBox' : 'shape';
      select(kind, at, false);
      if (kind === 'textBox') void focusTextBox(at, 0, 0, 0);
    }
  }

  function penStroke(): ShapeStroke | undefined {
    const tool = shapeTools.tool;
    const pen = tool.kind === 'pen' ? shapeTools.pen(tool.pen) : undefined;
    if (!pen) return undefined;
    const opacity = pen.kind === 'highlighter' ? HIGHLIGHTER_OPACITY : pen.kind === 'pencil' ? PENCIL_OPACITY : undefined;
    return { color: pen.color, weight: pen.weight, ...(opacity === undefined ? {} : { opacity }) };
  }

  function finishInk(points: readonly InkPoint[]): void {
    const stroke = penStroke();
    const first = points[0];
    if (!stroke || !first || points.length < 2) return;
    const xs = points.map((p) => p[0]);
    const ys = points.map((p) => p[1]);
    const x1 = Math.min(...xs);
    const y1 = Math.min(...ys);
    const anchor = anchorFor(x1, y1);
    if (!anchor) return;
    const local = points.map(([x, y]): [number, number] => [(x - x1) * k(), (y - y1) * k()]);
    const width = Math.max(1, (Math.max(...xs) - x1) * k());
    const height = Math.max(1, (Math.max(...ys) - y1) * k());
    const shape = shapeTools.inkToShape ? recognizeInkShape(local) : undefined;
    if (shape) {
      session.apply(commands.insertShapeCommand, {
        at: anchor.at,
        options: {
          preset: shape.preset,
          left: anchor.left + shape.left,
          top: anchor.top + shape.top,
          width: Math.max(1, shape.width),
          height: Math.max(1, shape.height),
          flipH: shape.flipH,
          flipV: shape.flipV,
          fill: { type: 'none' },
          stroke,
        },
      });
      return;
    }
    // Strokes drawn inside a selected drawing canvas go into it.
    const sel = session.selectedObject;
    const model = session.model;
    const canvasEl = mine(sel) ? objectElement(sel.at) : undefined;
    const intoCanvas = model && mine(sel) && sel.kind === 'shape' && isCanvas(sel.at) && canvasEl ? canvasEl.getBoundingClientRect() : undefined;
    session.apply(commands.insertInkCommand, {
      ...(intoCanvas && mine(sel) ? { canvas: sel.at } : { at: anchor.at }),
      options: {
        left: intoCanvas ? (x1 - intoCanvas.left) * k() : anchor.left,
        top: intoCanvas ? (y1 - intoCanvas.top) * k() : anchor.top,
        width,
        height,
        points: local,
        stroke,
        fill: { type: 'none' },
      },
    });
  }

  function isCanvas(at: DocPosition): boolean {
    const model = session.model;
    if (!model) return false;
    try {
      return shapeKind(commands.shapeAt(model.doc, at)) === 'canvas';
    } catch {
      return false;
    }
  }

  function eraseAt(x: number, y: number): void {
    for (const hit of document.elementsFromPoint(x, y)) {
      const el = hit.closest<HTMLElement>('[data-wk-object="ink"]');
      const at = el && parseAt(el);
      if (!el || !at || erased.has(el)) continue;
      erased.add(el);
      session.apply(commands.deleteShapeCommand, { at });
      return;
    }
  }

  function finishLasso(g: Extract<Gesture, { kind: 'lasso' }>): void {
    if (!canvas) return;
    const l = Math.min(g.x0, g.x);
    const r = Math.max(g.x0, g.x);
    const t = Math.min(g.y0, g.y);
    const b = Math.max(g.y0, g.y);
    const picked: Array<{ kind: SelectedObject['kind']; at: DocPosition }> = [];
    for (const el of canvas.querySelectorAll<HTMLElement>('[data-wk-object]')) {
      const kind = el.dataset.wkObject ?? '';
      const at = parseAt(el);
      const rect = el.getBoundingClientRect();
      if (!MY_KINDS.has(kind) || !at) continue;
      if (rect.left >= l && rect.right <= r && rect.top >= t && rect.bottom <= b) picked.push({ kind: kind as SelectedObject['kind'], at });
    }
    shapeTools.setTool({ kind: 'none' });
    const [first] = picked;
    if (!first) return deselect();
    session.selectedObject = first;
    shapeTools.multi = picked.length > 1 ? picked.map((p) => p.at) : [];
    session.tick++;
  }

  function startResize(e: PointerEvent, handle: Handle): void {
    const sel = session.selectedObject;
    if (!mine(sel)) return;
    const el = objectElement(sel.at);
    const model = session.model;
    if (!el || !model) return;
    e.preventDefault();
    e.stopPropagation();
    let smartArt: { width: number; height: number } | undefined;
    if (sel.kind === 'smartArt') {
      const info = getSmartArt(model.doc, commands.smartArtAt(model.doc, sel.at));
      smartArt = { width: info.width, height: info.height };
    }
    const layout = sel.kind === 'smartArt' ? undefined : layoutAt(sel.at);
    gesture = { kind: 'resize', at: sel.at, el, handle, x0: e.clientX, y0: e.clientY, dx: 0, dy: 0, ...(layout ? { layout } : {}), ...(smartArt ? { smartArt } : {}) };
    (e.currentTarget as Element).setPointerCapture(e.pointerId);
  }

  function startRotate(e: PointerEvent): void {
    const sel = session.selectedObject;
    if (!mine(sel)) return;
    const el = objectElement(sel.at);
    const layout = layoutAt(sel.at);
    if (!el || !layout) return;
    e.preventDefault();
    e.stopPropagation();
    const r = el.getBoundingClientRect();
    gesture = { kind: 'rotate', at: sel.at, el, cx: r.left + r.width / 2, cy: r.top + r.height / 2, angle: layout.rotation, layout };
    (e.currentTarget as Element).setPointerCapture(e.pointerId);
  }

  /** Edit Points: the selected freeform's vertices, in points from its box. */
  const editablePoints = $derived.by(() => {
    const sel = session.selectedObject;
    const model = session.model;
    if (!shapeTools.editPoints || !mine(sel) || !model || session.tick < 0) return undefined;
    try {
      const shape = commands.shapeAt(model.doc, sel.at);
      const points = getShapePoints(shape);
      return points ? { at: sel.at, points, layout: getShapeLayout(shape) } : undefined;
    } catch {
      return undefined;
    }
  });

  function startPoint(e: PointerEvent, index: number): void {
    const ep = editablePoints;
    if (!ep) return;
    e.preventDefault();
    e.stopPropagation();
    gesture = { kind: 'point', at: ep.at, index, points: ep.points, layout: ep.layout, x0: e.clientX, y0: e.clientY };
    (e.currentTarget as Element).setPointerCapture(e.pointerId);
  }

  const livePoints = $derived(gesture?.kind === 'point' ? gesture.points : editablePoints?.points);

  // ---- Keyboard ---------------------------------------------------------

  function onKeyDown(e: KeyboardEvent): void {
    const target = e.target as Element;
    if (target.closest?.('.wk-txbx')) return onTextBoxKey(e);
    const sel = session.selectedObject;
    if (e.key === 'Escape' && shapeTools.tool.kind !== 'none') {
      shapeTools.setTool({ kind: 'none' });
      return;
    }
    if (!mine(sel)) return;
    if (e.key === 'Delete' || e.key === 'Backspace') {
      e.preventDefault();
      e.stopPropagation();
      // Delete from the last run backwards so earlier positions stay valid.
      const targets = (shapeTools.multi.length ? shapeTools.multi : [sel.at]).toSorted((a, b) => b.block - a.block || (b.inline ?? 0) - (a.inline ?? 0));
      for (const at of targets) {
        if (sel.kind === 'smartArt') session.apply(commands.deleteSmartArtCommand, { at });
        else session.apply(commands.deleteShapeCommand, { at });
      }
      deselect();
      return;
    }
    if (e.key === 'Escape') {
      e.preventDefault();
      deselect();
      return;
    }
    const step = e.altKey ? NUDGE / 4 : NUDGE * (e.shiftKey ? 10 : 1);
    const delta: Record<string, [number, number]> = { ArrowLeft: [-step, 0], ArrowRight: [step, 0], ArrowUp: [0, -step], ArrowDown: [0, step] };
    const d = delta[e.key];
    const layout = d && layoutAt(sel.at);
    if (d && layout && !layout.inline) {
      e.preventDefault();
      e.stopPropagation();
      session.apply(commands.shapeLayoutCommand, { at: sel.at, layout: { left: layout.left + d[0], top: layout.top + d[1] } });
    }
  }

  // ---- Text boxes -------------------------------------------------------

  let typingOpen = false;

  function boxOf(el: Element): { at: DocPosition; box: HTMLElement } | undefined {
    const host = el.closest<HTMLElement>('[data-wk-object]');
    const box = el.closest<HTMLElement>('.wk-txbx');
    const at = host && parseAt(host);
    return at && box ? { at, box } : undefined;
  }

  /** The caret inside a text box, as (paragraph, run, offset). */
  function caretInBox(): { para: number; run: number; offset: number } | undefined {
    const sel = window.getSelection();
    const node = sel?.anchorNode;
    if (!sel || !node) return undefined;
    const el = node instanceof Element ? node : node.parentElement;
    const run = el?.closest<HTMLElement>('[data-wk-txbx-run]');
    if (run) {
      const [para = 0, r = 0] = (run.dataset.wkTxbxRun ?? '0,0').split(',').map(Number);
      const before = node === run ? (sel.anchorOffset > 0 ? (run.textContent ?? '') : '') : (node.textContent ?? '').slice(0, sel.anchorOffset);
      // The placeholder zero-width space of an empty run is not document text.
      return { para, run: r, offset: before.replace(/​/g, '').length };
    }
    const p = el?.closest<HTMLElement>('[data-wk-txbx-para]');
    return p ? { para: Number(p.dataset.wkTxbxPara ?? 0), run: 0, offset: 0 } : undefined;
  }

  async function focusTextBox(at: DocPosition, para: number, run: number, offset: number): Promise<void> {
    await tick();
    requestAnimationFrame(() => {
      const host = objectElement(at);
      const box = host?.querySelector<HTMLElement>('.wk-txbx');
      if (!box) return;
      box.focus();
      const span = box.querySelector<HTMLElement>(`[data-wk-txbx-run="${para},${run}"]`);
      const p = box.querySelector<HTMLElement>(`[data-wk-txbx-para="${para}"]`);
      const node = span?.firstChild ?? span ?? p;
      if (!node) return;
      const range = document.createRange();
      range.setStart(node, Math.min(offset, node.textContent?.length ?? 0));
      range.collapse(true);
      const sel = window.getSelection();
      sel?.removeAllRanges();
      sel?.addRange(range);
    });
  }

  function onTextBoxKey(e: KeyboardEvent): void {
    // Keep the body's Enter / Backspace / formatting handlers out of the box.
    e.stopPropagation();
    if (e.isComposing) return;
    const found = boxOf(e.target as Element);
    const caret = caretInBox();
    if (!found || !caret) return;
    if (e.key === 'Enter') {
      e.preventDefault();
      typingOpen = false;
      session.apply(commands.shapeSplitTextCommand, { at: found.at, ...caret });
      void focusTextBox(found.at, caret.para + 1, 0, 0);
    } else if (e.key === 'Backspace' && caret.run === 0 && caret.offset === 0 && caret.para > 0) {
      e.preventDefault();
      typingOpen = false;
      const prev = found.box.querySelector(`[data-wk-txbx-para="${caret.para - 1}"]`);
      const runs = prev?.querySelectorAll('[data-wk-txbx-run]') ?? [];
      const lastRun = runs.length - 1;
      const lastLen = runs[lastRun]?.textContent?.length ?? 0;
      session.apply(commands.shapeMergeTextCommand, { at: found.at, para: caret.para });
      void focusTextBox(found.at, caret.para - 1, Math.max(0, lastRun), lastLen);
    } else if (e.key === 'Escape') {
      e.preventDefault();
      (e.target as HTMLElement).blur();
    }
  }

  /** Text typed into a text box element, without the empty-run placeholder. */
  function typedText(el: Element): string {
    return (el.textContent ?? '').replace(/\u200b/g, '');
  }

  function onInput(e: Event): void {
    const found = boxOf(e.target as Element);
    if (!found) return;
    e.stopPropagation();
    const model = session.model;
    if (!model) return;
    const paragraphs = [...found.box.querySelectorAll<HTMLElement>('[data-wk-txbx-para]')].map((p) => {
      const runs = [...p.querySelectorAll<HTMLElement>('[data-wk-txbx-run]')];
      return runs.length ? runs.map(typedText) : [typedText(p)];
    });
    // One undo step per typing burst, as on the body.
    if (!typingOpen) {
      model.beginEdit();
      typingOpen = true;
    }
    try {
      commands.shapeTextEditCommand.run(model, { at: found.at, paragraphs });
    } catch (err) {
      model.abortEdit();
      typingOpen = false;
      session.status = (err as Error).message;
      return;
    }
    model.commit();
    session.edited();
  }

  function onBeforeInputCapture(e: Event): void {
    if (boxOf(e.target as Element)) e.stopPropagation();
  }

  function onDoubleClick(e: MouseEvent): void {
    const el = (e.target as Element).closest<HTMLElement>('[data-wk-object]');
    const at = el && parseAt(el);
    const model = session.model;
    if (!el || !at || !model || el.dataset.wkObject !== 'shape') return;
    let kind;
    try {
      kind = shapeKind(commands.shapeAt(model.doc, at));
    } catch {
      return;
    }
    if (kind === 'wordArt') {
      session.openDialog('shape.wordArtText');
    } else if (kind === 'shape') {
      // Word's Add Text: the shape gets a text box and the caret goes into it.
      session.apply(commands.shapeTextCommand, { at, text: '' });
      session.selectedObject = { kind: 'textBox', at };
      void focusTextBox(at, 0, 0, 0);
    }
  }

  $effect(() => {
    const el = canvas;
    if (!el) return;
    const blurBox = (): void => {
      typingOpen = false;
    };
    el.addEventListener('pointerdown', onPointerDown, true);
    el.addEventListener('pointermove', onPointerMove);
    el.addEventListener('pointerup', onPointerUp);
    el.addEventListener('keydown', onKeyDown, true);
    el.addEventListener('input', onInput, true);
    el.addEventListener('beforeinput', onBeforeInputCapture, true);
    el.addEventListener('dblclick', onDoubleClick);
    el.addEventListener('focusout', blurBox);
    return () => {
      el.removeEventListener('pointerdown', onPointerDown, true);
      el.removeEventListener('pointermove', onPointerMove);
      el.removeEventListener('pointerup', onPointerUp);
      el.removeEventListener('keydown', onKeyDown, true);
      el.removeEventListener('input', onInput, true);
      el.removeEventListener('beforeinput', onBeforeInputCapture, true);
      el.removeEventListener('dblclick', onDoubleClick);
      el.removeEventListener('focusout', blurBox);
    };
  });

  const inkPath = $derived(
    gesture?.kind === 'ink' || (gesture?.kind === 'draw' && shapeTools.tool.kind === 'shape' && (shapeTools.tool.preset === 'freeform' || shapeTools.tool.preset === 'scribble'))
      ? gesture.points.map(([x, y], i) => `${i ? 'L' : 'M'}${x} ${y}`).join(' ')
      : '',
  );
  const inkStroke = $derived(penStroke());
  const rubber = $derived.by(() => {
    const g = gesture;
    if (g?.kind !== 'draw' && g?.kind !== 'lasso') return undefined;
    return { left: Math.min(g.x0, g.x), top: Math.min(g.y0, g.y), width: Math.abs(g.x - g.x0), height: Math.abs(g.y - g.y0) };
  });
  const selectedWrap = $derived.by(() => {
    const sel = session.selectedObject;
    const model = session.model;
    if (!mine(sel) || !model || sel.kind === 'smartArt' || session.tick < 0) return undefined;
    try {
      return getShapeWrap(commands.shapeAt(model.doc, sel.at));
    } catch {
      return undefined;
    }
  });
</script>

<div class="wk-shape-layer" aria-hidden="true">
  {#each frames as frame, i (i)}
    <div class="wk-shape-frame" class:secondary={!frame.primary} style="left:{frame.rect.left}px;top:{frame.rect.top}px;width:{frame.rect.width}px;height:{frame.rect.height}px">
      {#if frame.primary}
        {#if editablePoints && livePoints}
          {#each livePoints as [x, y], i (i)}
            <span role="presentation" class="wk-handle point" style="left:{x / k()}px;top:{y / k()}px" onpointerdown={(e) => startPoint(e, i)} onpointermove={onPointerMove} onpointerup={onPointerUp}></span>
          {/each}
        {:else}
        {#each HANDLES as handle (handle)}
          <span role="presentation" class="wk-handle {handle}" onpointerdown={(e) => startResize(e, handle)} onpointermove={onPointerMove} onpointerup={onPointerUp}></span>
        {/each}
        {#if session.selectedObject?.kind !== 'smartArt' && selectedWrap !== 'inline'}
          <span role="presentation" class="wk-handle rotate" onpointerdown={startRotate} onpointermove={onPointerMove} onpointerup={onPointerUp}></span>
        {/if}
        {/if}
      {/if}
    </div>
  {/each}
  {#if rubber}
    <div class="wk-rubber" class:lasso={gesture?.kind === 'lasso'} style="left:{rubber.left}px;top:{rubber.top}px;width:{rubber.width}px;height:{rubber.height}px"></div>
  {/if}
  {#if inkPath}
    <svg class="wk-live-ink">
      <path d={inkPath} fill="none" stroke={inkStroke ? `#${inkStroke.color}` : '#000'} stroke-opacity={inkStroke?.opacity ?? 1} stroke-width={(inkStroke?.weight ?? 1) / k()} stroke-linecap="round" stroke-linejoin="round" />
    </svg>
  {/if}
</div>

<WordArtTextDialog />
