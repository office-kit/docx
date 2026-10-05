<script lang="ts">
  /**
   * Picture and chart selection on the canvas, drawn over the page in
   * viewport coordinates: Word's frame with eight sizing handles and a rotate
   * handle, crop handles in crop mode, and wrap-point handles while editing
   * wrap points. It also lays out floating objects after every render and
   * takes image drops / pastes into the canvas.
   *
   * Only `[data-wk-object="picture"|"chart"]` is handled here; other objects
   * (shapes, SmartArt) have their own handler.
   */
  import {
    caretAt,
    commands,
    EMU_PER_PX,
    floatFrameStart,
    layoutFloatingObjects,
    positionFromDom,
    readFloat,
    type DocPosition,
  } from '@office-kit/docx-editor';
  import { getSession } from '../session.svelte';
  import { drawingElement, pageBoxOf, selected } from './state';
  import { pictureModes, resetPictureModes } from './ui.svelte';
  import { insertPictureFiles, pixelColorAt } from './insert';

  type Props = { canvas: HTMLElement | null };
  const { canvas }: Props = $props();
  const session = getSession();

  type Handle = 'nw' | 'n' | 'ne' | 'e' | 'se' | 's' | 'sw' | 'w';
  const HANDLES: readonly Handle[] = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'];
  const POLYGON_UNITS = 21600;
  // Pointer travel (px) before a press becomes a drag.
  const DRAG_THRESHOLD = 3;
  const ROTATE_SNAP_DEG = 15;
  const OUR_OBJECTS = '[data-wk-object="picture"],[data-wk-object="chart"]';

  interface Box { x: number; y: number; w: number; h: number }
  type Drag =
    | { kind: 'move'; start: Box; px: number; py: number; dx: number; dy: number }
    | { kind: 'resize' | 'crop'; handle: Handle; start: Box; px: number; py: number; dx: number; dy: number }
    | { kind: 'rotate'; start: Box; angle: number }
    | { kind: 'point'; point: number; start: Box; px: number; py: number; dx: number; dy: number };

  let box = $state<Box | null>(null);
  let drag = $state<Drag | null>(null);
  const current = $derived(selected(session));

  /** The canvas' zoom: CSS px of the page per CSS px of the viewport. */
  function zoomScale(): number {
    return canvas ? canvas.getBoundingClientRect().width / canvas.offsetWidth || 1 : 1;
  }

  function relayout(): void {
    const model = session.model;
    if (!canvas || !model) return;
    layoutFloatingObjects(canvas, pageBoxOf(model.doc));
    measure();
  }

  function measure(): void {
    const s = current;
    const el = s ? drawingElement(s.index) : null;
    if (!el) {
      box = null;
      // The object went away (deleted, undone): drop the stale selection.
      if (session.selectedObject && (session.selectedObject.kind === 'picture' || session.selectedObject.kind === 'chart') && session.tick >= 0 && !s) {
        session.selectedObject = null;
        resetPictureModes();
      }
      return;
    }
    const r = el.getBoundingClientRect();
    box = { x: r.left, y: r.top, w: r.width, h: r.height };
  }

  // Re-render → lay floats out again (the canvas replaces its children).
  $effect(() => {
    if (!canvas) return;
    const observer = new MutationObserver(relayout);
    observer.observe(canvas, { childList: true });
    relayout();
    return () => observer.disconnect();
  });

  // Typing reflows text; selection and zoom changes move the frame.
  $effect(() => {
    void session.tick;
    void session.zoom;
    void session.selectedObject;
    queueMicrotask(relayout);
  });

  $effect(() => {
    const onScroll = (): void => measure();
    window.addEventListener('scroll', onScroll, true);
    window.addEventListener('resize', onScroll);
    return () => {
      window.removeEventListener('scroll', onScroll, true);
      window.removeEventListener('resize', onScroll);
    };
  });

  function objectPosition(el: Element): DocPosition | null {
    const raw = el.getAttribute('data-wk-at');
    // Written by our renderer; JSON.parse only fails on a renderer bug.
    return raw ? (JSON.parse(raw) as DocPosition) : null;
  }

  /** Image decoding failures (unsupported or corrupt files) go to the status bar. */
  function reportError(err: unknown): void {
    session.status = (err as Error).message;
  }

  async function pickTransparent(el: HTMLElement, e: PointerEvent): Promise<void> {
    const s = current;
    if (!s) return;
    const r = el.getBoundingClientRect();
    const color = await pixelColorAt(s.info, (e.clientX - r.left) / r.width, (e.clientY - r.top) / r.height);
    pictureModes.pickingTransparent = false;
    if (!color) return;
    session.apply(commands.pictureAdjustCommand, {
      index: s.index,
      adjustments: { ...s.info.picture?.adjustments, transparentColor: color },
    });
  }

  function onCanvasPointerDown(e: PointerEvent): void {
    const target = e.target as Element;
    const obj = target.closest<HTMLElement>(OUR_OBJECTS);
    if (!obj) {
      // Clicks on other objects belong to their own handler.
      if (!target.closest('[data-wk-object]') && session.selectedObject && (session.selectedObject.kind === 'picture' || session.selectedObject.kind === 'chart')) {
        session.selectedObject = null;
        resetPictureModes();
      }
      return;
    }
    e.preventDefault();
    const href = obj.dataset.wkHref;
    if (href && (e.metaKey || e.ctrlKey)) {
      window.open(href, '_blank', 'noopener');
      return;
    }
    const at = objectPosition(obj);
    if (!at) return;
    const kind = obj.dataset.wkObject === 'chart' ? 'chart' : 'picture';
    const same = session.selectedObject && JSON.stringify(session.selectedObject.at) === JSON.stringify(at);
    if (!same) {
      resetPictureModes();
      session.selectedObject = { kind, at };
    } else if (pictureModes.pickingTransparent) {
      pickTransparent(obj, e).catch(reportError);
      return;
    }
    // Floating objects move with the pointer.
    if (obj.dataset.wkFloat) {
      const r = obj.getBoundingClientRect();
      startDrag({ kind: 'move', start: { x: r.left, y: r.top, w: r.width, h: r.height }, px: e.clientX, py: e.clientY, dx: 0, dy: 0 }, e);
    }
  }

  function startDrag(d: Drag, e: PointerEvent): void {
    drag = d;
    window.addEventListener('pointermove', onPointerMove);
    window.addEventListener('pointerup', onPointerUp, { once: true });
    e.stopPropagation();
  }

  function onHandleDown(handle: Handle, e: PointerEvent): void {
    if (!box) return;
    e.preventDefault();
    startDrag({ kind: pictureModes.cropping ? 'crop' : 'resize', handle, start: { ...box }, px: e.clientX, py: e.clientY, dx: 0, dy: 0 }, e);
  }

  function onRotateDown(e: PointerEvent): void {
    if (!box) return;
    e.preventDefault();
    startDrag({ kind: 'rotate', start: { ...box }, angle: current?.info.rotation ?? 0 }, e);
  }

  function onPointDown(point: number, e: PointerEvent): void {
    if (!box) return;
    e.preventDefault();
    startDrag({ kind: 'point', point, start: { ...box }, px: e.clientX, py: e.clientY, dx: 0, dy: 0 }, e);
  }

  function onPointerMove(e: PointerEvent): void {
    const d = drag;
    if (!d) return;
    if (d.kind === 'rotate') {
      const cx = d.start.x + d.start.w / 2;
      const cy = d.start.y + d.start.h / 2;
      let deg = (Math.atan2(e.clientY - cy, e.clientX - cx) * 180) / Math.PI + 90;
      if (e.shiftKey) deg = Math.round(deg / ROTATE_SNAP_DEG) * ROTATE_SNAP_DEG;
      drag = { ...d, angle: ((deg % 360) + 360) % 360 };
      return;
    }
    drag = { ...d, dx: e.clientX - d.px, dy: e.clientY - d.py };
  }

  /** The box a resize / crop drag produces, keeping the aspect on corners when locked. */
  function resizedBox(d: Extract<Drag, { handle: Handle }>, keepAspect: boolean): Box {
    const { start, handle } = d;
    let { dx, dy } = d;
    if (!handle.includes('e') && !handle.includes('w')) dx = 0;
    if (!handle.includes('n') && !handle.includes('s')) dy = 0;
    let w = start.w + (handle.includes('e') ? dx : handle.includes('w') ? -dx : 0);
    let h = start.h + (handle.includes('s') ? dy : handle.includes('n') ? -dy : 0);
    const corner = handle.length === 2;
    if (corner && keepAspect) {
      const k = Math.max(w / start.w, h / start.h);
      w = start.w * k;
      h = start.h * k;
    }
    w = Math.max(4, w);
    h = Math.max(4, h);
    return {
      x: handle.includes('w') ? start.x + start.w - w : start.x,
      y: handle.includes('n') ? start.y + start.h - h : start.y,
      w,
      h,
    };
  }

  const ghost = $derived.by((): Box | null => {
    const d = drag;
    if (!d || d.kind === 'rotate' || d.kind === 'point') return null;
    if (d.kind === 'move') return { ...d.start, x: d.start.x + d.dx, y: d.start.y + d.dy };
    const lock = (current?.info.lockAspect ?? false) !== (d.kind === 'resize' && d.dx !== 0 && false);
    return resizedBox(d, d.kind === 'resize' && lock);
  });

  /** Offsets (EMU) that put a floating object's top-left at viewport point (x, y). */
  function offsetsFor(el: HTMLElement, x: number, y: number) {
    const model = session.model;
    const data = readFloat(el);
    if (!canvas || !model || !data) return undefined;
    const scale = zoomScale();
    const c = canvas.getBoundingClientRect();
    const frame = floatFrameStart(canvas, el, pageBoxOf(model.doc));
    const cx = (x - c.left) / scale - canvas.clientLeft;
    const cy = (y - c.top) / scale - canvas.clientTop;
    return {
      horizontal: { relativeTo: data.h.relativeTo, offsetEmu: Math.round((cx - frame.x) * EMU_PER_PX) },
      vertical: { relativeTo: data.v.relativeTo, offsetEmu: Math.round((cy - frame.y) * EMU_PER_PX) },
    };
  }

  function onPointerUp(): void {
    window.removeEventListener('pointermove', onPointerMove);
    const d = drag;
    drag = null;
    const s = current;
    const el = s ? drawingElement(s.index) : null;
    if (!d || !s || !el) return;
    const scale = zoomScale();
    if (d.kind === 'rotate') {
      session.apply(commands.pictureTransformCommand, { index: s.index, rotation: Math.round(d.angle) });
      return;
    }
    if (Math.abs(d.dx) < DRAG_THRESHOLD && Math.abs(d.dy) < DRAG_THRESHOLD) return;
    if (d.kind === 'move') {
      const offsets = offsetsFor(el, d.start.x + d.dx, d.start.y + d.dy);
      if (offsets) session.apply(commands.drawingPositionCommand, { index: s.index, position: offsets });
      return;
    }
    if (d.kind === 'point') {
      commitWrapPoint(d, s.index);
      return;
    }
    const next = resizedBox(d, d.kind === 'resize' && s.info.lockAspect);
    const cxEmu = (next.w / scale) * EMU_PER_PX;
    const cyEmu = (next.h / scale) * EMU_PER_PX;
    if (d.kind === 'resize') {
      session.apply(commands.resizeImageCommand, { index: s.index, cxEmu, cyEmu });
    } else {
      const crop = s.info.picture?.crop ?? { left: 0, top: 0, right: 0, bottom: 0 };
      // The full image's size on screen: the visible part is (100 - crops) % of it.
      const fullW = d.start.w / Math.max(0.01, 1 - (crop.left + crop.right) / 100);
      const fullH = d.start.h / Math.max(0.01, 1 - (crop.top + crop.bottom) / 100);
      const dl = next.x - d.start.x;
      const dt = next.y - d.start.y;
      const dr = d.start.x + d.start.w - (next.x + next.w);
      const db = d.start.y + d.start.h - (next.y + next.h);
      session.apply(commands.cropPictureCommand, {
        index: s.index,
        crop: {
          left: crop.left + (dl / fullW) * 100,
          top: crop.top + (dt / fullH) * 100,
          right: crop.right + (dr / fullW) * 100,
          bottom: crop.bottom + (db / fullH) * 100,
        },
        size: { cxEmu, cyEmu },
      });
    }
    // Dragging a left / top handle keeps the opposite edge where it was.
    if (el.dataset.wkFloat && (next.x !== d.start.x || next.y !== d.start.y)) {
      const moved = drawingElement(s.index);
      const offsets = moved ? offsetsFor(moved, next.x, next.y) : undefined;
      if (offsets) session.apply(commands.drawingPositionCommand, { index: s.index, position: offsets });
    }
  }

  const polygon = $derived.by(() => {
    const a = current?.info.anchor;
    if (!pictureModes.editingWrapPoints || !a) return [];
    return a.wrapPolygon ?? [
      { x: 0, y: 0 },
      { x: 0, y: POLYGON_UNITS },
      { x: POLYGON_UNITS, y: POLYGON_UNITS },
      { x: POLYGON_UNITS, y: 0 },
      { x: 0, y: 0 },
    ];
  });

  function commitWrapPoint(d: Extract<Drag, { kind: 'point' }>, index: number): void {
    const points = polygon.map((p, i) =>
      i === d.point || (d.point === 0 && i === polygon.length - 1)
        ? { x: p.x + (d.dx / d.start.w) * POLYGON_UNITS, y: p.y + (d.dy / d.start.h) * POLYGON_UNITS }
        : p,
    );
    session.apply(commands.drawingAnchorOptionsCommand, { index, options: { wrapPolygon: points } });
  }

  function handleStyle(h: Handle, b: Box): string {
    const x = h.includes('w') ? 0 : h.includes('e') ? b.w : b.w / 2;
    const y = h.includes('n') ? 0 : h.includes('s') ? b.h : b.h / 2;
    return `left:${x}px;top:${y}px`;
  }

  function onKeydown(e: KeyboardEvent): void {
    const s = current;
    if (!s) return;
    const target = e.target as Element;
    if (target.closest('input,textarea,select,[role="dialog"]')) return;
    if (e.key === 'Escape') {
      e.preventDefault();
      if (pictureModes.cropping || pictureModes.pickingTransparent || pictureModes.editingWrapPoints) resetPictureModes();
      else session.selectedObject = null;
      return;
    }
    if (e.key === 'Delete' || e.key === 'Backspace') {
      e.preventDefault();
      session.apply(commands.drawingDeleteCommand, { index: s.index });
      session.selectedObject = null;
      resetPictureModes();
      return;
    }
    const nudge: Record<string, [number, number]> = { ArrowLeft: [-1, 0], ArrowRight: [1, 0], ArrowUp: [0, -1], ArrowDown: [0, 1] };
    const step = nudge[e.key];
    const el = drawingElement(s.index);
    if (step && el?.dataset.wkFloat && box) {
      e.preventDefault();
      const offsets = offsetsFor(el, box.x + step[0], box.y + step[1]);
      if (offsets) session.apply(commands.drawingPositionCommand, { index: s.index, position: offsets });
    }
  }

  function imageFiles(data: DataTransfer | null): File[] {
    return data ? Array.from(data.files).filter((f) => f.type.startsWith('image/')) : [];
  }

  function onDragOver(e: DragEvent): void {
    if (e.dataTransfer && Array.from(e.dataTransfer.types).includes('Files')) e.preventDefault();
  }

  function caretFromPoint(x: number, y: number): DocPosition | null {
    const p = document.caretPositionFromPoint(x, y);
    return p ? positionFromDom(p.offsetNode, p.offset) : null;
  }

  function onDrop(e: DragEvent): void {
    const files = imageFiles(e.dataTransfer);
    if (!files.length) return;
    e.preventDefault();
    const pos = caretFromPoint(e.clientX, e.clientY);
    if (pos) session.model?.setSelection(caretAt(pos));
    insertPictureFiles(session, files).catch(reportError);
  }

  function onPaste(e: ClipboardEvent): void {
    const files = imageFiles(e.clipboardData);
    if (!files.length) return;
    // Ahead of the canvas' own text paste, which would insert nothing.
    e.preventDefault();
    e.stopImmediatePropagation();
    insertPictureFiles(session, files).catch(reportError);
  }

  $effect(() => {
    const el = canvas;
    if (!el) return;
    el.addEventListener('pointerdown', onCanvasPointerDown, true);
    el.addEventListener('dragover', onDragOver);
    el.addEventListener('drop', onDrop);
    el.addEventListener('paste', onPaste, true);
    document.addEventListener('keydown', onKeydown, true);
    return () => {
      el.removeEventListener('pointerdown', onCanvasPointerDown, true);
      el.removeEventListener('dragover', onDragOver);
      el.removeEventListener('drop', onDrop);
      el.removeEventListener('paste', onPaste, true);
      document.removeEventListener('keydown', onKeydown, true);
    };
  });

  const rotation = $derived(drag?.kind === 'rotate' ? drag.angle : (current?.info.rotation ?? 0));
  const isPicture = $derived(current?.info.kind === 'picture');
  const fullImage = $derived.by((): Box | null => {
    const crop = current?.info.picture?.crop;
    if (!box || !crop || !pictureModes.cropping) return null;
    const fw = box.w / Math.max(0.01, 1 - (crop.left + crop.right) / 100);
    const fh = box.h / Math.max(0.01, 1 - (crop.top + crop.bottom) / 100);
    return { x: (-crop.left / 100) * fw, y: (-crop.top / 100) * fh, w: fw, h: fh };
  });
</script>

{#if box && current}
  <div
    class="wk-objsel"
    class:cropping={pictureModes.cropping}
    style="left:{box.x}px;top:{box.y}px;width:{box.w}px;height:{box.h}px;transform:rotate({rotation}deg)"
    aria-hidden="true"
  >
    {#if fullImage}<div class="full" style="left:{fullImage.x}px;top:{fullImage.y}px;width:{fullImage.w}px;height:{fullImage.h}px"></div>{/if}
    {#each HANDLES as h (h)}
      <span class="handle {h}" role="button" tabindex="-1" aria-label={h} class:crop={pictureModes.cropping} style={handleStyle(h, box)} onpointerdown={(e) => onHandleDown(h, e)}></span>
    {/each}
    {#if isPicture && !pictureModes.cropping}
      <span class="rotate" role="button" tabindex="-1" aria-label="rotate" style="left:{box.w / 2}px" onpointerdown={onRotateDown}></span>
    {/if}
    {#if polygon.length}
      <svg class="wrap-poly" width={box.w} height={box.h} overflow="visible">
        <polygon points={polygon.map((p) => `${(p.x / POLYGON_UNITS) * box!.w},${(p.y / POLYGON_UNITS) * box!.h}`).join(' ')} />
      </svg>
      {#each polygon.slice(0, -1) as p, i (i)}
        <span class="point" role="button" tabindex="-1" aria-label="wrap point" style="left:{(p.x / POLYGON_UNITS) * box.w}px;top:{(p.y / POLYGON_UNITS) * box.h}px" onpointerdown={(e) => onPointDown(i, e)}></span>
      {/each}
    {/if}
  </div>
{/if}
{#if ghost}
  <div class="wk-objghost" style="left:{ghost.x}px;top:{ghost.y}px;width:{ghost.w}px;height:{ghost.h}px" aria-hidden="true"></div>
{/if}

<style>
  /* Word for Mac: a thin frame, round white handles, a circular-arrow rotate handle. */
  .wk-objsel {
    position: fixed;
    z-index: 30;
    border: 1px solid #7a7a7a;
    box-sizing: border-box;
    pointer-events: none;
  }
  .handle {
    position: absolute;
    width: 9px;
    height: 9px;
    margin: -5px 0 0 -5px;
    border: 1px solid #7a7a7a;
    border-radius: 50%;
    background: #fff;
    pointer-events: auto;
  }
  .handle.nw, .handle.se { cursor: nwse-resize; }
  .handle.ne, .handle.sw { cursor: nesw-resize; }
  .handle.n, .handle.s { cursor: ns-resize; }
  .handle.e, .handle.w { cursor: ew-resize; }
  /* Crop handles: black bars along the edges and L-shaped corners. */
  .handle.crop {
    border: none;
    border-radius: 0;
    background: #000;
  }
  .handle.crop.n, .handle.crop.s { width: 16px; height: 4px; margin: -2px 0 0 -8px; }
  .handle.crop.e, .handle.crop.w { width: 4px; height: 16px; margin: -8px 0 0 -2px; }
  .handle.crop.nw, .handle.crop.ne, .handle.crop.se, .handle.crop.sw { width: 10px; height: 10px; margin: -5px 0 0 -5px; }
  .rotate {
    position: absolute;
    top: -26px;
    width: 14px;
    height: 14px;
    margin-left: -7px;
    border: 2px solid #7a7a7a;
    border-top-color: transparent;
    border-radius: 50%;
    background: #fff;
    cursor: grab;
    pointer-events: auto;
  }
  .full {
    position: absolute;
    border: 1px dashed #7a7a7a;
    background: rgba(255, 255, 255, 0.35);
  }
  .wrap-poly {
    position: absolute;
    inset: 0;
  }
  .wrap-poly polygon {
    fill: none;
    stroke: #c00000;
    stroke-dasharray: 3 2;
  }
  .point {
    position: absolute;
    width: 7px;
    height: 7px;
    margin: -4px 0 0 -4px;
    background: #000;
    cursor: move;
    pointer-events: auto;
  }
  .wk-objghost {
    position: fixed;
    z-index: 31;
    border: 1px dashed #365695;
    background: rgba(54, 86, 149, 0.08);
    pointer-events: none;
  }
</style>
