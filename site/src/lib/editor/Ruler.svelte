<script lang="ts">
  /**
   * View ▸ Ruler: the horizontal ruler over the page, with the caret
   * paragraph's indent markers (first line, hanging, left, right) that can be
   * dragged, and in Print Layout the vertical ruler beside the first page.
   * Positions are measured from the rendered page, so they follow zoom and
   * the window size. Tab stops are not drawn: they are edited in Paragraph ▸
   * Tabs.
   */
  import { commands, createStyleResolver, pageGeometry, paragraphAt } from '@office-kit/docx-editor';
  import { getSession } from './session.svelte';
  import { t } from './i18n/index.svelte';

  const session = getSession();
  // Word's ruler snaps indents to 1/16 inch.
  const SNAP_TWIPS = 90;
  const TWIPS_PER_INCH = 1440;
  const EIGHTHS = 8;
  // The vertical ruler sits this far left of the page.
  const VERTICAL_GAP_PX = 20;

  let ruler = $state<HTMLDivElement | null>(null);
  /** The page's left / top edge relative to the ruler, and CSS px per twip. */
  let frame = $state({ left: 0, scale: 0 });
  /** Where the vertical ruler goes inside the scrolling surface. */
  let vertical = $state({ left: 0, top: 0 });

  const page = $derived(session.version >= 0 && session.model ? pageGeometry(session.model.doc) : undefined);
  const indent = $derived.by(() => {
    const model = session.model;
    if (session.tick < 0 || !model?.selection) return undefined;
    const para = paragraphAt(model.doc, model.selection.focus);
    if (!para) return undefined;
    const f = createStyleResolver(model.doc).paragraph(para);
    return { left: f.left ?? 0, right: f.right ?? 0, first: (f.firstLine ?? 0) - (f.hanging ?? 0) };
  });

  /** Re-measure where the page is, after every render, scroll, zoom or resize. */
  function measure(): void {
    const host = ruler?.parentElement;
    const pageEl = host?.querySelector<HTMLElement>('.wk-pagebox, .wk-canvas');
    if (!ruler || !host || !pageEl || !page) return;
    const box = pageEl.getBoundingClientRect();
    const own = ruler.getBoundingClientRect();
    const surface = host.getBoundingClientRect();
    frame = { left: box.left - own.left, scale: box.width / page.width };
    vertical = {
      left: box.left - surface.left + host.scrollLeft - VERTICAL_GAP_PX,
      top: box.top - surface.top + host.scrollTop,
    };
  }

  $effect(() => {
    // Re-run when any of these change.
    void [session.version, session.zoom, session.viewMode, page];
    const id = requestAnimationFrame(measure);
    return () => cancelAnimationFrame(id);
  });

  const ticks = $derived.by(() => {
    if (!page) return [];
    const out: { x: number; major: boolean; label?: number }[] = [];
    const total = Math.floor((page.width / TWIPS_PER_INCH) * EIGHTHS);
    for (let i = 0; i <= total; i++) {
      // Inches are numbered from the left margin, as Word numbers them.
      const twips = (i * TWIPS_PER_INCH) / EIGHTHS;
      const fromMargin = (twips - page.left) / TWIPS_PER_INCH;
      const whole = Number.isInteger(fromMargin);
      out.push({ x: twips, major: i % (EIGHTHS / 2) === 0, ...(whole && fromMargin !== 0 ? { label: Math.abs(fromMargin) } : {}) });
    }
    return out;
  });

  type Marker = 'first' | 'hanging' | 'left' | 'right';
  let drag: { marker: Marker; startX: number; start: { left: number; right: number; first: number } } | null = null;
  let preview = $state<{ left: number; right: number; first: number } | null>(null);
  const shown = $derived(preview ?? indent);

  function down(marker: Marker, e: PointerEvent): void {
    if (!indent) return;
    e.preventDefault();
    (e.currentTarget as Element).setPointerCapture(e.pointerId);
    drag = { marker, startX: e.clientX, start: indent };
  }

  function move(e: PointerEvent): void {
    if (!drag || frame.scale === 0) return;
    const delta = Math.round((e.clientX - drag.startX) / frame.scale / SNAP_TWIPS) * SNAP_TWIPS;
    const s = drag.start;
    switch (drag.marker) {
      // The first-line marker moves alone; hanging moves the left indent but keeps the first line in place.
      case 'first':
        preview = { ...s, first: s.first + delta };
        break;
      case 'hanging':
        preview = { ...s, left: s.left + delta, first: s.first - delta };
        break;
      // The left marker (the box) moves the whole paragraph.
      case 'left':
        preview = { ...s, left: s.left + delta };
        break;
      case 'right':
        preview = { ...s, right: s.right - delta };
        break;
    }
  }

  function up(): void {
    if (drag && preview) {
      const { left, right, first } = preview;
      session.apply(commands.setIndentCommand, { left, right, ...(first >= 0 ? { firstLine: first } : { hanging: -first }) });
    }
    drag = null;
    preview = null;
  }

  const px = (twips: number): number => frame.left + twips * frame.scale;
</script>

<svelte:window onresize={measure} />

{#if page}
  <div class="ruler" bind:this={ruler} role="toolbar" tabindex="-1" aria-label={t('view.ruler')} onpointermove={move} onpointerup={up}>
    <div class="track" style="left: {px(0)}px; width: {page.width * frame.scale}px">
      <div class="margin" style="width: {page.left * frame.scale}px"></div>
      <div class="margin right" style="width: {page.right * frame.scale}px"></div>
      {#each ticks as tick (tick.x)}
        <span class="tick" class:major={tick.major} style="left: {tick.x * frame.scale}px">{tick.label ?? ''}</span>
      {/each}
    </div>
    {#if shown}
      {@const origin = px(page.left)}
      <button class="marker first" style="left: {origin + (shown.left + shown.first) * frame.scale}px" onpointerdown={(e) => down('first', e)} aria-label={t('view.firstLineIndent')}></button>
      <button class="marker hanging" style="left: {origin + shown.left * frame.scale}px" onpointerdown={(e) => down('hanging', e)} aria-label={t('view.hangingIndent')}></button>
      <button class="marker left" style="left: {origin + shown.left * frame.scale}px" onpointerdown={(e) => down('left', e)} aria-label={t('view.leftIndent')}></button>
      <button class="marker hanging right" style="left: {px(page.width - page.right) - shown.right * frame.scale}px" onpointerdown={(e) => down('right', e)} aria-label={t('view.rightIndent')}></button>
    {/if}
  </div>
  {#if session.viewMode === 'print'}
    <div class="vruler" style="left: {vertical.left}px; top: {vertical.top}px; height: {page.height * frame.scale}px" aria-hidden="true">
      <div class="margin" style="height: {page.top * frame.scale}px"></div>
      <div class="margin bottom" style="height: {page.bottom * frame.scale}px"></div>
    </div>
  {/if}
{/if}

<style>
  .ruler {
    position: sticky;
    top: -24px;
    z-index: 5;
    height: 22px;
    margin: -24px -24px 8px;
    background: var(--canvas-bg);
    user-select: none;
    font-size: 9px;
    color: #555;
  }
  .track { position: absolute; top: 3px; height: 16px; background: #fff; border: 1px solid #c6c6c6; box-sizing: border-box; }
  .margin { position: absolute; top: 0; bottom: 0; left: 0; background: #d6d6d6; }
  .margin.right { left: auto; right: 0; }
  .tick { position: absolute; bottom: 2px; width: 1px; height: 3px; background: #888; transform: translateX(-0.5px); line-height: 1; text-indent: -3px; }
  .tick.major { height: 5px; }
  .tick:not(:empty) { height: auto; background: none; bottom: 3px; }
  .marker { position: absolute; width: 9px; height: 7px; padding: 0; border: 1px solid #555; background: #f4f4f4; transform: translateX(-4.5px); cursor: ew-resize; }
  .marker.first { top: 1px; clip-path: polygon(0 0, 100% 0, 50% 100%); }
  .marker.hanging { top: 12px; clip-path: polygon(50% 0, 100% 100%, 0 100%); }
  .marker.left { top: 19px; height: 4px; }
  .vruler { position: absolute; width: 14px; background: #fff; border: 1px solid #c6c6c6; box-sizing: border-box; }
  .vruler .margin { right: 0; bottom: auto; }
  .vruler .margin.bottom { top: auto; bottom: 0; }
</style>
