<script lang="ts">
  /**
   * Word's Format Shape pane for the selected VML shape: Shape Options (Fill,
   * Line, Shadow, Size, Position) and Text Options (Text Box), plus Alt Text.
   * Fields commit on change, each through its own command, so each is one
   * undo step as in Word.
   */
  import {
    FILL_PATTERNS,
    getShapeAltText,
    getShapeFill,
    getShapeLayout,
    getShapeShadow,
    getShapeStroke,
    getTextBoxLayout,
    shapeKind,
    type FillPattern,
    type ShapeArrow,
    type ShapeFill,
    type ShapeLayout,
    type TextBoxLayout,
  } from '@office-kit/docx';
  import { commands } from '@office-kit/docx-editor';
  import RibbonIcon from '../RibbonIcon.svelte';
  import { getSession } from '../session.svelte';
  import { t, type MessageKey } from '../i18n/index.svelte';
  import { DASHES } from '../shapes/palette';
  import { selectedShape } from '../shapes/selection';

  const session = getSession();
  const current = $derived(selectedShape(session));
  const fill = $derived(current && getShapeFill(current.shape));
  const stroke = $derived(current && getShapeStroke(current.shape));
  const shadow = $derived(current && getShapeShadow(current.shape));
  const layout = $derived(current && getShapeLayout(current.shape));
  const kind = $derived(current && shapeKind(current.shape));
  const textLayout = $derived(current && (kind === 'textBox' || kind === 'shape') ? getTextBoxLayout(current.shape) : undefined);
  const altText = $derived(current ? getShapeAltText(current.shape) : '');

  const DEFAULT_COLOR = '4472C4';
  const FILL_TYPES = ['none', 'solid', 'gradient', 'pattern'] as const;
  // Internal margins in TextBoxLayout.inset order.
  const INSETS = ['draw.inset.left', 'draw.inset.top', 'draw.inset.right', 'draw.inset.bottom'] as const;
  const DEFAULT_SHADOW_OFFSET = 2;
  const PATTERNS = Object.keys(FILL_PATTERNS).filter((p): p is FillPattern => Object.hasOwn(FILL_PATTERNS, p));
  const ARROWS: readonly ShapeArrow[] = ['none', 'block', 'classic', 'open', 'oval', 'diamond'];
  const HREL: ReadonlyArray<ShapeLayout['horizontalRelativeTo']> = ['margin', 'page', 'column', 'character'];
  const VREL: ReadonlyArray<ShapeLayout['verticalRelativeTo']> = ['margin', 'page', 'paragraph', 'line'];
  const REL_LABEL: Record<string, MessageKey> = {
    margin: 'draw.rel.margin',
    page: 'draw.rel.page',
    column: 'draw.rel.column',
    character: 'draw.rel.character',
    paragraph: 'draw.rel.paragraph',
    line: 'draw.rel.line',
  };

  const hex = (v: string): string => v.slice(1).toUpperCase();
  const num = (e: Event): number => Number((e.currentTarget as HTMLInputElement).value);

  function setFill(type: (typeof FILL_TYPES)[number]): void {
    if (!current) return;
    const color = fill && 'color' in fill ? fill.color : DEFAULT_COLOR;
    if (type === 'none') session.apply(commands.shapeFillCommand, { at: current.at, fill: { type: 'none' } });
    if (type === 'solid') session.apply(commands.shapeFillCommand, { at: current.at, fill: { type: 'solid', color } });
    if (type === 'gradient') session.apply(commands.shapeFillCommand, { at: current.at, fill: { type: 'gradient', color, color2: 'FFFFFF' } });
    if (type === 'pattern') session.apply(commands.shapeFillCommand, { at: current.at, fill: { type: 'pattern', pattern: 'pct50', color, color2: 'FFFFFF' } });
  }

  async function pictureFill(e: Event): Promise<void> {
    const input = e.currentTarget as HTMLInputElement;
    const file = input.files?.[0];
    if (!file || !current) return;
    const bytes = new Uint8Array(await file.arrayBuffer());
    session.apply(commands.shapeFillCommand, { at: current.at, fill: { type: 'picture', bytes, contentType: file.type, tile: false } });
    input.value = '';
  }

  interface FillChange {
    color?: string;
    color2?: string;
    angle?: number;
    style?: 'linear' | 'radial';
    pattern?: FillPattern;
    opacity?: number;
  }

  /** Change one property of the current solid, gradient or pattern fill. */
  function patchFill(change: FillChange): void {
    if (!current || !fill) return;
    let next: ShapeFill | undefined;
    if (fill.type === 'solid') next = { type: 'solid', color: change.color ?? fill.color, opacity: change.opacity ?? fill.opacity ?? 1 };
    if (fill.type === 'gradient')
      next = {
        type: 'gradient',
        color: change.color ?? fill.color,
        color2: change.color2 ?? fill.color2,
        style: change.style ?? fill.style ?? 'linear',
        angle: change.angle ?? fill.angle ?? 0,
        opacity: change.opacity ?? fill.opacity ?? 1,
      };
    if (fill.type === 'pattern') next = { type: 'pattern', pattern: change.pattern ?? fill.pattern, color: change.color ?? fill.color, color2: change.color2 ?? fill.color2 };
    if (next) session.apply(commands.shapeFillCommand, { at: current.at, fill: next });
  }

  function patchStroke(change: Partial<NonNullable<typeof stroke>>): void {
    if (!current) return;
    session.apply(commands.shapeOutlineCommand, { at: current.at, stroke: { color: '000000', weight: 0.75, ...stroke, ...change } });
  }

  function patchShadow(change: Partial<NonNullable<typeof shadow>> | null): void {
    if (!current) return;
    const base = shadow ?? { color: '000000', offsetX: DEFAULT_SHADOW_OFFSET, offsetY: DEFAULT_SHADOW_OFFSET, opacity: 0.4 };
    session.apply(commands.shapeShadowCommand, { at: current.at, shadow: change === null ? null : { ...base, ...change } });
  }

  function patchLayout(change: Partial<ShapeLayout>): void {
    if (current) session.apply(commands.shapeLayoutCommand, { at: current.at, layout: change });
  }

  function patchText(change: Partial<TextBoxLayout>): void {
    if (current) session.apply(commands.textBoxLayoutCommand, { at: current.at, layout: change });
  }

  function setInset(i: number, value: number): void {
    if (!textLayout) return;
    const inset = [...textLayout.inset] as [number, number, number, number];
    inset[i] = value;
    patchText({ inset });
  }
</script>

<div class="pane-head">
  <span>{t('draw.formatShape')}</span>
  <button class="pane-close" onclick={() => (session.pane.right = null)} aria-label={t('draw.close')}><RibbonIcon name="close" size={14} /></button>
</div>
{#if current && layout}
  <div class="wk-format-pane">
    <details open>
      <summary>{t('draw.fill')}</summary>
      {#each FILL_TYPES as type (type)}
        <label><input type="radio" name="wk-fill" checked={fill?.type === type} onchange={() => setFill(type)} /> {t(`draw.fill.${type}`)}</label>
      {/each}
      <label><input type="radio" name="wk-fill" checked={fill?.type === 'picture'} disabled /> {t('draw.fill.picture')}
        <input type="file" accept="image/png,image/jpeg,image/gif" onchange={pictureFill} aria-label={t('draw.fill.picture')} /></label>
      {#if fill && 'color' in fill}
        <label class="field">{t('draw.color')} <input type="color" value="#{fill.color}" onchange={(e) => patchFill({ color: hex(e.currentTarget.value) })} /></label>
      {/if}
      {#if fill && 'color2' in fill}
        <label class="field">{t('draw.color2')} <input type="color" value="#{fill.color2}" onchange={(e) => patchFill({ color2: hex(e.currentTarget.value) })} /></label>
      {/if}
      {#if fill?.type === 'gradient'}
        <label class="field">{t('draw.angle')} <input type="number" min="0" max="359" value={fill.angle ?? 0} onchange={(e) => patchFill({ angle: num(e) })} /></label>
        <label class="field">{t('draw.gradientType')}
          <select value={fill.style ?? 'linear'} onchange={(e) => patchFill({ style: e.currentTarget.value === 'radial' ? 'radial' : 'linear' })}>
            <option value="linear">{t('draw.linear')}</option>
            <option value="radial">{t('draw.radial')}</option>
          </select></label>
      {/if}
      {#if fill?.type === 'pattern'}
        <label class="field">{t('draw.fill.pattern')}
          <select value={fill.pattern} onchange={(e) => patchFill({ pattern: PATTERNS.find((p) => p === e.currentTarget.value) ?? fill.pattern })}>
            {#each PATTERNS as p (p)}<option value={p}>{p}</option>{/each}
          </select></label>
      {/if}
      {#if fill && 'opacity' in fill}
        <label class="field">{t('draw.transparency')} <input type="number" min="0" max="100" value={Math.round((1 - (fill.opacity ?? 1)) * 100)} onchange={(e) => patchFill({ opacity: 1 - num(e) / 100 })} />%</label>
      {/if}
    </details>
    <details open>
      <summary>{t('draw.line')}</summary>
      <label><input type="radio" name="wk-line" checked={!stroke} onchange={() => current && session.apply(commands.shapeOutlineCommand, { at: current.at, stroke: null })} /> {t('draw.noLine')}</label>
      <label><input type="radio" name="wk-line" checked={!!stroke} onchange={() => patchStroke({})} /> {t('draw.solidLine')}</label>
      {#if stroke}
        <label class="field">{t('draw.color')} <input type="color" value="#{stroke.color}" onchange={(e) => patchStroke({ color: hex(e.currentTarget.value) })} /></label>
        <label class="field">{t('draw.weight')} <input type="number" min="0.25" max="1584" step="0.25" value={stroke.weight} onchange={(e) => patchStroke({ weight: num(e) })} /> pt</label>
        <label class="field">{t('draw.dashes')}
          <select value={stroke.dash ?? 'solid'} onchange={(e) => patchStroke({ dash: DASHES.find((d) => d === e.currentTarget.value) ?? 'solid' })}>
            {#each DASHES as d (d)}<option value={d}>{t(`draw.dash.${d}`)}</option>{/each}
          </select></label>
        <label class="field">{t('draw.beginArrow')}
          <select value={stroke.startArrow ?? 'none'} onchange={(e) => patchStroke({ startArrow: ARROWS.find((a) => a === e.currentTarget.value) ?? 'none' })}>
            {#each ARROWS as a (a)}<option value={a}>{t(`draw.arrow.${a}`)}</option>{/each}
          </select></label>
        <label class="field">{t('draw.endArrow')}
          <select value={stroke.endArrow ?? 'none'} onchange={(e) => patchStroke({ endArrow: ARROWS.find((a) => a === e.currentTarget.value) ?? 'none' })}>
            {#each ARROWS as a (a)}<option value={a}>{t(`draw.arrow.${a}`)}</option>{/each}
          </select></label>
      {/if}
    </details>
    <details>
      <summary>{t('draw.shadow')}</summary>
      <label><input type="checkbox" checked={!!shadow} onchange={(e) => patchShadow(e.currentTarget.checked ? {} : null)} /> {t('draw.shadow')}</label>
      {#if shadow}
        <label class="field">{t('draw.color')} <input type="color" value="#{shadow.color}" onchange={(e) => patchShadow({ color: hex(e.currentTarget.value) })} /></label>
        <label class="field">{t('draw.transparency')} <input type="number" min="0" max="100" value={Math.round((1 - (shadow.opacity ?? 1)) * 100)} onchange={(e) => patchShadow({ opacity: 1 - num(e) / 100 })} />%</label>
        <label class="field">{t('draw.offsetX')} <input type="number" step="0.5" value={shadow.offsetX} onchange={(e) => patchShadow({ offsetX: num(e) })} /> pt</label>
        <label class="field">{t('draw.offsetY')} <input type="number" step="0.5" value={shadow.offsetY} onchange={(e) => patchShadow({ offsetY: num(e) })} /> pt</label>
      {/if}
    </details>
    <details open>
      <summary>{t('draw.sizeAndPosition')}</summary>
      <label class="field">{t('draw.height')} <input type="number" min="0" step="0.5" value={+layout.height.toFixed(2)} onchange={(e) => patchLayout({ height: num(e) })} /> pt</label>
      <label class="field">{t('draw.width')} <input type="number" min="0" step="0.5" value={+layout.width.toFixed(2)} onchange={(e) => patchLayout({ width: num(e) })} /> pt</label>
      <label class="field">{t('draw.rotation')} <input type="number" min="-360" max="360" value={layout.rotation} onchange={(e) => patchLayout({ rotation: num(e) })} />°</label>
      {#if !layout.inline}
        <label class="field">{t('draw.horizontalPosition')} <input type="number" step="0.5" value={+layout.left.toFixed(2)} onchange={(e) => patchLayout({ left: num(e) })} /> pt</label>
        <label class="field">{t('draw.relativeTo')}
          <select value={layout.horizontalRelativeTo} onchange={(e) => patchLayout({ horizontalRelativeTo: HREL.find((r) => r === e.currentTarget.value) ?? 'margin' })}>
            {#each HREL as r (r)}<option value={r}>{t(REL_LABEL[r] ?? 'draw.rel.margin')}</option>{/each}
          </select></label>
        <label class="field">{t('draw.verticalPosition')} <input type="number" step="0.5" value={+layout.top.toFixed(2)} onchange={(e) => patchLayout({ top: num(e) })} /> pt</label>
        <label class="field">{t('draw.relativeTo')}
          <select value={layout.verticalRelativeTo} onchange={(e) => patchLayout({ verticalRelativeTo: VREL.find((r) => r === e.currentTarget.value) ?? 'paragraph' })}>
            {#each VREL as r (r)}<option value={r}>{t(REL_LABEL[r] ?? 'draw.rel.margin')}</option>{/each}
          </select></label>
      {/if}
    </details>
    {#if textLayout}
      <details open>
        <summary>{t('draw.textBox')}</summary>
        <label class="field">{t('draw.verticalAlignment')}
          <select value={textLayout.anchor} onchange={(e) => patchText({ anchor: e.currentTarget.value === 'middle' ? 'middle' : e.currentTarget.value === 'bottom' ? 'bottom' : 'top' })}>
            <option value="top">{t('draw.alignTop')}</option>
            <option value="middle">{t('draw.alignMiddle')}</option>
            <option value="bottom">{t('draw.alignBottom')}</option>
          </select></label>
        <label class="field">{t('draw.textDirection')}
          <select value={textLayout.direction} onchange={(e) => patchText({ direction: e.currentTarget.value === 'vertical' ? 'vertical' : e.currentTarget.value === 'vertical270' ? 'vertical270' : 'horizontal' })}>
            <option value="horizontal">{t('draw.dir.horizontal')}</option>
            <option value="vertical">{t('draw.dir.rotate90')}</option>
            <option value="vertical270">{t('draw.dir.rotate270')}</option>
          </select></label>
        <label><input type="checkbox" checked={textLayout.autoFit} onchange={(e) => patchText({ autoFit: e.currentTarget.checked })} /> {t('draw.resizeToFit')}</label>
        {#each INSETS as key, i (key)}
          <label class="field">{t(key)} <input type="number" min="0" step="0.5" value={textLayout.inset[i]} onchange={(e) => setInset(i, num(e))} /> pt</label>
        {/each}
      </details>
    {/if}
    <details>
      <summary>{t('draw.altText')}</summary>
      <textarea rows="4" value={altText} aria-label={t('draw.altText')} onchange={(e) => current && session.apply(commands.shapeAltTextCommand, { at: current.at, alt: e.currentTarget.value })}></textarea>
    </details>
  </div>
{:else}
  <p class="pane-note">{t('draw.selectShape')}</p>
{/if}

<style>
  .wk-format-pane {
    display: grid;
    gap: 4px;
    padding: 0 10px 10px;
    font-size: 12px;
  }
  .wk-format-pane details {
    border-bottom: 1px solid #ddd;
    padding: 4px 0;
  }
  .wk-format-pane summary {
    font-weight: 600;
    cursor: pointer;
  }
  .wk-format-pane label {
    display: flex;
    align-items: center;
    gap: 6px;
    margin: 3px 0;
  }
  .wk-format-pane input[type='number'] {
    width: 64px;
  }
  .wk-format-pane textarea {
    box-sizing: border-box;
    width: 100%;
  }
</style>
