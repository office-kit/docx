<script lang="ts">
  /** Word's Shapes menu: Recently Used Shapes, then each category's presets. */
  import type { ShapeCategory, ShapePreset } from '@office-kit/docx';
  import { t, type MessageKey } from '../i18n/index.svelte';
  import { presetsByCategory } from './palette';
  import ShapeIcon from './ShapeIcon.svelte';
  import { shapeTools } from './tools.svelte';

  type Props = { onpick: (preset: ShapePreset) => void; showRecent?: boolean };
  const { onpick, showRecent = true }: Props = $props();
  const categories = presetsByCategory();
  const CATEGORY_LABEL: Record<ShapeCategory, MessageKey> = {
    lines: 'draw.cat.lines',
    rectangles: 'draw.cat.rectangles',
    basicShapes: 'draw.cat.basicShapes',
    blockArrows: 'draw.cat.blockArrows',
    equationShapes: 'draw.cat.equationShapes',
    flowchart: 'draw.cat.flowchart',
    starsAndBanners: 'draw.cat.starsAndBanners',
    callouts: 'draw.cat.callouts',
  };
</script>

<div class="wk-shape-gallery">
  {#if showRecent && shapeTools.recent.length}
    <div class="menu-head">{t('draw.cat.recent')}</div>
    <div class="wk-shape-grid">
      {#each shapeTools.recent as preset (preset)}
        <button class="wk-shape-tile" role="menuitem" title={t(`shape.${preset}`)} aria-label={t(`shape.${preset}`)} onclick={() => onpick(preset)}><ShapeIcon {preset} /></button>
      {/each}
    </div>
  {/if}
  {#each categories as [category, presets] (category)}
    <div class="menu-head">{t(CATEGORY_LABEL[category])}</div>
    <div class="wk-shape-grid">
      {#each presets as preset (preset)}
        <button class="wk-shape-tile" role="menuitem" title={t(`shape.${preset}`)} aria-label={t(`shape.${preset}`)} onclick={() => onpick(preset)}><ShapeIcon {preset} /></button>
      {/each}
    </div>
  {/each}
</div>
