<script lang="ts">
  /**
   * Insert ▸ Illustrations: the Shapes gallery (pick a shape, then drag on the
   * page to draw it; New Drawing Canvas at the bottom) and the SmartArt
   * gallery (List, Process, Cycle, Hierarchy).
   */
  import { SMARTART_LAYOUTS, type SmartArtLayout, type SmartArtNode } from '@office-kit/docx';
  import { commands, DEFAULT_THEME_COLORS, themePalette } from '@office-kit/docx-editor';
  import SplitButton from './SplitButton.svelte';
  import { getSession } from '../session.svelte';
  import { t, type MessageKey } from '../i18n/index.svelte';
  import ShapeGallery from '../shapes/ShapeGallery.svelte';
  import SmartArtThumb from '../shapes/SmartArtThumb.svelte';
  import { shapeTools } from '../shapes/tools.svelte';

  // Word puts Icons between Shapes and SmartArt, so each button renders on its own.
  type Props = { part: 'shapes' | 'smartArt' };
  const { part }: Props = $props();
  const session = getSession();
  const accent = $derived((session.model && session.version >= 0 ? themePalette(session.model.doc) : DEFAULT_THEME_COLORS).accent1);
  // Word's new drawing canvas: 6" × 3".
  const CANVAS_WIDTH = 432;
  const CANVAS_HEIGHT = 216;

  const LAYOUTS = Object.keys(SMARTART_LAYOUTS).filter((k): k is SmartArtLayout => Object.hasOwn(SMARTART_LAYOUTS, k));
  const CATEGORY_LABEL: Record<string, MessageKey> = {
    list: 'draw.smartArt.list',
    process: 'draw.smartArt.process',
    cycle: 'draw.smartArt.cycle',
    hierarchy: 'draw.smartArt.hierarchy',
  };
  const categories = [...new Set(LAYOUTS.map((l) => SMARTART_LAYOUTS[l].category))];

  /** Word fills a new diagram with "[Text]" placeholders in the layout's shape. */
  function placeholderNodes(layout: SmartArtLayout): SmartArtNode[] {
    const text = t('draw.smartArt.placeholder');
    if (layout === 'hierarchy') return [{ text, children: [{ text, children: [{ text }, { text }] }, { text, children: [{ text }] }] }];
    if (layout === 'verticalBulletList') return [{ text, children: [{ text }] }, { text, children: [{ text }] }];
    const count = layout === 'basicBlockList' || layout === 'basicCycle' ? 5 : 3;
    return Array.from({ length: count }, () => ({ text }));
  }

  function insertSmartArt(layout: SmartArtLayout): void {
    const at = session.apply(commands.insertSmartArtCommand, { options: { layout, nodes: placeholderNodes(layout) } });
    if (!at) return;
    session.selectedObject = { kind: 'smartArt', at };
    session.pane.left = 'smartArtText';
  }

  function insertCanvas(): void {
    const at = session.apply(commands.insertDrawingCanvasCommand, { width: CANVAS_WIDTH, height: CANVAS_HEIGHT });
    if (at) session.selectedObject = { kind: 'shape', at };
  }
</script>

{#if part === 'shapes'}
<SplitButton id="ins.shapes" size="large" icon="shapes" tip={t('draw.shapes')} label={t('draw.shapes')}>
  {#snippet menu()}
    <ShapeGallery onpick={(p) => { session.openMenu = null; shapeTools.pickPreset(p); }} />
    <hr />
    <button class="mi" role="menuitem" onclick={insertCanvas}>{t('draw.newDrawingCanvas')}</button>
  {/snippet}
</SplitButton>
{:else}
<SplitButton id="ins.smartArt" size="large" icon="smartArt" tip={t('draw.smartArt')} label={t('draw.smartArt')}>
  {#snippet menu()}
    {#each categories as category (category)}
      <div class="menu-head">{t(CATEGORY_LABEL[category] ?? 'draw.smartArt.list')}</div>
      <div class="wk-smartart-gallery">
        {#each LAYOUTS.filter((l) => SMARTART_LAYOUTS[l].category === category) as layout (layout)}
          <button class="wk-smartart-tile" role="menuitem" title={t(`smartArt.${layout}`)} aria-label={t(`smartArt.${layout}`)} onclick={() => insertSmartArt(layout)}>
            <SmartArtThumb {layout} colors={[accent]} />
          </button>
        {/each}
      </div>
    {/each}
  {/snippet}
</SplitButton>
{/if}
