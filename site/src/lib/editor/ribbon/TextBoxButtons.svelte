<script lang="ts">
  /**
   * Insert ▸ Text: the Text Box menu (built-in text boxes, Draw Text Box,
   * Draw Vertical Text Box) and the WordArt gallery.
   */
  import type { AddShapeOptions } from '@office-kit/docx';
  import { commands, themeColors } from '@office-kit/docx-editor';
  import SplitButton from './SplitButton.svelte';
  import { getSession } from '../session.svelte';
  import { t } from '../i18n/index.svelte';
  import { fillCss, tint, wordArtStyles } from '../shapes/palette';
  import { shapeTools } from '../shapes/tools.svelte';

  const session = getSession();
  const theme = $derived(session.model && session.version >= 0 ? themeColors(session.model.doc) : {});
  const styles = $derived(wordArtStyles(theme));

  // Built-in text boxes: 3" × 1.5" floating at the top of the margin box, as Word's presets are.
  const PRESET_WIDTH = 216;
  const PRESET_HEIGHT = 108;
  const SIDEBAR_WIDTH = 144;
  const SIDEBAR_HEIGHT = 432;
  const WORDART_WIDTH = 288;
  const WORDART_HEIGHT = 72;

  type Preset = { id: 'simple' | 'quote' | 'sidebar'; options: () => AddShapeOptions };
  const PRESETS: readonly Preset[] = [
    {
      id: 'simple',
      options: () => ({ preset: 'textBox', width: PRESET_WIDTH, height: PRESET_HEIGHT, left: 0, top: 0, wrap: 'square', text: t('draw.textBoxPlaceholder') }),
    },
    {
      id: 'quote',
      options: () => ({
        preset: 'textBox',
        width: PRESET_WIDTH * 2,
        height: PRESET_HEIGHT / 2,
        left: 0,
        top: 0,
        wrap: 'topAndBottom',
        fill: { type: 'none' },
        stroke: { color: theme.accent1 ?? '4472C4', weight: 1 },
        text: t('draw.quotePlaceholder'),
        textLayout: { anchor: 'middle' },
      }),
    },
    {
      id: 'sidebar',
      options: () => ({
        preset: 'textBox',
        width: SIDEBAR_WIDTH,
        height: SIDEBAR_HEIGHT,
        left: 0,
        top: 0,
        horizontalRelativeTo: 'margin',
        verticalRelativeTo: 'margin',
        wrap: 'square',
        fill: { type: 'solid', color: tint(theme.accent1 ?? '4472C4', 0.8) },
        stroke: null,
        text: t('draw.sidebarPlaceholder'),
      }),
    },
  ];
  const PRESET_LABEL = { simple: 'draw.simpleTextBox', quote: 'draw.quoteTextBox', sidebar: 'draw.sidebarTextBox' } as const;

  function insertPreset(p: Preset): void {
    const at = session.apply(commands.insertShapeCommand, { options: p.options() });
    if (at) session.selectedObject = { kind: 'textBox', at };
  }

  function insertWordArt(index: number): void {
    const style = styles[index];
    if (!style) return;
    const at = session.apply(commands.insertWordArtCommand, {
      options: {
        text: t('draw.wordArtPlaceholder'),
        width: WORDART_WIDTH,
        height: WORDART_HEIGHT,
        left: 0,
        top: 0,
        wrap: 'square',
        fill: style.fill,
        stroke: style.stroke,
        shadow: style.shadow,
      },
    });
    if (at) {
      session.selectedObject = { kind: 'shape', at };
      session.openDialog('shape.wordArtText');
    }
  }
</script>

<SplitButton id="ins.textBox" size="large" icon="textBox" tip={t('draw.textBox')} label={t('draw.textBox')}>
  {#snippet menu()}
    <div class="menu-head">{t('draw.builtIn')}</div>
    {#each PRESETS as p (p.id)}
      <button class="mi" role="menuitem" onclick={() => insertPreset(p)}>{t(PRESET_LABEL[p.id])}</button>
    {/each}
    <hr />
    <button class="mi" role="menuitem" onclick={() => { session.openMenu = null; shapeTools.pickPreset('textBox'); }}>{t('draw.drawTextBox')}</button>
    <button class="mi" role="menuitem" onclick={() => { session.openMenu = null; shapeTools.pickPreset('textBox', { direction: 'vertical' }); }}>{t('draw.drawVerticalTextBox')}</button>
  {/snippet}
</SplitButton>
<SplitButton id="ins.wordArt" size="large" icon="wordArt" tip={t('draw.wordArt')} label={t('draw.wordArt')}>
  {#snippet menu()}
    <div class="wk-style-menu wk-wordart-menu">
      {#each styles as style, i (i)}
        <button class="wk-style-tile wk-wordart-tile" role="menuitem" title={t('draw.wordArt')} onclick={() => insertWordArt(i)}>
          <span
            style="color:{style.fill.type === 'none' ? 'transparent' : ''};background:{fillCss(style.fill)};background-clip:text;-webkit-background-clip:text;-webkit-text-fill-color:{style.fill.type === 'solid' ? `#${style.fill.color}` : 'transparent'};-webkit-text-stroke:{style.stroke ? `1px #${style.stroke.color}` : '0'};text-shadow:{style.shadow ? '2px 2px 2px rgb(0 0 0 / 35%)' : 'none'}"
            >A</span
          >
        </button>
      {/each}
    </div>
  {/snippet}
</SplitButton>
