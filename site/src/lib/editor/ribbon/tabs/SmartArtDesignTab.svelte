<script lang="ts">
  /**
   * SmartArt Design (contextual, while a SmartArt graphic is selected): Create
   * Graphic (Add Shape, Promote, Demote, Move Up / Down, Text Pane), Layouts,
   * Change Colors, SmartArt Styles and Reset Graphic.
   */
  import {
    getSmartArt,
    SMARTART_COLORS,
    SMARTART_LAYOUTS,
    SMARTART_STYLES,
    type SmartArtColors,
    type SmartArtLayout,
    type SmartArtStyle,
  } from '@office-kit/docx';
  import { applyOutlineOp, commands, outlineRows, outlineToNodes, themeColors, type OutlineOp } from '@office-kit/docx-editor';
  import Button from '../Button.svelte';
  import Group from '../Group.svelte';
  import SplitButton from '../SplitButton.svelte';
  import { getSession } from '../../session.svelte';
  import { t, type MessageKey } from '../../i18n/index.svelte';
  import SmartArtThumb from '../../shapes/SmartArtThumb.svelte';
  import { selectedSmartArt } from '../../shapes/selection';
  import { shapeTools } from '../../shapes/tools.svelte';

  const session = getSession();
  const current = $derived(selectedSmartArt(session));
  const info = $derived(current && session.model ? getSmartArt(session.model.doc, current.ref) : undefined);
  const theme = $derived(session.model && session.version >= 0 ? themeColors(session.model.doc) : {});

  const LAYOUTS = Object.keys(SMARTART_LAYOUTS).filter((k): k is SmartArtLayout => Object.hasOwn(SMARTART_LAYOUTS, k));
  const COLORS = Object.keys(SMARTART_COLORS).filter((k): k is SmartArtColors => Object.hasOwn(SMARTART_COLORS, k));
  const STYLES = Object.keys(SMARTART_STYLES).filter((k): k is SmartArtStyle => Object.hasOwn(SMARTART_STYLES, k));
  const DEFAULT_COLORS: SmartArtColors = 'accent1_2';
  const DEFAULT_STYLE: SmartArtStyle = 'simple1';
  // Change Colors groups its rows under these headings, in Word's order.
  const COLOR_HEADINGS: ReadonlyArray<[string, MessageKey]> = [
    ['colorful', 'draw.smartArt.colorful'],
    ['accent1', 'draw.smartArt.accent1'],
    ['accent2', 'draw.smartArt.accent2'],
    ['accent3', 'draw.smartArt.accent3'],
    ['accent4', 'draw.smartArt.accent4'],
    ['accent5', 'draw.smartArt.accent5'],
    ['accent6', 'draw.smartArt.accent6'],
  ];

  const palette = (c: SmartArtColors): string[] => SMARTART_COLORS[c].fill.map((slot) => theme[slot] ?? '4472C4');

  function outline(op: OutlineOp): void {
    if (!current || !info) return;
    const rows = outlineRows(info.nodes);
    const index = Math.min(shapeTools.smartArtRow, Math.max(0, rows.length - 1));
    const result = rows.length ? applyOutlineOp(rows, index, op, '') : { rows: [{ level: 0, text: '' }], index: 0 };
    shapeTools.smartArtRow = result.index;
    session.apply(commands.smartArtNodesCommand, { at: current.at, nodes: outlineToNodes(result.rows) });
  }

  const at = (): NonNullable<typeof current>['at'] | undefined => current?.at;
  function setLayout(layout: SmartArtLayout): void {
    const pos = at();
    if (pos) session.apply(commands.smartArtLayoutCommand, { at: pos, layout });
  }
  function setColors(colors: SmartArtColors): void {
    const pos = at();
    if (pos) session.apply(commands.smartArtColorsCommand, { at: pos, colors });
  }
  function setStyle(style: SmartArtStyle): void {
    const pos = at();
    if (pos) session.apply(commands.smartArtStyleCommand, { at: pos, style });
  }
</script>

<Group label={t('draw.smartArt.createGraphic')}>
  <SplitButton id="sa.addShape" size="large" icon="addShape" tip={t('draw.smartArt.addShape')} label={t('draw.smartArt.addShape')} onclick={() => outline('addAfter')} disabled={!current}>
    {#snippet menu()}
      <button class="mi" role="menuitem" onclick={() => outline('addAfter')}>{t('draw.smartArt.addAfter')}</button>
      <button class="mi" role="menuitem" onclick={() => outline('addBefore')}>{t('draw.smartArt.addBefore')}</button>
      <button class="mi" role="menuitem" onclick={() => outline('addAbove')}>{t('draw.smartArt.addAbove')}</button>
      <button class="mi" role="menuitem" onclick={() => outline('addBelow')}>{t('draw.smartArt.addBelow')}</button>
    {/snippet}
  </SplitButton>
  <div class="rows">
    <Button size="mid" icon="promote" tip={t('draw.smartArt.promote')} onclick={() => outline('promote')} disabled={!current} />
    <Button size="mid" icon="demote" tip={t('draw.smartArt.demote')} onclick={() => outline('demote')} disabled={!current} />
    <Button size="mid" icon="textPane" tip={t('draw.smartArt.textPane')} on={session.pane.left === 'smartArtText'} onclick={() => session.togglePane('left', 'smartArtText')} />
  </div>
  <div class="rows">
    <Button size="mid" icon="moveUp" tip={t('draw.smartArt.moveUp')} onclick={() => outline('moveUp')} disabled={!current} />
    <Button size="mid" icon="moveDown" tip={t('draw.smartArt.moveDown')} onclick={() => outline('moveDown')} disabled={!current} />
  </div>
</Group>

<Group label={t('draw.smartArt.layouts')} class="gallery-group">
  <div class="gallery">
    {#each LAYOUTS as layout (layout)}
      <button class="wk-smartart-tile" class:selected={info?.layout === layout} title={t(`smartArt.${layout}`)} aria-label={t(`smartArt.${layout}`)} onclick={() => setLayout(layout)} disabled={!current}>
        <SmartArtThumb {layout} colors={[theme.accent1 ?? '4472C4']} />
      </button>
    {/each}
  </div>
</Group>

<Group label={t('draw.smartArt.changeColors')}>
  <SplitButton id="sa.colors" size="large" icon="changeColors" tip={t('draw.smartArt.changeColors')} label={t('draw.smartArt.changeColors')} disabled={!current}>
    {#snippet menu()}
      {#each COLOR_HEADINGS as [category, heading] (category)}
        <div class="menu-head">{t(heading)}</div>
        <div class="wk-smartart-gallery">
          {#each COLORS.filter((c) => SMARTART_COLORS[c].category === category) as colors (colors)}
            <button class="wk-smartart-tile" class:selected={info?.colors === colors} role="menuitem" aria-label={t(heading)} onclick={() => setColors(colors)}>
              <SmartArtThumb layout={info?.layout ?? 'basicBlockList'} colors={palette(colors)} />
            </button>
          {/each}
        </div>
      {/each}
    {/snippet}
  </SplitButton>
</Group>

<Group label={t('draw.smartArt.styles')} class="gallery-group">
  <div class="gallery">
    {#each STYLES as style (style)}
      <button class="wk-smartart-tile" class:selected={info?.style === style} title={t(`smartArt.style.${style}`)} aria-label={t(`smartArt.style.${style}`)} onclick={() => setStyle(style)} disabled={!current}>
        <SmartArtThumb layout={info?.layout ?? 'basicBlockList'} colors={palette(info?.colors ?? DEFAULT_COLORS)} lineWidth={SMARTART_STYLES[style].line / 2} effect={SMARTART_STYLES[style].effect} />
      </button>
    {/each}
  </div>
</Group>

<Group label={t('draw.smartArt.reset')}>
  <Button size="large" icon="resetGraphic" tip={t('draw.smartArt.reset')} disabled={!current} onclick={() => { setColors(DEFAULT_COLORS); setStyle(DEFAULT_STYLE); }} />
</Group>
