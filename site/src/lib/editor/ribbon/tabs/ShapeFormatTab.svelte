<script lang="ts">
  /**
   * Shape Format (contextual, for shapes, text boxes, WordArt, ink and
   * SmartArt): Insert Shapes, Shape Styles, WordArt Styles, Text, Alt Text,
   * Arrange, Size and Format Pane, as in Word for Mac.
   */
  import {
    getShapeFill,
    getShapeLayout,
    getShapeStroke,
    getSmartArt,
    getTextBoxLayout,
    linkedTextBox,
    shapeKind,
    type ShapeFill,
    type ShapeStroke,
    type TextBoxLayout,
  } from '@office-kit/docx';
  import { commands, themeColors } from '@office-kit/docx-editor';
  import ArrangeGroup from '../ArrangeGroup.svelte';
  import Button from '../Button.svelte';
  import Group from '../Group.svelte';
  import SplitButton from '../SplitButton.svelte';
  import { getSession } from '../../session.svelte';
  import { t, type MessageKey } from '../../i18n/index.svelte';
  import ColorMenu from '../../shapes/ColorMenu.svelte';
  import ShapeGallery from '../../shapes/ShapeGallery.svelte';
  import { selectedShape, selectedSmartArt } from '../../shapes/selection';
  import { ARROW_STYLES, DASHES, fillCss, gradientVariations, OUTLINE_WEIGHTS, SHADOW_PRESETS, shapeStyles, wordArtStyles, type ShapeStylePreset } from '../../shapes/palette';
  import { shapeTools } from '../../shapes/tools.svelte';

  const session = getSession();
  const CM_PER_POINT = 2.54 / 72;
  // Shape Styles shows this many presets in the ribbon; the rest are in its menu.
  const INLINE_STYLES = 3;

  const current = $derived(selectedShape(session));
  const smartArt = $derived(selectedSmartArt(session));
  const kind = $derived(current && shapeKind(current.shape));
  const isWordArt = $derived(kind === 'wordArt');
  const theme = $derived(session.model && session.version >= 0 ? themeColors(session.model.doc) : {});
  const styles = $derived(shapeStyles(theme));
  const artStyles = $derived(wordArtStyles(theme));
  const fill = $derived(current && getShapeFill(current.shape));
  const stroke = $derived(current && getShapeStroke(current.shape));
  const fillColor = $derived(fill && 'color' in fill ? fill.color : (theme.accent1 ?? '4472C4'));
  const textLayout = $derived(current && kind !== 'ink' && kind !== 'wordArt' && kind !== 'group' && kind !== 'canvas' ? getTextBoxLayout(current.shape) : undefined);
  const size = $derived.by(() => {
    if (current) return getShapeLayout(current.shape);
    const model = session.model;
    if (smartArt && model) return getSmartArt(model.doc, smartArt.ref);
    return undefined;
  });
  const linked = $derived(current ? linkedTextBox(current.shape) !== undefined : false);

  function apply<P>(run: (at: NonNullable<typeof current>['at']) => P): P | undefined {
    return current ? run(current.at) : undefined;
  }

  const setFill = (f: ShapeFill): void => void apply((at) => session.apply(commands.shapeFillCommand, { at, fill: f }));
  const setStroke = (s: ShapeStroke | null): void => void apply((at) => session.apply(commands.shapeOutlineCommand, { at, stroke: s }));
  const patchStroke = (change: Partial<ShapeStroke>): void => setStroke({ color: stroke?.color ?? '000000', weight: stroke?.weight ?? 0.75, ...stroke, ...change });
  const setStyle = (p: ShapeStylePreset): void => void apply((at) => session.apply(commands.shapeStyleCommand, { at, fill: p.fill, stroke: p.stroke, shadow: p.shadow }));
  const setText = (change: Partial<TextBoxLayout>): void => void apply((at) => session.apply(commands.textBoxLayoutCommand, { at, layout: change }));

  async function pictureFill(e: Event): Promise<void> {
    const input = e.currentTarget as HTMLInputElement;
    const file = input.files?.[0];
    if (!file) return;
    const bytes = new Uint8Array(await file.arrayBuffer());
    setFill({ type: 'picture', bytes, contentType: file.type, tile: false });
    input.value = '';
  }

  function setSize(dim: 'width' | 'height', cm: number): void {
    if (!Number.isFinite(cm) || cm <= 0) return;
    const pt = cm / CM_PER_POINT;
    if (current) session.apply(commands.shapeLayoutCommand, { at: current.at, layout: dim === "width" ? { width: pt } : { height: pt } });
    else if (smartArt && size) session.apply(commands.smartArtSizeCommand, { at: smartArt.at, width: dim === 'width' ? pt : size.width, height: dim === 'height' ? pt : size.height });
  }

  function createLink(): void {
    if (!current) return;
    if (linked) session.apply(commands.linkTextBoxCommand, { at: current.at, to: null });
    else shapeTools.setTool({ kind: 'link', from: current.at });
  }

  const DIRECTIONS: ReadonlyArray<[TextBoxLayout['direction'], MessageKey]> = [
    ['horizontal', 'draw.dir.horizontal'],
    ['vertical', 'draw.dir.rotate90'],
    ['vertical270', 'draw.dir.rotate270'],
  ];
  const ANCHORS: ReadonlyArray<[TextBoxLayout['anchor'], MessageKey]> = [
    ['top', 'draw.alignTop'],
    ['middle', 'draw.alignMiddle'],
    ['bottom', 'draw.alignBottom'],
  ];
  const lineStyle = (dash: string): string => (dash === 'solid' ? 'solid' : dash.includes('dot') && !dash.includes('dash') ? 'dotted' : 'dashed');
</script>

{#snippet styleTile(p: ShapeStylePreset)}
  <button class="wk-style-tile" role="menuitem" title={t('draw.shapeStyles')} onclick={() => setStyle(p)} disabled={!current}>
    <i style="background:{fillCss(p.fill)};border:{p.stroke ? `${Math.min(p.stroke.weight, 3)}px solid #${p.stroke.color}` : 'none'};color:#{p.text};box-shadow:{p.shadow ? '1px 2px 3px rgb(0 0 0 / 35%)' : 'none'}">Abc</i>
  </button>
{/snippet}
{#snippet artTile(p: ShapeStylePreset)}
  <button class="wk-style-tile wk-wordart-tile" role="menuitem" title={t('draw.wordArtStyles')} onclick={() => setStyle(p)} disabled={!isWordArt}>
    <span style="color:{p.fill.type === 'solid' ? `#${p.fill.color}` : 'transparent'};background:{p.fill.type === 'gradient' ? fillCss(p.fill) : 'none'};-webkit-background-clip:text;background-clip:text;-webkit-text-stroke:{p.stroke ? `1px #${p.stroke.color}` : '0'}">A</span>
  </button>
{/snippet}

<Group label={t('draw.insertShapes')}>
  <SplitButton id="sf.shapes" size="large" icon="shapes" tip={t('draw.shapes')} label={t('draw.shapes')}>
    {#snippet menu()}<ShapeGallery onpick={(p) => { session.openMenu = null; shapeTools.pickPreset(p); }} />{/snippet}
  </SplitButton>
  <div class="rows">
    <SplitButton id="sf.editShape" icon="editShape" tip={t('draw.editShape')} disabled={!current || kind === 'wordArt' || kind === 'ink'}>
      {#snippet menu()}
        <div class="menu-head">{t('draw.changeShape')}</div>
        <ShapeGallery showRecent={false} onpick={(preset) => apply((at) => session.apply(commands.changeShapeCommand, { at, preset }))} />
        <hr />
        <button class="mi check" class:checked={shapeTools.editPoints} role="menuitem" onclick={() => { session.openMenu = null; shapeTools.editPoints = !shapeTools.editPoints; }}>{t('draw.editPoints')}</button>
      {/snippet}
    </SplitButton>
    <Button icon="textBox" tip={t('draw.drawTextBox')} onclick={() => shapeTools.pickPreset('textBox')} />
  </div>
</Group>

<Group label={t('draw.shapeStyles')} class="gallery-group">
  <div class="gallery">
    {#each (styles[1] ?? []).slice(0, INLINE_STYLES) as p, i (i)}{@render styleTile(p)}{/each}
  </div>
  <SplitButton id="sf.styles" tip={t('draw.shapeStyles')} disabled={!current}>
    {#snippet menu()}
      <div class="menu-head">{t('draw.themeStyles')}</div>
      <div class="wk-style-menu">
        {#each styles as row, r (r)}{#each row as p, c (c)}{@render styleTile(p)}{/each}{/each}
      </div>
    {/snippet}
  </SplitButton>
</Group>

<Group label={t('draw.shapeFill')}>
  <SplitButton id="sf.fill" size="large" icon="shapeFill" tip={t('draw.shapeFill')} label={t('draw.shapeFill')} onclick={() => setFill({ type: 'solid', color: fillColor })} disabled={!current}>
    {#snippet menu()}
      <ColorMenu onpick={(color) => setFill({ type: 'solid', color })} noneLabel={t('draw.noFill')} onnone={() => setFill({ type: 'none' })}>
        <label class="mi">{t('draw.fill.picture')}…<input type="file" accept="image/png,image/jpeg,image/gif" onchange={pictureFill} hidden /></label>
        <div class="menu-head">{t('draw.fill.gradient')}</div>
        <div class="grid five">
          {#each gradientVariations(fillColor) as g, i (i)}
            <button class="chip" role="menuitem" style="background:{fillCss(g)}" aria-label={t('draw.fill.gradient')} onclick={() => setFill(g)}></button>
          {/each}
        </div>
        <button class="mi" role="menuitem" onclick={() => { session.pane.right = 'formatShape'; session.openMenu = null; }}>{t('draw.moreFill')}</button>
      </ColorMenu>
    {/snippet}
  </SplitButton>
  <div class="rows">
    <SplitButton id="sf.outline" icon="shapeOutline" tip={t('draw.shapeOutline')} onclick={() => patchStroke({})} disabled={!current}>
      {#snippet menu()}
        <ColorMenu onpick={(color) => patchStroke({ color })} noneLabel={t('draw.noOutline')} onnone={() => setStroke(null)}>
          <div class="menu-head">{t('draw.weight')}</div>
          {#each OUTLINE_WEIGHTS as w (w)}
            <button class="mi check" class:checked={stroke?.weight === w} role="menuitem" onclick={() => patchStroke({ weight: w })}><span class="wk-line-sample" style="border-top:{Math.max(1, w * 1.33)}px solid"></span>{w} pt</button>
          {/each}
          <div class="menu-head">{t('draw.dashes')}</div>
          {#each DASHES as d (d)}
            <button class="mi check" class:checked={(stroke?.dash ?? 'solid') === d} role="menuitem" onclick={() => patchStroke({ dash: d })}><span class="wk-line-sample" style="border-top:2px {lineStyle(d)}"></span>{t(`draw.dash.${d}`)}</button>
          {/each}
          <div class="menu-head">{t('draw.arrows')}</div>
          {#each ARROW_STYLES as [start, end] (`${start}-${end}`)}
            <button class="mi" role="menuitem" onclick={() => patchStroke({ startArrow: start, endArrow: end })}>{t(`draw.arrow.${start}`)} → {t(`draw.arrow.${end}`)}</button>
          {/each}
        </ColorMenu>
      {/snippet}
    </SplitButton>
    <SplitButton id="sf.effects" icon="shapeEffects" tip={t('draw.shapeEffects')} disabled={!current}>
      {#snippet menu()}
        <div class="menu-head">{t('draw.shadow')}</div>
        <button class="mi" role="menuitem" onclick={() => apply((at) => session.apply(commands.shapeShadowCommand, { at, shadow: null }))}>{t('draw.noShadow')}</button>
        <div class="menu-head">{t('draw.outer')}</div>
        <div class="grid five">
          {#each SHADOW_PRESETS as p (p.id)}
            <button class="chip" role="menuitem" title={t(`draw.shadow.${p.id}`)} aria-label={t(`draw.shadow.${p.id}`)} style="background:#fff;box-shadow:{p.shadow.offsetX}px {p.shadow.offsetY}px 2px rgb(0 0 0 / 45%)" onclick={() => apply((at) => session.apply(commands.shapeShadowCommand, { at, shadow: p.shadow }))}></button>
          {/each}
        </div>
      {/snippet}
    </SplitButton>
  </div>
</Group>

<Group label={t('draw.wordArtStyles')} class="gallery-group">
  <div class="gallery">
    {#each artStyles.slice(0, INLINE_STYLES) as p, i (i)}{@render artTile(p)}{/each}
  </div>
  <SplitButton id="sf.wordArtStyles" tip={t('draw.wordArtStyles')} disabled={!isWordArt}>
    {#snippet menu()}
      <div class="wk-style-menu wk-wordart-menu">{#each artStyles as p, i (i)}{@render artTile(p)}{/each}</div>
    {/snippet}
  </SplitButton>
  <div class="rows">
    <SplitButton id="sf.textFill" icon="textFill" tip={t('draw.textFill')} disabled={!isWordArt}>
      {#snippet menu()}<ColorMenu onpick={(color) => setFill({ type: 'solid', color })} noneLabel={t('draw.noFill')} onnone={() => setFill({ type: 'none' })} />{/snippet}
    </SplitButton>
    <SplitButton id="sf.textOutline" icon="textOutline" tip={t('draw.textOutline')} disabled={!isWordArt}>
      {#snippet menu()}<ColorMenu onpick={(color) => patchStroke({ color })} noneLabel={t('draw.noOutline')} onnone={() => setStroke(null)} />{/snippet}
    </SplitButton>
    <SplitButton id="sf.textEffects" icon="textEffects" tip={t('draw.textEffects')} disabled={!isWordArt}>
      {#snippet menu()}
        <button class="mi" role="menuitem" onclick={() => apply((at) => session.apply(commands.shapeShadowCommand, { at, shadow: null }))}>{t('draw.noShadow')}</button>
        <div class="grid five">
          {#each SHADOW_PRESETS as p (p.id)}
            <button class="chip" role="menuitem" title={t(`draw.shadow.${p.id}`)} aria-label={t(`draw.shadow.${p.id}`)} style="background:#fff;box-shadow:{p.shadow.offsetX}px {p.shadow.offsetY}px 2px rgb(0 0 0 / 45%)" onclick={() => apply((at) => session.apply(commands.shapeShadowCommand, { at, shadow: p.shadow }))}></button>
          {/each}
        </div>
      {/snippet}
    </SplitButton>
  </div>
</Group>

<Group label={t('draw.text')}>
  <div class="rows">
    <SplitButton id="sf.textDirection" size="mid" icon="textDirection" tip={t('draw.textDirection')} label={t('draw.textDirection')} disabled={!textLayout}>
      {#snippet menu()}
        {#each DIRECTIONS as [dir, label] (dir)}
          <button class="mi check" class:checked={textLayout?.direction === dir} role="menuitem" onclick={() => setText({ direction: dir })}>{t(label)}</button>
        {/each}
      {/snippet}
    </SplitButton>
    <SplitButton id="sf.alignText" size="mid" icon="alignText" tip={t('draw.alignText')} label={t('draw.alignText')} disabled={!textLayout}>
      {#snippet menu()}
        {#each ANCHORS as [anchor, label] (anchor)}
          <button class="mi check" class:checked={textLayout?.anchor === anchor} role="menuitem" onclick={() => setText({ anchor })}>{t(label)}</button>
        {/each}
      {/snippet}
    </SplitButton>
    <Button size="mid" icon="createLink" tip={linked ? t('draw.breakLink') : t('draw.createLink')} onclick={createLink} disabled={kind !== 'textBox'} on={shapeTools.tool.kind === 'link'} />
  </div>
</Group>

<Group label={t('draw.accessibility')}>
  <Button size="large" icon="altText" tip={t('draw.altText')} onclick={() => (session.pane.right = 'formatShape')} disabled={!current} />
</Group>

<ArrangeGroup />

<Group label={t('draw.size')}>
  <div class="rows">
    <label class="wk-size-spin" title={t('draw.height')}>
      <span aria-hidden="true">↕</span>
      <input class="combo" type="number" min="0.01" step="0.1" aria-label={t('draw.height')} value={size ? +(size.height * CM_PER_POINT).toFixed(2) : ''} disabled={!size} onchange={(e) => setSize('height', Number(e.currentTarget.value))} /> cm
    </label>
    <label class="wk-size-spin" title={t('draw.width')}>
      <span aria-hidden="true">↔</span>
      <input class="combo" type="number" min="0.01" step="0.1" aria-label={t('draw.width')} value={size ? +(size.width * CM_PER_POINT).toFixed(2) : ''} disabled={!size} onchange={(e) => setSize('width', Number(e.currentTarget.value))} /> cm
    </label>
  </div>
</Group>

<Group label={t('draw.formatPane')}>
  <Button size="large" icon="formatPane" tip={t('draw.formatPane')} on={session.pane.right === 'formatShape'} onclick={() => session.togglePane('right', 'formatShape')} />
</Group>
