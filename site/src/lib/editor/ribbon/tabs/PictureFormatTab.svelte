<script lang="ts">
  /**
   * Word's Picture Format tab: Adjust, Picture Styles, Accessibility, Arrange
   * and Size. Effects without an ECMA-376 form (Remove Background, Artistic
   * Effects, Sharpen/Soften, Color Tone) are shown disabled.
   */
  import { commands, CROP_SHAPES, presetPath } from '@office-kit/docx-editor';
  import type { PictureShadow } from '@office-kit/docx';
  import RibbonIcon from '../../RibbonIcon.svelte';
  import ArrangeGroup from '../ArrangeGroup.svelte';
  import Button from '../Button.svelte';
  import Group from '../Group.svelte';
  import SplitButton from '../SplitButton.svelte';
  import CompressDialog from '../../picture/CompressDialog.svelte';
  import { getSession } from '../../session.svelte';
  import { t, type MessageKey } from '../../i18n/index.svelte';
  import { applyToSelected, EMU_PER_CM, EMU_PER_PT, selected } from '../../picture/state';
  import { patchAdjustments, patchEffects, RECOLOR_KEYS } from '../../picture/edit';
  import { changePictureFrom } from '../../picture/insert';
  import { picturePreview } from '../../picture/preview';
  import { pictureModes } from '../../picture/ui.svelte';
  import {
    ASPECT_RATIOS,
    BEVELS,
    BORDER_DASHES,
    BORDER_WEIGHTS,
    CORRECTION_STEPS,
    GLOW_COLORS,
    GLOW_OPACITY,
    GLOW_SIZES,
    PICTURE_STYLES,
    RECOLORS,
    REFLECTIONS,
    ROTATIONS_3D,
    SATURATIONS,
    SHADOWS,
    SOFT_EDGES,
    STANDARD_COLORS,
    THEME_COLORS,
    TRANSPARENCIES,
  } from '../../picture/presets';

  const session = getSession();
  const sel = $derived(selected(session));
  const info = $derived(sel?.info);
  const picture = $derived(info?.picture);
  const adj = $derived(picture?.adjustments ?? {});
  // Picture Styles shows this many tiles in the ribbon; the rest open from the chevron.
  const INLINE_STYLES = 4;
  const PREVIEW_W = 64;
  const PREVIEW_H = 48;
  const SMALL_W = 44;
  const SMALL_H = 34;

  type EffectSection = 'shadow' | 'reflection' | 'glow' | 'softEdges' | 'bevel' | 'rotation3d';
  const EFFECT_SECTIONS: ReadonlyArray<readonly [EffectSection, MessageKey]> = [
    ['shadow', 'pic.shadow'],
    ['reflection', 'pic.reflection'],
    ['glow', 'pic.glow'],
    ['softEdges', 'pic.softEdges'],
    ['bevel', 'pic.bevel'],
    ['rotation3d', 'pic.rotation3d'],
  ];
  let effectSection = $state<EffectSection>('shadow');
  type BorderSection = 'colors' | 'weight' | 'dashes';
  let borderSection = $state<BorderSection>('colors');
  type CropSection = 'shape' | 'aspect' | null;
  let cropSection = $state<CropSection>(null);

  let changeInput = $state<HTMLInputElement | null>(null);

  const cm = (emu: number): number => Math.round((emu / EMU_PER_CM) * 100) / 100;

  function setSize(dim: 'height' | 'width', value: number): void {
    const i = info;
    if (!i || !(value > 0)) return;
    const emu = value * EMU_PER_CM;
    const k = dim === 'height' ? emu / i.heightEmu : emu / i.widthEmu;
    const cxEmu = dim === 'width' ? emu : i.lockAspect ? i.widthEmu * k : i.widthEmu;
    const cyEmu = dim === 'height' ? emu : i.lockAspect ? i.heightEmu * k : i.heightEmu;
    applyToSelected(session, commands.resizeImageCommand, { cxEmu, cyEmu });
  }

  function setBorder(patch: Partial<{ color: string; widthEmu: number; dash: (typeof BORDER_DASHES)[number] }>): void {
    const cur = picture?.outline ?? { color: '000000', widthEmu: 0.75 * EMU_PER_PT, dash: 'solid' as const, compound: 'sng' as const };
    applyToSelected(session, commands.pictureBorderCommand, { outline: { ...cur, ...patch } });
  }

  async function onChangeFile(e: Event): Promise<void> {
    const input = e.currentTarget as HTMLInputElement;
    const file = input.files?.[0];
    input.value = '';
    if (file && sel) await changePictureFrom(session, sel.index, file);
  }

  function preview(patch: Parameters<typeof picturePreview>[1]): string {
    return picture ? picturePreview(picture, patch, PREVIEW_W, PREVIEW_H) : '';
  }

  function shadowLabel(s: PictureShadow): string {
    return s.kind === 'inner' ? t('pic.inner') : t('pic.outer');
  }

  function openPane(id: string): void {
    session.openMenu = null;
    if (session.pane.right !== id) session.togglePane('right', id);
  }
</script>

<Group label={t('pic.adjust')}>
  <Button size="large" icon="removeBackground" tip={t('pic.removeBackground')} onclick={() => {}} disabled />
  <SplitButton id="pic.corrections" size="large" icon="corrections" tip={t('pic.corrections')} disabled={!picture}>
    {#snippet menu()}
      <div class="menu-head">{t('pic.brightnessContrast')}</div>
      <div class="grid pic-five">
        {#each CORRECTION_STEPS as contrast (contrast)}
          {#each CORRECTION_STEPS as brightness (brightness)}
            {@const label = `${t('pic.brightness')}: ${brightness}% ${t('pic.contrast')}: ${contrast}%`}
            <button class="pic-tile" class:selected={(adj.brightness ?? 0) === brightness && (adj.contrast ?? 0) === contrast} title={label} aria-label={label}
              onclick={() => patchAdjustments(session, { brightness: brightness || undefined, contrast: contrast || undefined })}
              >{@html preview({ adjustments: { ...adj, brightness, contrast } })}</button>
          {/each}
        {/each}
      </div>
      <hr />
      <button class="mi" role="menuitem" onclick={() => openPane('formatPicture')}>{t('pic.correctionsOptions')}</button>
    {/snippet}
  </SplitButton>
  <div class="col">
  <SplitButton id="pic.color" size="mid" icon="pictureColor" tip={t('pic.color')} disabled={!picture}>
    {#snippet menu()}
      <div class="menu-head">{t('pic.saturation')}</div>
      <div class="pic-row">
        {#each SATURATIONS as sat (sat)}
          <button class="pic-tile" class:selected={(adj.saturation ?? 100) === sat} title="{t('pic.saturation')}: {sat}%" aria-label="{t('pic.saturation')}: {sat}%"
            onclick={() => patchAdjustments(session, { saturation: sat === 100 ? undefined : sat })}
            >{@html preview({ adjustments: { ...adj, saturation: sat } })}</button>
        {/each}
      </div>
      <div class="menu-head">{t('pic.recolor')}</div>
      <div class="grid pic-seven">
        {#each RECOLORS as r (r.id)}
          <button class="pic-tile" title={r.name} aria-label={r.name}
            onclick={() => patchAdjustments(session, { ...RECOLOR_KEYS, ...r.adjustments })}
            >{@html preview({ adjustments: r.adjustments })}</button>
        {/each}
      </div>
      <hr />
      <button class="mi" role="menuitem" onclick={() => { session.openMenu = null; pictureModes.pickingTransparent = true; }}>{t('pic.setTransparent')}</button>
      <button class="mi" role="menuitem" onclick={() => openPane('formatPicture')}>{t('pic.colorOptions')}</button>
    {/snippet}
  </SplitButton>
  <Button size="mid" icon="artisticEffects" tip={t('pic.artisticEffects')} onclick={() => {}} disabled />
  <SplitButton id="pic.transparency" size="mid" icon="transparency" tip={t('pic.transparency')} disabled={!picture}>
    {#snippet menu()}
      <div class="pic-row">
        {#each TRANSPARENCIES as tr (tr)}
          <button class="pic-tile" class:selected={(adj.transparency ?? 0) === tr} title="{t('pic.transparency')}: {tr}%" aria-label="{t('pic.transparency')}: {tr}%"
            onclick={() => patchAdjustments(session, { transparency: tr || undefined })}
            >{@html preview({ adjustments: { ...adj, transparency: tr } })}</button>
        {/each}
      </div>
      <hr />
      <button class="mi" role="menuitem" onclick={() => openPane('formatPicture')}>{t('pic.transparencyOptions')}</button>
    {/snippet}
  </SplitButton>
  </div>
  <div class="col">
    <Button icon="compress" tip={t('pic.compress')} onclick={() => session.openDialog('picture.compress')} disabled={!picture} />
    <SplitButton id="pic.change" icon="changePicture" tip={t('pic.change')} disabled={!picture}>
      {#snippet menu()}
        <button class="mi" role="menuitem" onclick={() => { session.openMenu = null; changeInput?.click(); }}>{t('pic.fromFile')}</button>
      {/snippet}
    </SplitButton>
    <SplitButton id="pic.reset" icon="resetPicture" tip={t('pic.reset')} disabled={!picture} onclick={() => applyToSelected(session, commands.resetPictureCommand, { size: false })}>
      {#snippet menu()}
        <button class="mi" role="menuitem" onclick={() => applyToSelected(session, commands.resetPictureCommand, { size: false })}>{t('pic.reset')}</button>
        <button class="mi" role="menuitem" onclick={() => applyToSelected(session, commands.resetPictureCommand, { size: true })}>{t('pic.resetSize')}</button>
      {/snippet}
    </SplitButton>
    <input bind:this={changeInput} type="file" accept="image/*" onchange={onChangeFile} hidden />
  </div>
</Group>

<Group label={t('pic.styles')}>
  <div class="gallery-group">
    <div class="gallery pic-styles">
      {#each PICTURE_STYLES.slice(0, INLINE_STYLES) as s (s.id)}
        <button class="pic-tile small" title={s.name} aria-label={s.name} disabled={!picture}
          onclick={() => applyToSelected(session, commands.pictureStyleCommand, { outline: s.outline, effects: s.effects, geometry: s.geometry })}
          >{@html picture ? picturePreview(picture, { outline: s.outline, effects: s.effects, geometry: s.geometry }, SMALL_W, SMALL_H) : ''}</button>
      {/each}
    </div>
    <SplitButton id="pic.stylesMore" tip={t('pic.styles')} disabled={!picture}>
      {#snippet menu()}
        <div class="grid pic-styles-all">
          {#each PICTURE_STYLES as s (s.id)}
            <button class="pic-tile" title={s.name} aria-label={s.name}
              onclick={() => applyToSelected(session, commands.pictureStyleCommand, { outline: s.outline, effects: s.effects, geometry: s.geometry })}
              >{@html preview({ outline: s.outline, effects: s.effects, geometry: s.geometry })}</button>
          {/each}
        </div>
      {/snippet}
    </SplitButton>
  </div>
  <div class="col">
    <SplitButton id="pic.border" size="mid" icon="pictureBorder" tip={t('pic.border')} disabled={!picture}>
      {#snippet menu()}
        <div class="pic-sections">
          <button class="mi" class:checked={borderSection === 'colors'} onclick={() => (borderSection = 'colors')}>{t('pic.colors')}</button>
          <button class="mi" class:checked={borderSection === 'weight'} onclick={() => (borderSection = 'weight')}>{t('pic.weight')}</button>
          <button class="mi" class:checked={borderSection === 'dashes'} onclick={() => (borderSection = 'dashes')}>{t('pic.dashes')}</button>
        </div>
        {#if borderSection === 'colors'}
          <div class="menu-head">{t('pic.themeColors')}</div>
          <div class="grid ten">
            {#each THEME_COLORS as c (c)}<button class="chip" style="background:#{c}" title="#{c}" aria-label="#{c}" onclick={() => setBorder({ color: c })}></button>{/each}
          </div>
          <div class="menu-head">{t('pic.standardColors')}</div>
          <div class="grid ten">
            {#each STANDARD_COLORS as c (c)}<button class="chip" style="background:#{c}" title="#{c}" aria-label="#{c}" onclick={() => setBorder({ color: c })}></button>{/each}
          </div>
          <label class="mi">{t('pic.moreColors')} <input type="color" onchange={(e) => setBorder({ color: e.currentTarget.value.slice(1).toUpperCase() })} /></label>
        {:else if borderSection === 'weight'}
          {#each BORDER_WEIGHTS as w (w)}
            <button class="mi check" class:checked={picture?.outline?.widthEmu === w * EMU_PER_PT} onclick={() => setBorder({ widthEmu: w * EMU_PER_PT })}>
              <span class="pic-line" style="border-top-width:{Math.max(1, w * 1.33)}px"></span>{w} pt
            </button>
          {/each}
        {:else}
          {#each BORDER_DASHES as d (d)}
            <button class="mi check" class:checked={picture?.outline?.dash === d} title={d} aria-label={d} onclick={() => setBorder({ dash: d })}>
              <svg width="90" height="6" aria-hidden="true"><path d="M0 3h90" stroke="#000" stroke-width="2" stroke-dasharray={d === 'solid' ? undefined : d.includes('Dot') && !d.includes('Dash') ? '2 2' : d.startsWith('lg') ? '8 3' : '4 3'} /></svg>
            </button>
          {/each}
        {/if}
        <hr />
        <button class="mi" role="menuitem" onclick={() => applyToSelected(session, commands.pictureBorderCommand, { outline: undefined })}>{t('pic.noOutline')}</button>
      {/snippet}
    </SplitButton>
    <SplitButton id="pic.effects" size="mid" icon="pictureEffects" tip={t('pic.effects')} disabled={!picture}>
      {#snippet menu()}
        <div class="pic-sections">
          {#each EFFECT_SECTIONS as [id, key] (id)}
            <button class="mi" class:checked={effectSection === id} onclick={() => (effectSection = id)}>{t(key)}</button>
          {/each}
        </div>
        {#if effectSection === 'shadow'}
          <button class="mi" onclick={() => patchEffects(session, { shadow: undefined })}>{t('pic.noShadow')}</button>
          <div class="grid pic-nine">
            {#each SHADOWS as s (s.name)}
              <button class="pic-tile" title="{shadowLabel(s.shadow)}: {s.name}" aria-label={s.name} onclick={() => patchEffects(session, { shadow: s.shadow })}
                >{@html preview({ effects: { shadow: s.shadow }, outline: undefined })}</button>
            {/each}
          </div>
        {:else if effectSection === 'reflection'}
          <button class="mi" onclick={() => patchEffects(session, { reflection: undefined })}>{t('pic.noReflection')}</button>
          <div class="grid pic-three">
            {#each REFLECTIONS as r (r.name)}
              {@const { name, ...reflection } = r}
              <button class="pic-tile tall" title={name} aria-label={name} onclick={() => patchEffects(session, { reflection })}
                >{@html preview({ effects: { reflection }, outline: undefined })}</button>
            {/each}
          </div>
        {:else if effectSection === 'glow'}
          <button class="mi" onclick={() => patchEffects(session, { glow: undefined })}>{t('pic.noGlow')}</button>
          <div class="grid pic-six">
            {#each GLOW_SIZES as size (size)}
              {#each GLOW_COLORS as color (color)}
                <button class="chip pic-glow" title="{size} pt" aria-label="{size} pt #{color}" style="box-shadow:0 0 {size / 2}px {size / 3}px #{color}"
                  onclick={() => patchEffects(session, { glow: { radiusEmu: size * EMU_PER_PT, color, opacity: GLOW_OPACITY } })}></button>
              {/each}
            {/each}
          </div>
        {:else if effectSection === 'softEdges'}
          <button class="mi" onclick={() => patchEffects(session, { softEdgeEmu: undefined })}>{t('pic.noSoftEdges')}</button>
          {#each SOFT_EDGES as pt (pt)}
            <button class="mi check" class:checked={picture?.effects.softEdgeEmu === pt * EMU_PER_PT} onclick={() => patchEffects(session, { softEdgeEmu: pt * EMU_PER_PT })}>{pt} pt</button>
          {/each}
        {:else if effectSection === 'bevel'}
          <button class="mi" onclick={() => patchEffects(session, { bevel: undefined })}>{t('pic.noBevel')}</button>
          <div class="grid pic-four">
            {#each BEVELS as preset (preset)}
              <button class="pic-tile" title={preset} aria-label={preset}
                onclick={() => patchEffects(session, { bevel: { preset, widthEmu: 6 * EMU_PER_PT, heightEmu: 6 * EMU_PER_PT } })}
                >{@html preview({ effects: { bevel: { preset, widthEmu: 6 * EMU_PER_PT, heightEmu: 6 * EMU_PER_PT } }, outline: undefined })}</button>
            {/each}
          </div>
        {:else}
          <button class="mi" onclick={() => patchEffects(session, { rotation3d: undefined })}>{t('pic.noRotation')}</button>
          <div class="grid pic-four">
            {#each ROTATIONS_3D as camera (camera)}
              <button class="pic-tile" title={camera} aria-label={camera} onclick={() => patchEffects(session, { rotation3d: { camera } })}
                >{@html preview({ effects: { rotation3d: { camera } }, outline: undefined })}</button>
            {/each}
          </div>
        {/if}
        <hr />
        <button class="mi" role="menuitem" onclick={() => openPane('formatPicture')}>{t('pic.effectsOptions')}</button>
      {/snippet}
    </SplitButton>
  </div>
</Group>

<Group label={t('pic.accessibility')}>
  <Button size="large" icon="altText" tip={t('pic.altText')} on={session.pane.right === 'altText'} onclick={() => session.togglePane('right', 'altText')} disabled={!info} />
</Group>

<ArrangeGroup />

<Group label={t('pic.size')}>
  <SplitButton id="pic.crop" size="large" icon="picCrop" tip={t('pic.crop')} on={pictureModes.cropping} disabled={!picture} onclick={() => (pictureModes.cropping = !pictureModes.cropping)}>
    {#snippet menu()}
      <button class="mi check" class:checked={pictureModes.cropping} onclick={() => { session.openMenu = null; pictureModes.cropping = !pictureModes.cropping; }}>{t('pic.crop')}</button>
      <button class="mi" onclick={() => (cropSection = cropSection === 'shape' ? null : 'shape')}>{t('pic.cropToShape')} ›</button>
      {#if cropSection === 'shape'}
        <div class="grid pic-shapes">
          {#each CROP_SHAPES as prst (prst)}
            <button class="chip pic-shape" title={prst} aria-label={prst} onclick={() => applyToSelected(session, commands.pictureShapeCommand, { preset: prst })}>
              <svg viewBox="-1 -1 22 22" width="20" height="20" aria-hidden="true"><path d={presetPath(prst, 20, 20)} fill="#DAE3F3" stroke="#4472C4" /></svg>
            </button>
          {/each}
        </div>
      {/if}
      <button class="mi" onclick={() => (cropSection = cropSection === 'aspect' ? null : 'aspect')}>{t('pic.aspectRatio')} ›</button>
      {#if cropSection === 'aspect'}
        {#each Object.entries(ASPECT_RATIOS) as [group, ratios] (group)}
          <div class="menu-head">{t(`pic.aspect.${group as keyof typeof ASPECT_RATIOS}`)}</div>
          {#each ratios as [w, h] (`${w}:${h}`)}
            <button class="mi" onclick={() => applyToSelected(session, commands.cropToAspectCommand, { mode: { ratio: w / h } })}>{w}:{h}</button>
          {/each}
        {/each}
      {/if}
      <hr />
      <button class="mi" onclick={() => applyToSelected(session, commands.cropToAspectCommand, { mode: 'fill' })}>{t('pic.fill')}</button>
      <button class="mi" onclick={() => applyToSelected(session, commands.cropToAspectCommand, { mode: 'fit' })}>{t('pic.fit')}</button>
    {/snippet}
  </SplitButton>
  <div class="rows pic-size">
    <label class="field" title="{t('pic.height')} (cm)">
      <RibbonIcon name="picHeight" size={16} />
      <input type="number" min="0.01" step="0.1" value={info ? cm(info.heightEmu) : ''} disabled={!info} onchange={(e) => setSize('height', e.currentTarget.valueAsNumber)} />
    </label>
    <label class="field" title="{t('pic.width')} (cm)">
      <RibbonIcon name="picWidth" size={16} />
      <input type="number" min="0.01" step="0.1" value={info ? cm(info.widthEmu) : ''} disabled={!info} onchange={(e) => setSize('width', e.currentTarget.valueAsNumber)} />
    </label>
    <div class="row">
      <Button icon="lockAspect" tip={t('pic.lockAspect')} on={info?.lockAspect ?? false} disabled={!info} onclick={() => applyToSelected(session, commands.drawingAspectLockCommand, { locked: !info?.lockAspect })} />
      <Button icon="arrPosition" tip={t('pic.sizeDialog')} disabled={!info} onclick={() => session.openDialog('picture.size')} />
      <Button icon="formatPane" tip={t('pic.formatPane')} on={session.pane.right === 'formatPicture'} disabled={!info} onclick={() => session.togglePane('right', 'formatPicture')} />
    </div>
  </div>
</Group>

<CompressDialog />

<style>
  .pic-tile {
    display: grid;
    place-items: center;
    width: 72px;
    height: 56px;
    padding: 2px;
    overflow: visible;
    border: 1px solid transparent;
    border-radius: 3px;
    background: #fff;
    cursor: pointer;
  }
  .pic-tile.small {
    width: 52px;
    height: 44px;
  }
  .pic-tile.tall {
    height: 96px;
    align-items: start;
  }
  .pic-tile:hover,
  .pic-tile.selected {
    border-color: #f0a30a;
  }
  .pic-tile :global(svg) {
    overflow: visible;
  }
  .pic-row {
    display: flex;
    gap: 4px;
    padding: 4px;
  }
  .pic-five { grid-template-columns: repeat(5, 72px); }
  .pic-seven { grid-template-columns: repeat(7, 72px); }
  .pic-nine { grid-template-columns: repeat(3, 72px); }
  .pic-three { grid-template-columns: repeat(3, 72px); }
  .pic-four { grid-template-columns: repeat(4, 72px); }
  .pic-six { grid-template-columns: repeat(6, 28px); gap: 10px; padding: 10px; }
  .pic-shapes { grid-template-columns: repeat(8, 26px); }
  .pic-styles-all { grid-template-columns: repeat(5, 72px); }
  .pic-styles {
    gap: 2px;
  }
  .pic-sections {
    display: flex;
    gap: 2px;
    border-bottom: 1px solid var(--chrome-line);
    margin-bottom: 4px;
  }
  .pic-sections .mi {
    width: auto;
  }
  .pic-sections .mi.checked {
    background: rgba(0, 0, 0, 0.08);
  }
  .pic-line {
    display: inline-block;
    width: 60px;
    border-top: solid #000;
  }
  .pic-glow {
    width: 20px;
    background: #fff;
  }
  .pic-shape {
    aspect-ratio: auto;
    background: #fff;
  }
  .pic-size input {
    width: 48px;
  }
</style>
