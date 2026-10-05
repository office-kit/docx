<script lang="ts">
  /**
   * Word's Format Picture pane: Line, Effects, Size and Picture sections with
   * exact values for what the ribbon galleries set by preset.
   */
  import { commands } from '@office-kit/docx-editor';
  import type { LineDash, PictureCrop, PictureShadow } from '@office-kit/docx';
  import RibbonIcon from '../RibbonIcon.svelte';
  import { getSession } from '../session.svelte';
  import { t, type MessageKey } from '../i18n/index.svelte';
  import { applyToSelected, EMU_PER_CM, EMU_PER_PT, selected } from '../picture/state';
  import { patchAdjustments, patchEffects } from '../picture/edit';
  import { BORDER_DASHES } from '../picture/presets';

  const session = getSession();
  const info = $derived(selected(session)?.info);
  const picture = $derived(info?.picture);

  type Section = 'line' | 'effects' | 'size' | 'picture';
  const SECTIONS: ReadonlyArray<readonly [Section, MessageKey]> = [
    ['line', 'fmtPic.line'],
    ['effects', 'pic.effects'],
    ['size', 'pic.size'],
    ['picture', 'fmtPic.picture'],
  ];
  let section = $state<Section>('picture');

  const pt = (emu: number): number => Math.round((emu / EMU_PER_PT) * 100) / 100;
  const cm = (emu: number): number => Math.round((emu / EMU_PER_CM) * 100) / 100;
  /** A number input's value, or undefined while it is empty / invalid. */
  const num = (e: Event): number | undefined => {
    const v = (e.currentTarget as HTMLInputElement).valueAsNumber;
    return Number.isFinite(v) ? v : undefined;
  };

  const DEFAULT_SHADOW: PictureShadow = { kind: 'outer', color: '000000', opacity: 40, blurEmu: 4 * EMU_PER_PT, distanceEmu: 3 * EMU_PER_PT, directionDeg: 45 };

  function setLine(patch: Partial<{ color: string; widthEmu: number; dash: LineDash }>): void {
    const cur = picture?.outline ?? { color: '000000', widthEmu: 0.75 * EMU_PER_PT, dash: 'solid' as const, compound: 'sng' as const };
    applyToSelected(session, commands.pictureBorderCommand, { outline: { ...cur, ...patch } });
  }

  function setShadow(patch: Partial<PictureShadow>): void {
    patchEffects(session, { shadow: { ...(picture?.effects.shadow ?? DEFAULT_SHADOW), ...patch } });
  }

  function setCrop(side: keyof PictureCrop, value: number | undefined): void {
    const crop = picture?.crop;
    if (!crop || value === undefined) return;
    applyToSelected(session, commands.cropPictureCommand, { crop: { ...crop, [side]: value } });
  }

  function setSize(dim: 'w' | 'h', value: number | undefined): void {
    if (!info || !value || value <= 0) return;
    const emu = value * EMU_PER_CM;
    const k = emu / (dim === 'w' ? info.widthEmu : info.heightEmu);
    const keep = info.lockAspect;
    applyToSelected(session, commands.resizeImageCommand, {
      cxEmu: dim === 'w' ? emu : keep ? info.widthEmu * k : info.widthEmu,
      cyEmu: dim === 'h' ? emu : keep ? info.heightEmu * k : info.heightEmu,
    });
  }
</script>

<div class="pane-head">
  <span>{t('pic.formatPane')}</span>
  <button class="pane-close" onclick={() => (session.pane.right = null)} aria-label={t('pic.closePane')}><RibbonIcon name="close" size={14} /></button>
</div>
<div class="fmt-tabs" role="tablist">
  {#each SECTIONS as [id, key] (id)}
    <button role="tab" aria-selected={section === id} class:active={section === id} onclick={() => (section = id)}>{t(key)}</button>
  {/each}
</div>
{#if !info || !picture}
  <p class="pane-note">{t('altText.none')}</p>
{:else if section === 'line'}
  <label class="field"><input type="radio" name="line" checked={!picture.outline} onchange={() => applyToSelected(session, commands.pictureBorderCommand, { outline: undefined })} />{t('pic.noOutline')}</label>
  <label class="field"><input type="radio" name="line" checked={!!picture.outline} onchange={() => setLine({})} />{t('fmtPic.solidLine')}</label>
  <label class="field">{t('fmtPic.color')} <input type="color" value="#{picture.outline?.color ?? '000000'}" onchange={(e) => setLine({ color: e.currentTarget.value.slice(1).toUpperCase() })} /></label>
  <label class="field">{t('pic.weight')} <input type="number" min="0" step="0.25" value={pt(picture.outline?.widthEmu ?? 0)} onchange={(e) => { const v = num(e); if (v !== undefined) setLine({ widthEmu: v * EMU_PER_PT }); }} /> pt</label>
  <label class="field">{t('pic.dashes')}
    <select value={picture.outline?.dash ?? 'solid'} onchange={(e) => setLine({ dash: BORDER_DASHES.find((d) => d === e.currentTarget.value) ?? 'solid' })}>
      {#each BORDER_DASHES as d (d)}<option value={d}>{d}</option>{/each}
    </select>
  </label>
{:else if section === 'effects'}
  {@const fx = picture.effects}
  <fieldset>
    <legend>{t('pic.shadow')}</legend>
    <label class="field"><input type="checkbox" checked={!!fx.shadow} onchange={(e) => (e.currentTarget.checked ? setShadow({}) : patchEffects(session, { shadow: undefined }))} />{t('pic.shadow')}</label>
    {#if fx.shadow}
      <label class="field">{t('fmtPic.color')} <input type="color" value="#{fx.shadow.color}" onchange={(e) => setShadow({ color: e.currentTarget.value.slice(1).toUpperCase() })} /></label>
      <label class="field">{t('pic.transparency')} <input type="number" min="0" max="100" value={100 - fx.shadow.opacity} onchange={(e) => { const v = num(e); if (v !== undefined) setShadow({ opacity: 100 - v }); }} /> %</label>
      <label class="field">{t('fmtPic.blur')} <input type="number" min="0" value={pt(fx.shadow.blurEmu)} onchange={(e) => { const v = num(e); if (v !== undefined) setShadow({ blurEmu: v * EMU_PER_PT }); }} /> pt</label>
      <label class="field">{t('fmtPic.angle')} <input type="number" min="0" max="359" value={fx.shadow.directionDeg} onchange={(e) => { const v = num(e); if (v !== undefined) setShadow({ directionDeg: v }); }} />°</label>
      <label class="field">{t('fmtPic.distance')} <input type="number" min="0" value={pt(fx.shadow.distanceEmu)} onchange={(e) => { const v = num(e); if (v !== undefined) setShadow({ distanceEmu: v * EMU_PER_PT }); }} /> pt</label>
    {/if}
  </fieldset>
  <fieldset>
    <legend>{t('pic.glow')}</legend>
    <label class="field">{t('fmtPic.size')} <input type="number" min="0" value={pt(fx.glow?.radiusEmu ?? 0)} onchange={(e) => { const v = num(e); if (v !== undefined) patchEffects(session, { glow: v ? { color: fx.glow?.color ?? '4472C4', opacity: fx.glow?.opacity ?? 60, radiusEmu: v * EMU_PER_PT } : undefined }); }} /> pt</label>
    {#if fx.glow}
      <label class="field">{t('fmtPic.color')} <input type="color" value="#{fx.glow.color}" onchange={(e) => fx.glow && patchEffects(session, { glow: { ...fx.glow, color: e.currentTarget.value.slice(1).toUpperCase() } })} /></label>
    {/if}
  </fieldset>
  <fieldset>
    <legend>{t('pic.softEdges')}</legend>
    <label class="field">{t('fmtPic.size')} <input type="number" min="0" value={pt(fx.softEdgeEmu ?? 0)} onchange={(e) => { const v = num(e); if (v !== undefined) patchEffects(session, { softEdgeEmu: v ? v * EMU_PER_PT : undefined }); }} /> pt</label>
  </fieldset>
  <fieldset>
    <legend>{t('pic.rotation3d')}</legend>
    {#each [['latitude', 'X'], ['longitude', 'Y'], ['revolution', 'Z']] as const as [axis, label] (axis)}
      <label class="field">{label} <input type="number" min="0" max="359" value={fx.rotation3d?.[axis] ?? 0}
        onchange={(e) => { const v = num(e); if (v !== undefined) patchEffects(session, { rotation3d: { camera: fx.rotation3d?.camera ?? 'orthographicFront', ...fx.rotation3d, [axis]: v } }); }} />°</label>
    {/each}
  </fieldset>
{:else if section === 'size'}
  <label class="field">{t('pic.height')} <input type="number" min="0.01" step="0.1" value={cm(info.heightEmu)} onchange={(e) => setSize('h', num(e))} /> cm</label>
  <label class="field">{t('pic.width')} <input type="number" min="0.01" step="0.1" value={cm(info.widthEmu)} onchange={(e) => setSize('w', num(e))} /> cm</label>
  <label class="field">{t('objLayout.rotation')} <input type="number" step="1" value={info.rotation} onchange={(e) => { const v = num(e); if (v !== undefined) applyToSelected(session, commands.pictureTransformCommand, { rotation: ((v % 360) + 360) % 360 }); }} />°</label>
  <label class="field"><input type="checkbox" checked={info.lockAspect} onchange={(e) => applyToSelected(session, commands.drawingAspectLockCommand, { locked: e.currentTarget.checked })} />{t('pic.lockAspect')}</label>
{:else}
  {@const a = picture.adjustments}
  <fieldset>
    <legend>{t('pic.corrections')}</legend>
    <label class="field">{t('pic.brightness')} <input type="range" min="-100" max="100" value={a.brightness ?? 0} onchange={(e) => patchAdjustments(session, { brightness: num(e) || undefined })} /> {a.brightness ?? 0}%</label>
    <label class="field">{t('pic.contrast')} <input type="range" min="-100" max="100" value={a.contrast ?? 0} onchange={(e) => patchAdjustments(session, { contrast: num(e) || undefined })} /> {a.contrast ?? 0}%</label>
  </fieldset>
  <fieldset>
    <legend>{t('pic.color')}</legend>
    <label class="field">{t('pic.saturation')} <input type="range" min="0" max="400" value={a.saturation ?? 100} onchange={(e) => { const v = num(e); patchAdjustments(session, { saturation: v === 100 ? undefined : v }); }} /> {a.saturation ?? 100}%</label>
  </fieldset>
  <fieldset>
    <legend>{t('pic.transparency')}</legend>
    <label class="field">{t('pic.transparency')} <input type="range" min="0" max="100" value={a.transparency ?? 0} onchange={(e) => patchAdjustments(session, { transparency: num(e) || undefined })} /> {a.transparency ?? 0}%</label>
  </fieldset>
  <fieldset>
    <legend>{t('pic.crop')}</legend>
    {#each [['left', 'objLayout.left'], ['top', 'objLayout.top'], ['right', 'objLayout.right'], ['bottom', 'objLayout.bottom']] as const as [side, key] (side)}
      <label class="field">{t(key)} <input type="number" step="0.1" value={Math.round(picture.crop[side] * 10) / 10} onchange={(e) => setCrop(side, num(e))} /> %</label>
    {/each}
  </fieldset>
{/if}

<style>
  .fmt-tabs {
    display: flex;
    gap: 2px;
  }
  .fmt-tabs button {
    flex: 1;
    padding: 4px 2px;
    border: 1px solid var(--chrome-line);
    border-radius: 5px;
    background: #fff;
    font-size: 11px;
  }
  .fmt-tabs button.active {
    background: var(--accent, #2b579a);
    color: #fff;
  }
  fieldset {
    display: flex;
    flex-direction: column;
    gap: 4px;
  }
  input[type='number'] {
    width: 60px;
  }
</style>
