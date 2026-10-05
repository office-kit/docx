<script lang="ts">
  /**
   * Word's Layout dialog (Wrap Text ▸ More Layout Options…, Size ▸ dialog
   * launcher): Position, Text Wrapping and Size tabs for the selected picture
   * or chart. Opened as `picture.layout` (Position tab) or `picture.size`.
   */
  import { commands } from '@office-kit/docx-editor';
  import type {
    HorizontalAlign,
    HorizontalPosition,
    HorizontalRelativeTo,
    VerticalAlign,
    VerticalPosition,
    VerticalRelativeTo,
    WrapSide,
    WrapStyle,
  } from '@office-kit/docx';
  import { untrack } from 'svelte';
  import Dialog from '../Dialog.svelte';
  import { getSession } from '../session.svelte';
  import { t, type MessageKey } from '../i18n/index.svelte';
  import { applyToSelected, EMU_PER_CM, selected } from './state';

  const session = getSession();
  type Tab = 'position' | 'wrapping' | 'size';
  let tab = $state<Tab>('position');

  const H_REL: ReadonlyArray<readonly [HorizontalRelativeTo, MessageKey]> = [
    ['margin', 'objLayout.relMargin'],
    ['page', 'objLayout.relPage'],
    ['column', 'objLayout.relColumn'],
    ['character', 'objLayout.relCharacter'],
    ['leftMargin', 'objLayout.relLeftMargin'],
    ['rightMargin', 'objLayout.relRightMargin'],
    ['insideMargin', 'objLayout.relInsideMargin'],
    ['outsideMargin', 'objLayout.relOutsideMargin'],
  ];
  const V_REL: ReadonlyArray<readonly [VerticalRelativeTo, MessageKey]> = [
    ['margin', 'objLayout.relMargin'],
    ['page', 'objLayout.relPage'],
    ['paragraph', 'objLayout.relParagraph'],
    ['line', 'objLayout.relLine'],
    ['topMargin', 'objLayout.relTopMargin'],
    ['bottomMargin', 'objLayout.relBottomMargin'],
    ['insideMargin', 'objLayout.relInsideMargin'],
    ['outsideMargin', 'objLayout.relOutsideMargin'],
  ];
  const H_ALIGN: ReadonlyArray<readonly [HorizontalAlign, MessageKey]> = [
    ['left', 'objLayout.left'],
    ['center', 'objLayout.centered'],
    ['right', 'objLayout.right'],
  ];
  const V_ALIGN: ReadonlyArray<readonly [VerticalAlign, MessageKey]> = [
    ['top', 'objLayout.top'],
    ['center', 'objLayout.centered'],
    ['bottom', 'objLayout.bottom'],
  ];
  const WRAPS: ReadonlyArray<readonly [WrapStyle, MessageKey]> = [
    ['inline', 'arr.wrapInline'],
    ['square', 'arr.wrapSquare'],
    ['tight', 'arr.wrapTight'],
    ['through', 'arr.wrapThrough'],
    ['topAndBottom', 'arr.wrapTopBottom'],
    ['behindText', 'arr.wrapBehind'],
    ['inFrontOfText', 'arr.wrapInFront'],
  ];
  const SIDES: ReadonlyArray<readonly [WrapSide, MessageKey]> = [
    ['bothSides', 'objLayout.bothSides'],
    ['left', 'objLayout.leftOnly'],
    ['right', 'objLayout.rightOnly'],
    ['largest', 'objLayout.largestOnly'],
  ];

  interface Form {
    hMode: 'align' | 'offset';
    hAlign: HorizontalAlign;
    hAlignRel: HorizontalRelativeTo;
    hOffset: number;
    hOffsetRel: HorizontalRelativeTo;
    vMode: 'align' | 'offset';
    vAlign: VerticalAlign;
    vAlignRel: VerticalRelativeTo;
    vOffset: number;
    vOffsetRel: VerticalRelativeTo;
    locked: boolean;
    allowOverlap: boolean;
    layoutInCell: boolean;
    wrap: WrapStyle;
    side: WrapSide;
    distT: number;
    distB: number;
    distL: number;
    distR: number;
    height: number;
    width: number;
    scaleH: number;
    scaleW: number;
    rotation: number;
    lockAspect: boolean;
  }

  let form = $state<Form | null>(null);
  // The size the form was opened with, so Scale edits size relative to it.
  let base = { width: 1, height: 1 };

  const cm = (emu: number): number => Math.round((emu / EMU_PER_CM) * 100) / 100;
  const emu = (cmValue: number): number => Math.round(cmValue * EMU_PER_CM);

  function horizontalFields(p: HorizontalPosition | undefined) {
    if (!p) return { hMode: 'align' as const, hAlign: 'left' as const, hAlignRel: 'column' as const, hOffset: 0, hOffsetRel: 'column' as const };
    return 'align' in p
      ? { hMode: 'align' as const, hAlign: p.align, hAlignRel: p.relativeTo, hOffset: 0, hOffsetRel: p.relativeTo }
      : { hMode: 'offset' as const, hAlign: 'left' as const, hAlignRel: p.relativeTo, hOffset: cm(p.offsetEmu), hOffsetRel: p.relativeTo };
  }

  function verticalFields(p: VerticalPosition | undefined) {
    if (!p) return { vMode: 'offset' as const, vAlign: 'top' as const, vAlignRel: 'paragraph' as const, vOffset: 0, vOffsetRel: 'paragraph' as const };
    return 'align' in p
      ? { vMode: 'align' as const, vAlign: p.align, vAlignRel: p.relativeTo, vOffset: 0, vOffsetRel: p.relativeTo }
      : { vMode: 'offset' as const, vAlign: 'top' as const, vAlignRel: p.relativeTo, vOffset: cm(p.offsetEmu), vOffsetRel: p.relativeTo };
  }

  // Load the selection into the form each time the dialog opens.
  $effect(() => {
    const id = session.dialog;
    if (id !== 'picture.layout' && id !== 'picture.size') return;
    // Only opening the dialog loads it; later ticks must not reset the user's edits.
    const sel = untrack(() => selected(session));
    if (!sel) return;
    tab = id === 'picture.size' ? 'size' : 'position';
    const { info } = sel;
    const a = info.anchor;
    base = { width: info.widthEmu, height: info.heightEmu };
    form = {
      ...horizontalFields(a?.horizontal),
      ...verticalFields(a?.vertical),
      locked: a?.locked ?? false,
      allowOverlap: a?.allowOverlap ?? true,
      layoutInCell: a?.layoutInCell ?? true,
      wrap: info.wrap,
      side: a?.wrapSide ?? 'bothSides',
      distT: cm(a?.distance.top ?? 0),
      distB: cm(a?.distance.bottom ?? 0),
      distL: cm(a?.distance.left ?? 0),
      distR: cm(a?.distance.right ?? 0),
      height: cm(info.heightEmu),
      width: cm(info.widthEmu),
      scaleH: 100,
      scaleW: 100,
      rotation: info.rotation,
      lockAspect: info.lockAspect,
    };
  });

  /** Height / width / scale edits keep the aspect ratio when it is locked. */
  function setHeight(v: number): void {
    if (!form) return;
    form.height = v;
    form.scaleH = Math.round((emu(v) / base.height) * 100);
    if (form.lockAspect) {
      form.scaleW = form.scaleH;
      form.width = cm((base.width * form.scaleH) / 100);
    }
  }
  function setWidth(v: number): void {
    if (!form) return;
    form.width = v;
    form.scaleW = Math.round((emu(v) / base.width) * 100);
    if (form.lockAspect) {
      form.scaleH = form.scaleW;
      form.height = cm((base.height * form.scaleW) / 100);
    }
  }
  function setScaleH(v: number): void {
    setHeight(cm((base.height * v) / 100));
  }
  function setScaleW(v: number): void {
    setWidth(cm((base.width * v) / 100));
  }

  function ok(): boolean {
    const f = form;
    if (!f) return true;
    if (!(f.width > 0 && f.height > 0)) return false;
    const horizontal: HorizontalPosition =
      f.hMode === 'align' ? { relativeTo: f.hAlignRel, align: f.hAlign } : { relativeTo: f.hOffsetRel, offsetEmu: emu(f.hOffset) };
    const vertical: VerticalPosition =
      f.vMode === 'align' ? { relativeTo: f.vAlignRel, align: f.vAlign } : { relativeTo: f.vOffsetRel, offsetEmu: emu(f.vOffset) };
    applyToSelected(session, commands.drawingLayoutCommand, {
      wrap: f.wrap,
      position: { horizontal, vertical },
      options: {
        wrapSide: f.side,
        distance: { top: emu(f.distT), bottom: emu(f.distB), left: emu(f.distL), right: emu(f.distR) },
        locked: f.locked,
        allowOverlap: f.allowOverlap,
        layoutInCell: f.layoutInCell,
      },
      size: { cxEmu: emu(f.width), cyEmu: emu(f.height) },
      rotation: ((f.rotation % 360) + 360) % 360,
      lockAspect: f.lockAspect,
    });
    return true;
  }

  const floating = $derived(form ? form.wrap !== 'inline' : false);
  const dialogId = $derived(session.dialog === 'picture.size' ? 'picture.size' : 'picture.layout');
  const TABS: ReadonlyArray<readonly [Tab, MessageKey]> = [
    ['position', 'objLayout.tabPosition'],
    ['wrapping', 'objLayout.tabWrapping'],
    ['size', 'objLayout.tabSize'],
  ];
</script>

<Dialog id={dialogId} title={t('objLayout.title')} onok={ok}>
  {#if form}
    <div class="lay-tabs" role="tablist">
      {#each TABS as [id, key] (id)}
        <button type="button" role="tab" aria-selected={tab === id} class:active={tab === id} onclick={() => (tab = id)}>{t(key)}</button>
      {/each}
    </div>
    {#if tab === 'position'}
      <fieldset disabled={!floating}>
        <legend>{t('objLayout.horizontal')}</legend>
        <label class="field"><input type="radio" bind:group={form.hMode} value="align" />{t('objLayout.alignment')}
          <select bind:value={form.hAlign}>{#each H_ALIGN as [v, k] (v)}<option value={v}>{t(k)}</option>{/each}</select>
          {t('objLayout.relativeTo')}
          <select bind:value={form.hAlignRel}>{#each H_REL as [v, k] (v)}<option value={v}>{t(k)}</option>{/each}</select>
        </label>
        <label class="field"><input type="radio" bind:group={form.hMode} value="offset" />{t('objLayout.absolute')}
          <input type="number" step="0.01" bind:value={form.hOffset} /> cm {t('objLayout.rightOf')}
          <select bind:value={form.hOffsetRel}>{#each H_REL as [v, k] (v)}<option value={v}>{t(k)}</option>{/each}</select>
        </label>
      </fieldset>
      <fieldset disabled={!floating}>
        <legend>{t('objLayout.vertical')}</legend>
        <label class="field"><input type="radio" bind:group={form.vMode} value="align" />{t('objLayout.alignment')}
          <select bind:value={form.vAlign}>{#each V_ALIGN as [v, k] (v)}<option value={v}>{t(k)}</option>{/each}</select>
          {t('objLayout.relativeTo')}
          <select bind:value={form.vAlignRel}>{#each V_REL as [v, k] (v)}<option value={v}>{t(k)}</option>{/each}</select>
        </label>
        <label class="field"><input type="radio" bind:group={form.vMode} value="offset" />{t('objLayout.absolute')}
          <input type="number" step="0.01" bind:value={form.vOffset} /> cm {t('objLayout.below')}
          <select bind:value={form.vOffsetRel}>{#each V_REL as [v, k] (v)}<option value={v}>{t(k)}</option>{/each}</select>
        </label>
      </fieldset>
      <fieldset disabled={!floating}>
        <legend>{t('objLayout.options')}</legend>
        <label class="field"><input type="checkbox" checked={form.vOffsetRel === 'paragraph'} onchange={(e) => { if (form) { form.vMode = 'offset'; form.vOffsetRel = e.currentTarget.checked ? 'paragraph' : 'page'; } }} />{t('arr.moveWithText')}</label>
        <label class="field"><input type="checkbox" bind:checked={form.locked} />{t('objLayout.lockAnchor')}</label>
        <label class="field"><input type="checkbox" bind:checked={form.allowOverlap} />{t('objLayout.allowOverlap')}</label>
        <label class="field"><input type="checkbox" bind:checked={form.layoutInCell} />{t('objLayout.layoutInCell')}</label>
      </fieldset>
    {:else if tab === 'wrapping'}
      <fieldset>
        <legend>{t('objLayout.wrapStyle')}</legend>
        <div class="lay-grid">
          {#each WRAPS as [v, k] (v)}<label class="field"><input type="radio" bind:group={form.wrap} value={v} />{t(k)}</label>{/each}
        </div>
      </fieldset>
      <fieldset disabled={!floating || form.wrap === 'topAndBottom' || form.wrap === 'behindText' || form.wrap === 'inFrontOfText'}>
        <legend>{t('objLayout.wrapText')}</legend>
        <div class="lay-grid">
          {#each SIDES as [v, k] (v)}<label class="field"><input type="radio" bind:group={form.side} value={v} />{t(k)}</label>{/each}
        </div>
      </fieldset>
      <fieldset disabled={!floating}>
        <legend>{t('objLayout.distance')}</legend>
        <div class="lay-grid">
          <label class="field">{t('objLayout.top')} <input type="number" min="0" step="0.01" bind:value={form.distT} /> cm</label>
          <label class="field">{t('objLayout.left')} <input type="number" min="0" step="0.01" bind:value={form.distL} /> cm</label>
          <label class="field">{t('objLayout.bottom')} <input type="number" min="0" step="0.01" bind:value={form.distB} /> cm</label>
          <label class="field">{t('objLayout.right')} <input type="number" min="0" step="0.01" bind:value={form.distR} /> cm</label>
        </div>
      </fieldset>
    {:else}
      <fieldset>
        <legend>{t('pic.height')}</legend>
        <label class="field">{t('objLayout.absolute')} <input type="number" min="0.01" step="0.01" value={form.height} onchange={(e) => setHeight(e.currentTarget.valueAsNumber)} /> cm</label>
      </fieldset>
      <fieldset>
        <legend>{t('pic.width')}</legend>
        <label class="field">{t('objLayout.absolute')} <input type="number" min="0.01" step="0.01" value={form.width} onchange={(e) => setWidth(e.currentTarget.valueAsNumber)} /> cm</label>
      </fieldset>
      <fieldset>
        <legend>{t('objLayout.rotate')}</legend>
        <label class="field">{t('objLayout.rotation')} <input type="number" step="1" bind:value={form.rotation} />°</label>
      </fieldset>
      <fieldset>
        <legend>{t('objLayout.scale')}</legend>
        <label class="field">{t('pic.height')} <input type="number" min="1" step="1" value={form.scaleH} onchange={(e) => setScaleH(e.currentTarget.valueAsNumber)} /> %</label>
        <label class="field">{t('pic.width')} <input type="number" min="1" step="1" value={form.scaleW} onchange={(e) => setScaleW(e.currentTarget.valueAsNumber)} /> %</label>
        <label class="field"><input type="checkbox" bind:checked={form.lockAspect} />{t('pic.lockAspect')}</label>
      </fieldset>
    {/if}
  {/if}
</Dialog>

<style>
  .lay-tabs {
    display: flex;
    justify-content: center;
    gap: 2px;
    margin-bottom: 8px;
  }
  .lay-tabs button {
    padding: 3px 12px;
    border: 1px solid var(--chrome-line);
    border-radius: 5px;
    background: #fff;
  }
  .lay-tabs button.active {
    background: var(--accent, #2b579a);
    color: #fff;
  }
  .lay-grid {
    display: grid;
    grid-template-columns: repeat(2, auto);
    gap: 4px 16px;
  }
  fieldset {
    display: flex;
    flex-direction: column;
    gap: 4px;
    margin-bottom: 8px;
  }
  input[type='number'] {
    width: 64px;
  }
</style>
