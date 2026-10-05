<script lang="ts">
  /**
   * Modify Table Style and New Table Style (Word's "Create New Style from
   * Formatting" for tables): pick the region the formatting applies to (whole
   * table, header row, banded rows …) and set its text, fill, borders and
   * alignment. Only the fields the user touched are written.
   */
  import { BUILT_IN_TABLE_STYLES, listStyles, type TableBorder, type TableStyleFormatting, type TableStyleRegion } from '@office-kit/docx';
  import { commands } from '@office-kit/docx-editor';
  import Dialog from '../../Dialog.svelte';
  import { getSession } from '../../session.svelte';
  import { t, type MessageKey } from '../../i18n/index.svelte';
  import { tableTool } from './table-tool.svelte';

  type Props = { mode: 'new' | 'modify'; dialogId: string };
  const { mode, dialogId }: Props = $props();
  const session = getSession();

  const REGIONS: readonly TableStyleRegion[] = ['wholeTable', 'firstRow', 'lastRow', 'firstCol', 'lastCol', 'band1Horz', 'band2Horz', 'band1Vert', 'band2Vert', 'nwCell', 'neCell', 'swCell', 'seCell'];
  const regionLabel = (r: TableStyleRegion): MessageKey => `tbl.region.${r}`;
  type BorderPreset = 'keep' | 'none' | 'outside' | 'all' | 'bottom' | 'insideH';

  const current = $derived(session.tick >= 0 && session.model ? commands.currentTableStyle(session.model) : undefined);
  const styleChoices = $derived.by(() => {
    if (!session.model || session.version < 0) return [];
    const builtIn = BUILT_IN_TABLE_STYLES.map((s) => ({ id: s.styleId, name: s.name }));
    const known = new Set(builtIn.map((s) => s.id));
    const custom = listStyles(session.model.doc)
      .filter((s) => s.type === 'table' && !known.has(s.styleId))
      .map((s) => ({ id: s.styleId, name: s.styleId }));
    return [...custom, ...builtIn];
  });

  let name = $state('');
  let basedOn = $state('TableGrid');
  let region = $state<TableStyleRegion>('wholeTable');
  let bold = $state<boolean | undefined>(undefined);
  let italic = $state<boolean | undefined>(undefined);
  let color = $state<string | undefined>(undefined);
  let fill = $state<string | undefined>(undefined);
  let alignment = $state<TableStyleFormatting['alignment'] | ''>('');
  let borders = $state<BorderPreset>('keep');

  const opened = $derived(session.dialog === dialogId);
  $effect(() => {
    if (!opened) return;
    name = mode === 'new' ? 'Table Style 1' : (current ?? '');
    basedOn = current ?? 'TableGrid';
    region = 'wholeTable';
    bold = italic = color = fill = undefined;
    alignment = '';
    borders = 'keep';
  });

  function borderFormatting(pen: TableBorder): TableStyleFormatting['borders'] {
    switch (borders) {
      case 'keep':
        return undefined;
      case 'none':
        return { top: null, left: null, bottom: null, right: null, insideH: null, insideV: null };
      case 'outside':
        return { top: pen, left: pen, bottom: pen, right: pen };
      case 'all':
        return { top: pen, left: pen, bottom: pen, right: pen, insideH: pen, insideV: pen };
      case 'bottom':
        return { bottom: pen };
      case 'insideH':
        return { insideH: pen };
    }
  }

  function formatting(): TableStyleFormatting {
    const b = borderFormatting(tableTool.pen);
    return {
      ...(bold === undefined ? {} : { bold }),
      ...(italic === undefined ? {} : { italic }),
      ...(color === undefined ? {} : { color }),
      ...(fill === undefined ? {} : { fill }),
      ...(alignment ? { alignment } : {}),
      ...(b ? { borders: b } : {}),
    };
  }

  function ok(): boolean {
    let styleId = current;
    if (mode === 'new') {
      styleId = session.apply(commands.newTableStyleCommand, { name, basedOn });
      if (styleId === undefined) return false;
    }
    if (styleId === undefined) return true;
    const f = formatting();
    if (Object.keys(f).length > 0) session.apply(commands.modifyTableStyleCommand, { styleId, region, formatting: f });
    return true;
  }

  const hex = (v: string | undefined): string => (v && v !== 'auto' ? `#${v}` : '#000000');
  const fromInput = (e: Event): string => (e.currentTarget instanceof HTMLInputElement ? e.currentTarget.value.slice(1).toUpperCase() : 'auto');
</script>

<Dialog id={dialogId} title={t(mode === 'new' ? 'tbl.dlg.newStyle' : 'tbl.dlg.modifyStyle')} onok={ok} okDisabled={mode === 'new' && !name.trim()}>
  <label class="field">{t('tbl.dlg.name')}<input type="text" bind:value={name} readonly={mode === 'modify'} /></label>
  {#if mode === 'new'}
    <label class="field">{t('tbl.dlg.basedOn')}
      <select bind:value={basedOn}>
        {#each styleChoices as s (s.id)}<option value={s.id}>{s.name}</option>{/each}
      </select>
    </label>
  {/if}
  <label class="field">{t('tbl.dlg.applyTo')}
    <select bind:value={region}>
      {#each REGIONS as r (r)}<option value={r}>{t(regionLabel(r))}</option>{/each}
    </select>
  </label>
  <fieldset>
    <div class="tbl-dlg-row">
      <label class="field"><input type="checkbox" checked={bold ?? false} onchange={(e) => (bold = e.currentTarget.checked)} />{t('tbl.dlg.bold')}</label>
      <label class="field"><input type="checkbox" checked={italic ?? false} onchange={(e) => (italic = e.currentTarget.checked)} />{t('tbl.dlg.italic')}</label>
      <label class="field">{t('tbl.dlg.fontColor')}<input type="color" value={hex(color)} onchange={(e) => (color = fromInput(e))} /></label>
    </div>
    <div class="tbl-dlg-row">
      <label class="field">{t('tbl.dlg.cellFill')}<input type="color" value={fill && fill !== 'auto' ? `#${fill}` : '#ffffff'} onchange={(e) => (fill = fromInput(e))} /></label>
      <button type="button" class="push" onclick={() => (fill = 'auto')}>{t('tbl.color.none')}</button>
    </div>
    <div class="tbl-dlg-row">
      <label class="field">{t('tbl.dlg.border')}
        <select bind:value={borders}>
          <option value="keep">—</option>
          <option value="none">{t('tbl.border.none')}</option>
          <option value="outside">{t('tbl.border.outside')}</option>
          <option value="all">{t('tbl.border.all')}</option>
          <option value="bottom">{t('tbl.border.bottom')}</option>
          <option value="insideH">{t('tbl.border.insideH')}</option>
        </select>
      </label>
      <label class="field">{t('tbl.dlg.alignment')}
        <select bind:value={alignment}>
          <option value="">—</option>
          <option value="left">{t('tbl.dlg.alignLeft')}</option>
          <option value="center">{t('tbl.dlg.alignCenter')}</option>
          <option value="right">{t('tbl.dlg.alignRight')}</option>
        </select>
      </label>
    </div>
  </fieldset>
</Dialog>

<style>
  .tbl-dlg-row {
    display: flex;
    flex-wrap: wrap;
    align-items: center;
    gap: 12px;
    margin: 4px 0;
  }
</style>
