<script lang="ts">
  /**
   * Table Properties: the Table, Row, Column, Cell and Alt Text tabs. Opens on
   * the selection's current settings; on OK only what changed is written, so
   * opening and confirming the dialog does not stamp widths on every cell.
   */
  import type { TableAlignment, TableCellVerticalAlign, TableWidth } from '@office-kit/docx';
  import { commands, tablePropertiesSnapshot, type TablePropertiesSnapshot } from '@office-kit/docx-editor';
  import Dialog from '../../Dialog.svelte';
  import { getSession } from '../../session.svelte';
  import { t, type MessageKey } from '../../i18n/index.svelte';
  import LengthField from './LengthField.svelte';
  import { DIALOGS } from './lines';
  import { lengthUnit } from './table-tool.svelte';

  const session = getSession();
  type Tab = 'table' | 'row' | 'column' | 'cell' | 'alt';
  const TABS: readonly { id: Tab; label: MessageKey }[] = [
    { id: 'table', label: 'tbl.dlg.tabTable' },
    { id: 'row', label: 'tbl.dlg.tabRow' },
    { id: 'column', label: 'tbl.dlg.tabColumn' },
    { id: 'cell', label: 'tbl.dlg.tabCell' },
    { id: 'alt', label: 'tbl.dlg.tabAlt' },
  ];
  // pct widths are fiftieths of a percent (ST_TblWidth).
  const PCT_SCALE = 50;
  // Word's suggestion when a preferred width is first turned on: 5″ for a table, 1″ for a cell.
  const DEFAULT_TABLE_WIDTH = 7200;
  const DEFAULT_CELL_WIDTH = 1440;
  const DEFAULT_ROW_HEIGHT = 360;

  let tab = $state<Tab>('table');
  let snap = $state<TablePropertiesSnapshot | undefined>(undefined);

  // Editable copies of the snapshot.
  let tableWidthOn = $state(false);
  let tableWidth = $state<TableWidth>({ type: 'dxa', value: DEFAULT_TABLE_WIDTH });
  let alignment = $state<TableAlignment>('left');
  let indent = $state(0);
  let heightOn = $state(false);
  let height = $state(DEFAULT_ROW_HEIGHT);
  let heightRule = $state<'atLeast' | 'exact'>('atLeast');
  let allowBreak = $state(true);
  let header = $state(false);
  let columnWidth = $state(0);
  let cellWidthOn = $state(false);
  let cellWidth = $state<TableWidth>({ type: 'dxa', value: DEFAULT_CELL_WIDTH });
  let vAlign = $state<TableCellVerticalAlign>('top');
  let noWrap = $state(false);
  let fitText = $state(false);
  let title = $state('');
  let description = $state('');

  const opened = $derived(session.dialog === DIALOGS.properties);
  $effect(() => {
    if (!opened || !session.model) return;
    const s = tablePropertiesSnapshot(session.model);
    snap = s;
    if (!s) return;
    tableWidthOn = !!s.table.width;
    tableWidth = s.table.width ?? { type: 'dxa', value: DEFAULT_TABLE_WIDTH };
    alignment = s.table.alignment;
    indent = s.table.indentTwips;
    heightOn = !!s.row.height;
    height = s.row.height?.twips ?? DEFAULT_ROW_HEIGHT;
    heightRule = s.row.height?.rule === 'exact' ? 'exact' : 'atLeast';
    allowBreak = !s.row.cantSplit;
    header = s.row.header;
    columnWidth = s.column.widthTwips;
    cellWidthOn = !!s.cell.width;
    cellWidth = s.cell.width ?? { type: 'dxa', value: DEFAULT_CELL_WIDTH };
    vAlign = s.cell.verticalAlign;
    noWrap = s.cell.noWrap;
    fitText = s.cell.fitText;
    title = s.altText.title;
    description = s.altText.description;
  });

  const sameWidth = (a: TableWidth | undefined, b: TableWidth | undefined): boolean => a?.type === b?.type && a?.value === b?.value;

  function ok(): void {
    const s = snap;
    if (!s) return;
    const newTableWidth = tableWidthOn ? tableWidth : undefined;
    const newCellWidth = cellWidthOn ? cellWidth : undefined;
    const newHeight = heightOn ? { twips: height, rule: heightRule } : undefined;
    const params: commands.TablePropertiesParams = {
      table: {
        ...(sameWidth(newTableWidth, s.table.width) ? {} : { width: newTableWidth ?? null }),
        ...(alignment === s.table.alignment ? {} : { alignment }),
        ...(indent === s.table.indentTwips ? {} : { indentTwips: indent }),
      },
      rows: {
        ...(newHeight?.twips === s.row.height?.twips && newHeight?.rule === s.row.height?.rule ? {} : { height: newHeight ?? null }),
        ...(allowBreak === !s.row.cantSplit ? {} : { cantSplit: !allowBreak }),
        ...(header === s.row.header ? {} : { header }),
      },
      ...(columnWidth === s.column.widthTwips ? {} : { columnWidthTwips: columnWidth }),
      cells: {
        ...(sameWidth(newCellWidth, s.cell.width) ? {} : { width: newCellWidth ?? null }),
        ...(vAlign === s.cell.verticalAlign ? {} : { verticalAlign: vAlign }),
        ...(noWrap === s.cell.noWrap ? {} : { noWrap }),
        ...(fitText === s.cell.fitText ? {} : { fitText }),
      },
      ...(title === s.altText.title && description === s.altText.description ? {} : { altText: { title, description } }),
    };
    session.apply(commands.tablePropertiesCommand, params);
  }

  function widthValue(w: TableWidth): number {
    return w.type === 'pct' ? w.value / PCT_SCALE : w.value;
  }
  function pctWidth(pct: number): TableWidth {
    return { type: 'pct', value: Math.round(pct * PCT_SCALE) };
  }
  function switchMeasure(w: TableWidth, type: string): TableWidth {
    // Switching to percent starts at 100 %, back to a length at Word's default.
    if (type === 'pct') return w.type === 'pct' ? w : { type: 'pct', value: 100 * PCT_SCALE };
    return w.type === 'dxa' ? w : { type: 'dxa', value: DEFAULT_CELL_WIDTH };
  }
</script>

{#snippet widthFields(on: boolean, w: TableWidth, toggle: (v: boolean) => void, set: (w: TableWidth) => void)}
  <div class="tp-row">
    <label class="field"><input type="checkbox" checked={on} onchange={(e) => toggle(e.currentTarget.checked)} />{t('tbl.dlg.preferredWidth')}</label>
    {#if w.type === 'pct'}
      <label class="field"><input type="number" min="0" max="600" step="any" disabled={!on} value={widthValue(w)} onchange={(e) => set(pctWidth(Number(e.currentTarget.value)))} />%</label>
    {:else}
      <LengthField label="" twips={w.value} disabled={!on} onchange={(value) => set({ type: 'dxa', value })} />
    {/if}
    <label class="field">{t('tbl.dlg.measureIn')}
      <select disabled={!on} value={w.type === 'pct' ? 'pct' : 'dxa'} onchange={(e) => set(switchMeasure(w, e.currentTarget.value))}>
        <option value="dxa">{t(lengthUnit() === 'in' ? 'tbl.unit.inches' : lengthUnit() === 'mm' ? 'tbl.unit.millimeters' : 'tbl.unit.centimeters')}</option>
        <option value="pct">{t('tbl.unit.pct')}</option>
      </select>
    </label>
  </div>
{/snippet}

<Dialog id={DIALOGS.properties} title={t('tbl.dlg.properties')} onok={ok}>
  <div class="tp-tabs" role="tablist">
    {#each TABS as tb (tb.id)}
      <button type="button" role="tab" class:active={tab === tb.id} aria-selected={tab === tb.id} onclick={() => (tab = tb.id)}>{t(tb.label)}</button>
    {/each}
  </div>
  {#if tab === 'table'}
    <fieldset>
      <legend>{t('tbl.dlg.size')}</legend>
      {@render widthFields(tableWidthOn, tableWidth, (v) => (tableWidthOn = v), (w) => (tableWidth = w))}
    </fieldset>
    <fieldset>
      <legend>{t('tbl.dlg.alignment')}</legend>
      <div class="tp-row">
        {#each [['left', 'tbl.dlg.alignLeft'], ['center', 'tbl.dlg.alignCenter'], ['right', 'tbl.dlg.alignRight']] as const as [value, label] (value)}
          <label class="field"><input type="radio" name="tp-align" checked={alignment === value} onchange={() => (alignment = value)} />{t(label)}</label>
        {/each}
        <LengthField label={t('tbl.dlg.indent')} bind:twips={indent} disabled={alignment !== 'left'} />
      </div>
    </fieldset>
    <div class="tp-row end">
      <button type="button" class="push" onclick={() => session.openDialog(DIALOGS.cellMargins)}>{t('tbl.dlg.optionsButton')}</button>
      <button type="button" class="push" onclick={() => session.openDialog(DIALOGS.bordersShading)}>{t('tbl.border.dialog')}</button>
    </div>
  {:else if tab === 'row'}
    <fieldset>
      <legend>{t('tbl.dlg.size')}</legend>
      <div class="tp-row">
        <label class="field"><input type="checkbox" bind:checked={heightOn} />{t('tbl.dlg.specifyHeight')}</label>
        <LengthField label="" bind:twips={height} disabled={!heightOn} />
        <label class="field">{t('tbl.dlg.heightRule')}
          <select bind:value={heightRule} disabled={!heightOn}>
            <option value="atLeast">{t('tbl.dlg.atLeast')}</option>
            <option value="exact">{t('tbl.dlg.exactly')}</option>
          </select>
        </label>
      </div>
    </fieldset>
    <label class="field"><input type="checkbox" bind:checked={allowBreak} />{t('tbl.dlg.allowBreak')}</label>
    <label class="field"><input type="checkbox" bind:checked={header} />{t('tbl.dlg.repeatHeader')}</label>
  {:else if tab === 'column'}
    <fieldset>
      <legend>{t('tbl.dlg.size')}</legend>
      <LengthField label={t('tbl.dlg.preferredWidth')} bind:twips={columnWidth} />
    </fieldset>
  {:else if tab === 'cell'}
    <fieldset>
      <legend>{t('tbl.dlg.size')}</legend>
      {@render widthFields(cellWidthOn, cellWidth, (v) => (cellWidthOn = v), (w) => (cellWidth = w))}
    </fieldset>
    <fieldset>
      <legend>{t('tbl.dlg.verticalAlignment')}</legend>
      <div class="tp-row">
        {#each [['top', 'tbl.dlg.vTop'], ['center', 'tbl.dlg.vCenter'], ['bottom', 'tbl.dlg.vBottom']] as const as [value, label] (value)}
          <label class="field"><input type="radio" name="tp-valign" checked={vAlign === value} onchange={() => (vAlign = value)} />{t(label)}</label>
        {/each}
      </div>
    </fieldset>
    <div class="tp-row">
      <label class="field"><input type="checkbox" checked={!noWrap} onchange={(e) => (noWrap = !e.currentTarget.checked)} />{t('tbl.dlg.wrapText')}</label>
      <label class="field"><input type="checkbox" bind:checked={fitText} />{t('tbl.dlg.fitText')}</label>
    </div>
  {:else}
    <label class="field">{t('tbl.dlg.title')}<input type="text" bind:value={title} /></label>
    <label class="tp-col">{t('tbl.dlg.description')}<textarea rows="4" bind:value={description}></textarea></label>
  {/if}
</Dialog>

<style>
  .tp-tabs {
    display: flex;
    justify-content: center;
    gap: 2px;
  }
  .tp-tabs button {
    padding: 3px 12px;
    border: 1px solid var(--control-line);
    border-radius: 5px;
    background: #fff;
    cursor: pointer;
  }
  .tp-tabs button.active {
    background: #2f6fe0;
    border-color: #2f6fe0;
    color: #fff;
  }
  .tp-row {
    display: flex;
    flex-wrap: wrap;
    align-items: center;
    gap: 10px;
  }
  .tp-row.end {
    justify-content: flex-end;
  }
  .tp-col {
    display: flex;
    flex-direction: column;
    gap: 4px;
  }
</style>
