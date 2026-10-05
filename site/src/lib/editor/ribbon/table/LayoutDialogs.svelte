<script lang="ts">
  /**
   * The Table Layout tab's smaller dialogs: Delete Cells, Split Cells, Table
   * Options (Cell Margins), Sort, Convert Table to Text and Formula.
   */
  import { getTableCellText, type TableSortKey, type TableTextSeparator } from '@office-kit/docx';
  import { commands, tablePropertiesSnapshot, tableSelection } from '@office-kit/docx-editor';
  import Dialog from '../../Dialog.svelte';
  import { getSession } from '../../session.svelte';
  import { t, type MessageKey } from '../../i18n/index.svelte';
  import LengthField from './LengthField.svelte';
  import SeparatorFields from './SeparatorFields.svelte';
  import { DIALOGS } from './lines';

  const session = getSession();
  const ts = $derived(session.tick >= 0 && session.model ? tableSelection(session.model) : undefined);

  // --- Delete Cells ---------------------------------------------------------------------
  type Shift = 'left' | 'up' | 'row' | 'column';
  const SHIFTS: readonly { value: Shift; label: MessageKey }[] = [
    { value: 'left', label: 'tbl.dlg.shiftLeft' },
    { value: 'up', label: 'tbl.dlg.shiftUp' },
    { value: 'row', label: 'tbl.dlg.entireRow' },
    { value: 'column', label: 'tbl.dlg.entireColumn' },
  ];
  let shift = $state<Shift>('left');

  // --- Split Cells ----------------------------------------------------------------------
  // Word proposes two columns and one row.
  let splitColumns = $state(2);
  let splitRows = $state(1);

  // --- Table Options --------------------------------------------------------------------
  let margins = $state({ top: 0, bottom: 0, left: 0, right: 0 });
  let spacingOn = $state(false);
  let spacing = $state(0);
  let autoResize = $state(true);
  $effect(() => {
    if (session.dialog !== DIALOGS.cellMargins || !session.model) return;
    const s = tablePropertiesSnapshot(session.model);
    if (!s) return;
    const m = s.options.defaultMargins;
    margins = { top: m.top ?? 0, bottom: m.bottom ?? 0, left: m.left ?? 0, right: m.right ?? 0 };
    spacingOn = s.options.cellSpacingTwips > 0;
    spacing = s.options.cellSpacingTwips;
    autoResize = s.options.autoResize;
  });

  // --- Sort -----------------------------------------------------------------------------
  // Word's Sort dialog has three keys: Sort by, Then by, Then by.
  const SORT_KEYS = 3;
  type KeyState = { column: number; type: NonNullable<TableSortKey['type']>; order: NonNullable<TableSortKey['order']> };
  let sortKeys = $state<KeyState[]>([]);
  let sortHeader = $state(false);
  const columnNames = $derived.by(() => {
    if (!ts) return [];
    const count = ts.placements[0]?.reduce((n, p) => n + p.gridSpan, 0) ?? 0;
    return Array.from({ length: count }, (_, i) => {
      const first = ts.placements[0]?.find((p) => p.gridStart === i);
      const text = sortHeader && first ? getTableCellText(ts.table, 0, first.cell).trim() : '';
      return text || `${t('tbl.dlg.column')} ${i + 1}`;
    });
  });
  $effect(() => {
    if (session.dialog !== DIALOGS.sort || !ts) return;
    sortHeader = ts.table.rows[0]?.trPr?.children.some((c) => c.kind === 'element' && c.name.local === 'tblHeader') ?? false;
    sortKeys = Array.from({ length: SORT_KEYS }, (_, i) => ({ column: i === 0 ? ts.range.firstColumn : -1, type: 'text', order: 'ascending' }));
  });
  function sort(): void {
    const keys = sortKeys.filter((k) => k.column >= 0).map((k) => ({ column: k.column, type: k.type, order: k.order }));
    if (keys.length > 0) session.apply(commands.sortTableCommand, { keys, headerRow: sortHeader });
  }

  // --- Convert to Text ------------------------------------------------------------------
  let toTextSeparator = $state<TableTextSeparator>('tab');

  // --- Formula --------------------------------------------------------------------------
  const FUNCTIONS = ['ABS', 'AND', 'AVERAGE', 'COUNT', 'DEFINED', 'FALSE', 'IF', 'INT', 'MAX', 'MIN', 'MOD', 'NOT', 'OR', 'PRODUCT', 'ROUND', 'SIGN', 'SUM', 'TRUE'] as const;
  const NUMBER_FORMATS = ['#,##0', '#,##0.00', '$#,##0.00;($#,##0.00)', '0', '0%', '0.00', '0.00%'] as const;
  const NUMERIC = /^\s*[-+(]?[$¥€£]?\s*[\d,.]+\)?\s*%?\s*$/;
  let formula = $state('=SUM(ABOVE)');
  let numberFormat = $state('');
  $effect(() => {
    if (session.dialog !== DIALOGS.formula || !ts) return;
    // Word proposes SUM(ABOVE) when the cells above hold numbers, else SUM(LEFT).
    const { row, cell } = ts.focus;
    const above = row > 0 ? getTableCellText(ts.table, row - 1, Math.min(cell, (ts.table.rows[row - 1]?.cells.length ?? 1) - 1)) : '';
    formula = NUMERIC.test(above) ? '=SUM(ABOVE)' : '=SUM(LEFT)';
    numberFormat = '';
  });
  function pasteFunction(e: Event): void {
    if (!(e.currentTarget instanceof HTMLSelectElement) || !e.currentTarget.value) return;
    formula += `${e.currentTarget.value}()`;
    e.currentTarget.value = '';
  }
</script>

<Dialog id={DIALOGS.deleteCells} title={t('tbl.dlg.deleteCells')} onok={() => { session.apply(commands.deleteCellsCommand, { shift }); }}>
  {#each SHIFTS as s (s.value)}
    <label class="field"><input type="radio" name="tbl-shift" checked={shift === s.value} onchange={() => (shift = s.value)} />{t(s.label)}</label>
  {/each}
</Dialog>

<Dialog id={DIALOGS.splitCells} title={t('tbl.dlg.splitCells')} onok={() => { session.apply(commands.splitCellsCommand, { columns: splitColumns, rows: splitRows }); }}>
  <label class="field">{t('tbl.dlg.columns')}<input type="number" min="1" max="63" bind:value={splitColumns} /></label>
  <label class="field">{t('tbl.dlg.rows')}<input type="number" min="1" max="100" bind:value={splitRows} /></label>
</Dialog>

<Dialog
  id={DIALOGS.cellMargins}
  title={t('tbl.dlg.options')}
  onok={() => {
    session.apply(commands.tablePropertiesCommand, { options: { defaultMargins: margins, cellSpacingTwips: spacingOn ? spacing : 0, autoResize } });
  }}
>
  <fieldset>
    <legend>{t('tbl.dlg.defaultMargins')}</legend>
    <div class="ld-grid">
      <LengthField label={t('tbl.dlg.top')} bind:twips={margins.top} />
      <LengthField label={t('tbl.dlg.left')} bind:twips={margins.left} />
      <LengthField label={t('tbl.dlg.bottom')} bind:twips={margins.bottom} />
      <LengthField label={t('tbl.dlg.right')} bind:twips={margins.right} />
    </div>
  </fieldset>
  <fieldset>
    <legend>{t('tbl.dlg.defaultSpacing')}</legend>
    <div class="ld-row">
      <label class="field"><input type="checkbox" bind:checked={spacingOn} />{t('tbl.dlg.allowSpacing')}</label>
      <LengthField label="" bind:twips={spacing} disabled={!spacingOn} />
    </div>
  </fieldset>
  <label class="field"><input type="checkbox" bind:checked={autoResize} />{t('tbl.dlg.autoResize')}</label>
</Dialog>

<Dialog id={DIALOGS.sort} title={t('tbl.dlg.sort')} onok={sort}>
  {#each sortKeys as key, i (i)}
    <fieldset>
      <legend>{t(i === 0 ? 'tbl.dlg.sortBy' : 'tbl.dlg.thenBy')}</legend>
      <div class="ld-row">
        <select bind:value={key.column} aria-label={t(i === 0 ? 'tbl.dlg.sortBy' : 'tbl.dlg.thenBy')}>
          {#if i > 0}<option value={-1}>{t('tbl.dlg.noneChoice')}</option>{/if}
          {#each columnNames as name, c (c)}<option value={c}>{name}</option>{/each}
        </select>
        <label class="field">{t('tbl.dlg.type')}
          <select bind:value={key.type}>
            <option value="text">{t('tbl.dlg.text')}</option>
            <option value="number">{t('tbl.dlg.number')}</option>
            <option value="date">{t('tbl.dlg.date')}</option>
          </select>
        </label>
        <label class="field"><input type="radio" name="tbl-order-{i}" checked={key.order === 'ascending'} onchange={() => (key.order = 'ascending')} />{t('tbl.dlg.ascending')}</label>
        <label class="field"><input type="radio" name="tbl-order-{i}" checked={key.order === 'descending'} onchange={() => (key.order = 'descending')} />{t('tbl.dlg.descending')}</label>
      </div>
    </fieldset>
  {/each}
  <div class="ld-row">
    {t('tbl.dlg.myList')}
    <label class="field"><input type="radio" name="tbl-sort-header" checked={sortHeader} onchange={() => (sortHeader = true)} />{t('tbl.dlg.headerRow')}</label>
    <label class="field"><input type="radio" name="tbl-sort-header" checked={!sortHeader} onchange={() => (sortHeader = false)} />{t('tbl.dlg.noHeaderRow')}</label>
  </div>
</Dialog>

<Dialog id={DIALOGS.convertToText} title={t('tbl.dlg.convertToText')} onok={() => { session.apply(commands.convertTableToTextCommand, { separator: toTextSeparator }); }}>
  <SeparatorFields legend="tbl.dlg.separateWith" paragraphLabel="tbl.dlg.paragraphMarks" bind:separator={toTextSeparator} />
</Dialog>

<Dialog
  id={DIALOGS.formula}
  title={t('tbl.dlg.formula')}
  onok={() => {
    session.apply(commands.insertFormulaCommand, { formula, ...(numberFormat ? { numberFormat } : {}) });
  }}
  okDisabled={!formula.trim()}
>
  <label class="field">{t('tbl.dlg.formulaField')}<input type="text" class="ld-wide" bind:value={formula} /></label>
  <label class="field">{t('tbl.dlg.numberFormat')}
    <input type="text" list="tbl-number-formats" bind:value={numberFormat} />
    <datalist id="tbl-number-formats">{#each NUMBER_FORMATS as f (f)}<option value={f}></option>{/each}</datalist>
  </label>
  <label class="field">{t('tbl.dlg.pasteFunction')}
    <select onchange={pasteFunction}>
      <option value=""></option>
      {#each FUNCTIONS as f (f)}<option value={f}>{f}</option>{/each}
    </select>
  </label>
</Dialog>

<style>
  .ld-grid {
    display: grid;
    grid-template-columns: auto auto;
    gap: 6px 16px;
  }
  .ld-row {
    display: flex;
    flex-wrap: wrap;
    align-items: center;
    gap: 10px;
  }
  .ld-wide {
    width: 240px;
  }
</style>
