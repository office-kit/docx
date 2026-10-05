<script lang="ts">
  /**
   * Insert ▸ Table: the size grid (hover to size, click to insert), Insert
   * Table…, Draw Table, Convert Text to Table… and Quick Tables, with their
   * dialogs.
   */
  import type { TableTextSeparator } from '@office-kit/docx';
  import { commands } from '@office-kit/docx-editor';
  import Dialog from '../Dialog.svelte';
  import RibbonIcon from '../RibbonIcon.svelte';
  import { getSession } from '../session.svelte';
  import { t, type MessageKey } from '../i18n/index.svelte';
  import SplitButton from './SplitButton.svelte';
  import LengthField from './table/LengthField.svelte';
  import SeparatorFields from './table/SeparatorFields.svelte';
  import { DIALOGS } from './table/lines';
  import { tableTool } from './table/table-tool.svelte';

  const session = getSession();
  // Word for Mac's grid: 10 columns by 8 rows.
  const GRID_COLS = 10;
  const GRID_ROWS = 8;
  // Insert Table's defaults and limits (Word caps columns at 63).
  const DEFAULT_COLS = 5;
  const DEFAULT_ROWS = 2;
  const MAX_COLS = 63;
  const MAX_ROWS = 1000;
  const DEFAULT_COLUMN_WIDTH = 1440;

  let hover = $state({ rows: 0, cols: 0 });
  const sizeLabel = $derived(hover.rows ? `${hover.cols} × ${hover.rows} ${t('tbl.btn.size')}` : t('ins.table'));

  function insert(rows: number, cols: number): void {
    session.apply(commands.insertTableCommand, { rows, cols });
    hover = { rows: 0, cols: 0 };
  }

  // --- Insert Table… ----------------------------------------------------------------------
  let cols = $state(DEFAULT_COLS);
  let rows = $state(DEFAULT_ROWS);
  let autoFit = $state<'fixed' | 'contents' | 'window'>('fixed');
  let fixedAuto = $state(true);
  let columnWidth = $state(DEFAULT_COLUMN_WIDTH);
  const sizeValid = $derived(Number.isInteger(cols) && cols >= 1 && cols <= MAX_COLS && Number.isInteger(rows) && rows >= 1 && rows <= MAX_ROWS);

  function insertFromDialog(): boolean {
    if (!sizeValid) return false;
    session.apply(commands.insertTableCommand, {
      rows,
      cols,
      autoFit,
      ...(autoFit === 'fixed' && !fixedAuto ? { columnWidthTwips: columnWidth } : {}),
    });
    return true;
  }

  // --- Convert Text to Table… -------------------------------------------------------------
  let separator = $state<TableTextSeparator>('tab');
  const canConvert = $derived(session.enabled(commands.convertTextToTableCommand));

  // --- Quick Tables -----------------------------------------------------------------------
  const QUICK: readonly { label: MessageKey; styleId: string; cells: string[][] }[] = [
    {
      label: 'tbl.quick.matrix',
      styleId: 'GridTable4-Accent1',
      cells: [
        ['', 'Point A', 'Point B', 'Point C'],
        ['Point A', '—', '', ''],
        ['Point B', '87', '—', ''],
        ['Point C', '64', '56', '—'],
      ],
    },
    {
      label: 'tbl.quick.tabular',
      styleId: 'PlainTable4',
      cells: [
        ['Item', 'Needed'],
        ['Books', '1'],
        ['Magazines', '3'],
        ['Notebooks', '1'],
      ],
    },
    {
      label: 'tbl.quick.double',
      styleId: 'ListTable3-Accent1',
      cells: [
        ['Class', 'Grad', 'Undergrad', 'Class', 'Grad', 'Undergrad'],
        ['Cedar', '9', '28', 'Elm', '7', '25'],
        ['Maple', '4', '17', 'Oak', '6', '31'],
      ],
    },
    {
      label: 'tbl.quick.subheads',
      styleId: 'GridTable1Light-Accent1',
      cells: [
        ['College', 'New students', 'Graduating students', 'Change'],
        ['Undergraduate', '', '', ''],
        ['Cedar University', '110', '103', '+7'],
        ['Elm College', '223', '214', '+9'],
        ['Graduate', '', '', ''],
        ['Cedar University', '24', '20', '+4'],
      ],
    },
  ];
  let quickOpen = $state(false);
</script>

<SplitButton id="ins-table" size="large" tip={t('ins.table')} icon="table">
  {#snippet menu()}
    <div class="tb-head">{sizeLabel}</div>
    <div class="tb-grid" role="grid" tabindex="-1" onmouseleave={() => (hover = { rows: 0, cols: 0 })}>
      {#each Array.from({ length: GRID_ROWS }, (_, r) => r) as r (r)}
        {#each Array.from({ length: GRID_COLS }, (_, c) => c) as c (c)}
          <button
            class="tb-cell"
            class:on={r < hover.rows && c < hover.cols}
            aria-label={`${c + 1} × ${r + 1}`}
            onmouseenter={() => (hover = { rows: r + 1, cols: c + 1 })}
            onfocus={() => (hover = { rows: r + 1, cols: c + 1 })}
            onclick={() => insert(r + 1, c + 1)}
          ></button>
        {/each}
      {/each}
    </div>
    <hr />
    <button class="mi" role="menuitem" onclick={() => session.openDialog(DIALOGS.insertTable)}><RibbonIcon name="table" size={16} />{t('tbl.btn.insertTable')}</button>
    <button class="mi" role="menuitem" onclick={() => { tableTool.mode = 'draw'; session.openMenu = null; }}><RibbonIcon name="tblDraw" size={16} />{t('tbl.drawTable')}</button>
    <button class="mi" role="menuitem" disabled={!canConvert} onclick={() => session.openDialog(DIALOGS.convertTextToTable)}><RibbonIcon name="tblConvertTextToTable" size={16} />{t('tbl.btn.convertText')}</button>
    <button class="mi" role="menuitem" aria-expanded={quickOpen} onclick={() => (quickOpen = !quickOpen)}><RibbonIcon name="tblQuickTables" size={16} />{t('tbl.btn.quickTables')} ▸</button>
    {#if quickOpen}
      {#each QUICK as q (q.label)}
        <button class="mi tb-sub" role="menuitem" onclick={() => session.apply(commands.insertTableCommand, { rows: q.cells.length, cols: q.cells[0]?.length ?? 1, styleId: q.styleId, cells: q.cells })}>{t(q.label)}</button>
      {/each}
    {/if}
  {/snippet}
</SplitButton>

<Dialog id={DIALOGS.insertTable} title={t('tbl.dlg.insertTable')} onok={insertFromDialog} okDisabled={!sizeValid}>
  <fieldset>
    <legend>{t('tbl.dlg.tableSize')}</legend>
    <div class="tb-fields">
      <label class="field">{t('tbl.dlg.columns')}<input type="number" min="1" max={MAX_COLS} bind:value={cols} /></label>
      <label class="field">{t('tbl.dlg.rows')}<input type="number" min="1" max={MAX_ROWS} bind:value={rows} /></label>
    </div>
  </fieldset>
  <fieldset>
    <legend>{t('tbl.dlg.autoFitBehavior')}</legend>
    <div class="tb-fields">
      <label class="field"><input type="radio" name="tb-autofit" checked={autoFit === 'fixed'} onchange={() => (autoFit = 'fixed')} />{t('tbl.dlg.initialWidth')}</label>
      <span class="tb-inline">
        <label class="field"><input type="checkbox" bind:checked={fixedAuto} disabled={autoFit !== 'fixed'} />{t('tbl.dlg.auto')}</label>
        <LengthField label="" bind:twips={columnWidth} disabled={autoFit !== 'fixed' || fixedAuto} />
      </span>
      <label class="field"><input type="radio" name="tb-autofit" checked={autoFit === 'contents'} onchange={() => (autoFit = 'contents')} />{t('tbl.dlg.fitContents')}</label>
      <label class="field"><input type="radio" name="tb-autofit" checked={autoFit === 'window'} onchange={() => (autoFit = 'window')} />{t('tbl.dlg.fitWindow')}</label>
    </div>
  </fieldset>
</Dialog>

<Dialog id={DIALOGS.convertTextToTable} title={t('tbl.dlg.convertToTable')} onok={() => { session.apply(commands.convertTextToTableCommand, { separator }); }}>
  <SeparatorFields legend="tbl.dlg.separateAt" paragraphLabel="tbl.dlg.paragraphs" bind:separator />
</Dialog>

<style>
  .tb-head {
    padding: 2px 4px 6px;
    font-size: 12px;
    text-align: center;
  }
  .tb-grid {
    display: grid;
    grid-template-columns: repeat(10, 16px);
    gap: 3px;
    padding: 0 4px 4px;
  }
  .tb-cell {
    width: 16px;
    height: 16px;
    padding: 0;
    border: 1px solid #bdbdbd;
    background: #fff;
    cursor: pointer;
  }
  .tb-cell.on {
    border-color: #e36c0a;
    background: #fde9d9;
  }
  .tb-sub {
    padding-left: 32px;
  }
  .tb-fields {
    display: flex;
    flex-direction: column;
    gap: 6px;
  }
  .tb-inline {
    display: flex;
    align-items: center;
    gap: 8px;
    padding-left: 22px;
  }
</style>
