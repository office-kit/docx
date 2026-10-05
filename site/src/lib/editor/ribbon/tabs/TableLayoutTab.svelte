<script lang="ts">
  /**
   * Table Layout (contextual, while the caret is in a table): Select, View
   * Gridlines, Properties, Draw Table / Eraser, Delete, Insert Above / Below /
   * Left / Right, Merge / Split, AutoFit, the Height and Width spinners,
   * Distribute, the Alignment grid, Text Direction, Cell Margins, and the Data
   * group (Sort, Repeat Header Rows, Convert to Text, Formula).
   */
  import type { TableCellTextDirection, TableCellVerticalAlign } from '@office-kit/docx';
  import { commands, resolveTable, selectInTable, tableSelection, type TableSelectTarget } from '@office-kit/docx-editor';
  import Button from '../Button.svelte';
  import Group from '../Group.svelte';
  import SplitButton from '../SplitButton.svelte';
  import RibbonIcon, { type IconName } from '../../RibbonIcon.svelte';
  import { getSession } from '../../session.svelte';
  import { t, type MessageKey } from '../../i18n/index.svelte';
  import LayoutDialogs from '../table/LayoutDialogs.svelte';
  import TablePropertiesDialog from '../table/TablePropertiesDialog.svelte';
  import BordersShadingDialog from '../table/BordersShadingDialog.svelte';
  import { measuredRowHeight } from '../../table-canvas';
  import { DIALOGS } from '../table/lines';
  import { lengthUnit, tableTool, toggleTableTool, twipsToUnit, unitStep, unitToTwips } from '../table/table-tool.svelte';

  const session = getSession();
  const ts = $derived(session.tick >= 0 && session.model ? tableSelection(session.model) : undefined);
  const unit = $derived(lengthUnit());

  const SELECTS: readonly { what: TableSelectTarget; label: MessageKey }[] = [
    { what: 'cell', label: 'tbl.select.cell' },
    { what: 'column', label: 'tbl.select.column' },
    { what: 'row', label: 'tbl.select.row' },
    { what: 'table', label: 'tbl.select.table' },
  ];
  function select(what: TableSelectTarget): void {
    session.openMenu = null;
    const model = session.model;
    const sel = model && selectInTable(model, what);
    if (!model || !sel) return;
    model.setSelection(sel);
    session.changed();
  }

  // The spinners show the focused row's height and column's width as laid out.
  const size = $derived.by(() => {
    const model = session.model;
    if (!model || !ts) return { height: 0, width: 0 };
    const format = resolveTable(model.doc, ts.table);
    const p = ts.placements[ts.focus.row]?.[ts.focus.cell];
    const width = p ? format.columns.slice(p.gridStart, p.gridStart + p.gridSpan).reduce((a, b) => a + b, 0) : 0;
    // A row without a set height shows its laid-out height, as Word does.
    const height = format.rows[ts.focus.row]?.height?.value ?? measuredRowHeight(ts.block, ts.focus.row, ts.focus.row) ?? 0;
    return { height, width };
  });
  function spin(e: Event, apply: (twips: number) => void): void {
    if (!(e.currentTarget instanceof HTMLInputElement)) return;
    const value = Number(e.currentTarget.value);
    if (Number.isFinite(value) && value >= 0) apply(unitToTwips(value, unit));
  }

  const ALIGN: readonly { h: 'left' | 'center' | 'right'; v: TableCellVerticalAlign; icon: IconName; label: MessageKey }[] = [
    { h: 'left', v: 'top', icon: 'tblAlignTL', label: 'tbl.align.topLeft' },
    { h: 'center', v: 'top', icon: 'tblAlignTC', label: 'tbl.align.topCenter' },
    { h: 'right', v: 'top', icon: 'tblAlignTR', label: 'tbl.align.topRight' },
    { h: 'left', v: 'center', icon: 'tblAlignCL', label: 'tbl.align.centerLeft' },
    { h: 'center', v: 'center', icon: 'tblAlignCC', label: 'tbl.align.center' },
    { h: 'right', v: 'center', icon: 'tblAlignCR', label: 'tbl.align.centerRight' },
    { h: 'left', v: 'bottom', icon: 'tblAlignBL', label: 'tbl.align.bottomLeft' },
    { h: 'center', v: 'bottom', icon: 'tblAlignBC', label: 'tbl.align.bottomCenter' },
    { h: 'right', v: 'bottom', icon: 'tblAlignBR', label: 'tbl.align.bottomRight' },
  ];

  // Word's Text Direction button cycles horizontal → top-to-bottom → bottom-to-top.
  const DIRECTION_CYCLE: readonly TableCellTextDirection[] = ['lrTb', 'tbRl', 'btLr'];
  // Where each ST_TextDirection value sits in the cycle (the East Asian variants with their base).
  const CYCLE_POSITION: Readonly<Record<string, number>> = { lrTb: 0, lrTbV: 0, tbRl: 1, tbRlV: 1, btLr: 2, tbLrV: 2 };
  function cycleDirection(): void {
    const model = session.model;
    if (!model || !ts) return;
    const current = resolveTable(model.doc, ts.table).cells[ts.focus.row]?.[ts.focus.cell]?.textDirection ?? 'lrTb';
    const direction = DIRECTION_CYCLE[((CYCLE_POSITION[current] ?? 0) + 1) % DIRECTION_CYCLE.length] ?? 'lrTb';
    session.apply(commands.cellTextDirectionCommand, { direction });
  }

  const enabled = $derived(!!ts);

  /** Word gives the selected rows their average laid-out height. */
  function distributeRows(): void {
    if (!ts) return;
    const heightTwips = measuredRowHeight(ts.block, ts.range.firstRow, ts.range.lastRow);
    if (heightTwips !== undefined) session.apply(commands.distributeRowsCommand, { heightTwips });
  }
</script>

<Group label={t('tbl.group.select')}>
  <SplitButton id="tbl-select" size="large" tip={t('tbl.select')} icon="tblSelect" disabled={!enabled}>
    {#snippet menu()}
      {#each SELECTS as s (s.what)}<button class="mi" role="menuitem" onclick={() => select(s.what)}>{t(s.label)}</button>{/each}
    {/snippet}
  </SplitButton>
  <Button size="large" icon="tblGridlines" tip={t('tbl.gridlines')} on={tableTool.gridlines} onclick={() => (tableTool.gridlines = !tableTool.gridlines)} />
  <Button size="large" icon="tblProperties" tip={t('tbl.properties')} disabled={!enabled} onclick={() => session.openDialog(DIALOGS.properties)} />
</Group>

<Group label={t('tbl.group.draw')}>
  <Button size="large" icon="tblDraw" tip={t('tbl.drawTable')} on={tableTool.mode === 'draw'} onclick={() => toggleTableTool('draw')} />
  <Button size="large" icon="tblEraser" tip={t('tbl.eraser')} on={tableTool.mode === 'eraser'} onclick={() => toggleTableTool('eraser')} />
</Group>

<Group label={t('tbl.group.rowsCols')}>
  <SplitButton id="tbl-delete" size="large" tip={t('tbl.delete')} icon="tblDelete" disabled={!enabled}>
    {#snippet menu()}
      <button class="mi" role="menuitem" onclick={() => session.openDialog(DIALOGS.deleteCells)}>{t('tbl.delete.cells')}</button>
      <button class="mi" role="menuitem" onclick={() => session.apply(commands.deleteColumnsCommand, undefined)}>{t('tbl.delete.columns')}</button>
      <button class="mi" role="menuitem" onclick={() => session.apply(commands.deleteRowCommand, {})}>{t('tbl.delete.rows')}</button>
      <button class="mi" role="menuitem" onclick={() => session.apply(commands.deleteTableCommand, undefined)}>{t('tbl.delete.table')}</button>
    {/snippet}
  </SplitButton>
  <Button size="large" icon="tblInsertAbove" tip={t('tbl.insertAbove')} disabled={!enabled} onclick={() => session.apply(commands.insertRowsCommand, { where: 'above' })} />
  <Button size="large" icon="tblInsertBelow" tip={t('tbl.insertBelow')} disabled={!enabled} onclick={() => session.apply(commands.insertRowsCommand, { where: 'below' })} />
  <Button size="large" icon="tblInsertLeft" tip={t('tbl.insertLeft')} disabled={!enabled} onclick={() => session.apply(commands.insertColumnsCommand, { where: 'left' })} />
  <Button size="large" icon="tblInsertRight" tip={t('tbl.insertRight')} disabled={!enabled} onclick={() => session.apply(commands.insertColumnsCommand, { where: 'right' })} />
</Group>

<Group label={t('tbl.group.merge')}>
  <Button size="large" icon="tblMerge" tip={t('tbl.mergeCells')} disabled={!ts?.multiCell} onclick={() => session.apply(commands.mergeCellsCommand, {})} />
  <Button size="large" icon="tblSplit" tip={t('tbl.splitCells')} disabled={!enabled} onclick={() => session.openDialog(DIALOGS.splitCells)} />
  <Button size="large" icon="tblSplitTable" tip={t('tbl.splitTable')} disabled={!enabled} onclick={() => session.apply(commands.splitTableCommand, undefined)} />
</Group>

<Group label={t('tbl.group.cellSize')}>
  <SplitButton id="tbl-autofit" size="large" tip={t('tbl.autoFit')} icon="tblAutoFit" disabled={!enabled}>
    {#snippet menu()}
      <button class="mi" role="menuitem" onclick={() => session.apply(commands.autoFitCommand, { mode: 'contents' })}>{t('tbl.autoFit.contents')}</button>
      <button class="mi" role="menuitem" onclick={() => session.apply(commands.autoFitCommand, { mode: 'window' })}>{t('tbl.autoFit.window')}</button>
      <button class="mi" role="menuitem" onclick={() => session.apply(commands.autoFitCommand, { mode: 'fixed' })}>{t('tbl.autoFit.fixed')}</button>
    {/snippet}
  </SplitButton>
  <div class="rows">
    <label class="field"
      >{t('tbl.height')}<input
        type="number"
        min="0"
        step={unitStep(unit)}
        disabled={!enabled}
        value={twipsToUnit(size.height, unit)}
        onchange={(e) => spin(e, (twips) => session.apply(commands.rowHeightSelectionCommand, { twips }))}
      />{t(`tbl.unit.${unit}`)}</label
    >
    <label class="field"
      >{t('tbl.width')}<input
        type="number"
        min="0"
        step={unitStep(unit)}
        disabled={!enabled}
        value={twipsToUnit(size.width, unit)}
        onchange={(e) => spin(e, (twips) => session.apply(commands.columnWidthSelectionCommand, { twips }))}
      />{t(`tbl.unit.${unit}`)}</label
    >
  </div>
  <div class="col">
    <Button size="mid" icon="tblDistributeRows" tip={t('tbl.distributeRows')} disabled={!enabled} onclick={distributeRows} />
    <Button size="mid" icon="tblDistributeColumns" tip={t('tbl.distributeColumns')} disabled={!enabled} onclick={() => session.apply(commands.distributeColumnsCommand, undefined)} />
  </div>
</Group>

<Group label={t('tbl.group.alignment')}>
  <div class="tbl-align">
    {#each ALIGN as a (a.icon)}
      <Button icon={a.icon} tip={t(a.label)} disabled={!enabled} onclick={() => session.apply(commands.cellAlignmentCommand, { horizontal: a.h, vertical: a.v })} />
    {/each}
  </div>
  <Button size="large" icon="tblTextDirection" tip={t('tbl.textDirection')} disabled={!enabled} onclick={cycleDirection} />
  <Button size="large" icon="tblCellMargins" tip={t('tbl.cellMargins')} disabled={!enabled} onclick={() => session.openDialog(DIALOGS.cellMargins)} />
</Group>

<Group label={t('tbl.group.data')}>
  <Button size="large" icon="tblSort" tip={t('tbl.sort')} disabled={!enabled} onclick={() => session.openDialog(DIALOGS.sort)} />
  <Button size="large" icon="tblRepeatHeader" tip={t('tbl.repeatHeader')} disabled={!enabled} on={session.active(commands.repeatHeaderRowsCommand)} onclick={() => session.apply(commands.repeatHeaderRowsCommand, undefined)} />
  <Button size="large" icon="tblConvertToText" tip={t('tbl.convertToText')} disabled={!enabled} onclick={() => session.openDialog(DIALOGS.convertToText)} />
  <Button size="large" icon="tblFormula" tip={t('tbl.formula')} disabled={!enabled} onclick={() => session.openDialog(DIALOGS.formula)} />
</Group>

<TablePropertiesDialog />
<!-- Table Properties ▸ Borders and Shading… opens it from this tab. -->
<BordersShadingDialog />
<LayoutDialogs />

<style>
  .tbl-align {
    display: grid;
    grid-template-columns: repeat(3, 28px);
    gap: 0;
  }
</style>
