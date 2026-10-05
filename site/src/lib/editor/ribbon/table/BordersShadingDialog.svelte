<script lang="ts">
  /**
   * Borders and Shading for table cells: a Setting preset (None / Box / All /
   * Grid), the line's style, color and width, which edges get it (toggled on
   * the preview like Word's), whether it applies to the selected cells or the
   * whole table, and the cells' fill.
   */
  import { tableColumnCount, type TableBorder, type TableCellRange } from '@office-kit/docx';
  import { commands, tableSelection } from '@office-kit/docx-editor';
  import Dialog from '../../Dialog.svelte';
  import { getSession } from '../../session.svelte';
  import { t, type MessageKey } from '../../i18n/index.svelte';
  import { DIALOGS, LINE_STYLES, LINE_WEIGHTS } from './lines';
  import { tableTool } from './table-tool.svelte';

  const session = getSession();

  type Edge = 'top' | 'bottom' | 'left' | 'right' | 'insideH' | 'insideV' | 'tl2br' | 'tr2bl';
  const EDGES: readonly Edge[] = ['top', 'bottom', 'left', 'right', 'insideH', 'insideV', 'tl2br', 'tr2bl'];
  const EDGE_LABEL: Record<Edge, MessageKey> = {
    top: 'tbl.border.top',
    bottom: 'tbl.border.bottom',
    left: 'tbl.border.left',
    right: 'tbl.border.right',
    insideH: 'tbl.border.insideH',
    insideV: 'tbl.border.insideV',
    tl2br: 'tbl.border.tl2br',
    tr2bl: 'tbl.border.tr2bl',
  };
  type Setting = 'none' | 'box' | 'all' | 'grid';
  const SETTINGS: readonly Setting[] = ['none', 'box', 'all', 'grid'];
  const SETTING_LABEL: Record<Setting, MessageKey> = { none: 'tbl.bs.none', box: 'tbl.bs.box', all: 'tbl.bs.all', grid: 'tbl.bs.grid' };
  const OUTER: ReadonlySet<Edge> = new Set(['top', 'bottom', 'left', 'right']);
  // Grid: Word frames the table with a heavier (1½ pt) line than the inside.
  const GRID_OUTER_SIZE = 12;

  let setting = $state<Setting>('all');
  let style = $state('single');
  let size = $state(4);
  let color = $state('000000');
  let on = $state<Record<Edge, boolean>>(edgesFor('all'));
  let applyTo = $state<'cell' | 'table'>('cell');
  let fill = $state<string | undefined>(undefined);

  function edgesFor(s: Setting): Record<Edge, boolean> {
    const inner = s === 'all' || s === 'grid';
    return { top: s !== 'none', bottom: s !== 'none', left: s !== 'none', right: s !== 'none', insideH: inner, insideV: inner, tl2br: false, tr2bl: false };
  }

  const opened = $derived(session.dialog === DIALOGS.bordersShading);
  $effect(() => {
    if (!opened) return;
    const pen = tableTool.pen;
    style = pen.style;
    size = pen.size;
    color = pen.color === 'auto' ? '000000' : pen.color;
    setting = 'all';
    on = edgesFor('all');
    applyTo = session.model && tableSelection(session.model)?.multiCell ? 'cell' : 'table';
    fill = undefined;
  });

  function choose(s: Setting): void {
    setting = s;
    on = edgesFor(s);
  }

  function ok(): void {
    const model = session.model;
    const ts = model ? tableSelection(model) : undefined;
    if (!ts) return;
    const range: TableCellRange =
      applyTo === 'table' ? { firstRow: 0, lastRow: ts.table.rows.length - 1, firstColumn: 0, lastColumn: tableColumnCount(ts.table) - 1 } : ts.range;
    const line: TableBorder = { style, size, color };
    const outer: TableBorder = setting === 'grid' ? { ...line, size: Math.max(size, GRID_OUTER_SIZE) } : line;
    const borders: Partial<Record<Edge, TableBorder | null>> = {};
    for (const edge of EDGES) borders[edge] = on[edge] ? (OUTER.has(edge) ? outer : line) : null;
    session.apply(commands.bordersAndShadingCommand, { range, borders, ...(fill === undefined ? {} : { fill }) });
  }
</script>

<Dialog id={DIALOGS.bordersShading} title={t('tbl.bs.title')} onok={ok}>
  <div class="bs">
    <fieldset class="bs-settings">
      <legend>{t('tbl.bs.setting')}</legend>
      {#each SETTINGS as s (s)}
        <label class="field"><input type="radio" name="bs-setting" checked={setting === s} onchange={() => choose(s)} />{t(SETTING_LABEL[s])}</label>
      {/each}
    </fieldset>
    <fieldset>
      <label class="field">{t('tbl.bs.style')}
        <select bind:value={style}>{#each LINE_STYLES as l (l.style)}<option value={l.style}>{t(l.label)}</option>{/each}</select>
      </label>
      <label class="field">{t('tbl.bs.color')}<input type="color" value={`#${color}`} onchange={(e) => (color = e.currentTarget.value.slice(1).toUpperCase())} /></label>
      <label class="field">{t('tbl.bs.width')}
        <select bind:value={size}>{#each LINE_WEIGHTS as w (w.size)}<option value={w.size}>{w.label}</option>{/each}</select>
      </label>
    </fieldset>
    <fieldset class="bs-edges">
      {#each EDGES as edge (edge)}
        <label class="field"><input type="checkbox" bind:checked={on[edge]} />{t(EDGE_LABEL[edge])}</label>
      {/each}
    </fieldset>
  </div>
  <div class="bs-row">
    <label class="field">{t('tbl.bs.applyTo')}
      <select bind:value={applyTo}>
        <option value="cell">{t('tbl.bs.cell')}</option>
        <option value="table">{t('tbl.bs.table')}</option>
      </select>
    </label>
    <label class="field">{t('tbl.bs.fill')}<input type="color" value={fill && fill !== 'auto' ? `#${fill}` : '#ffffff'} onchange={(e) => (fill = e.currentTarget.value.slice(1).toUpperCase())} /></label>
    <button type="button" class="push" onclick={() => (fill = 'auto')}>{t('tbl.color.none')}</button>
  </div>
</Dialog>

<style>
  .bs {
    display: flex;
    gap: 10px;
  }
  .bs fieldset {
    display: flex;
    flex-direction: column;
    gap: 6px;
  }
  .bs-row {
    display: flex;
    align-items: center;
    gap: 12px;
  }
</style>
