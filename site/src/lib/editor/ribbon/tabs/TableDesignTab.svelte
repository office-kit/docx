<script lang="ts">
  /**
   * Table Design (contextual, while the caret is in a table): Table Style
   * Options (`w:tblLook`), the Table Styles gallery, Shading, and the Borders
   * group — Border Styles, the pen (Line Style / Line Weight / Pen Color), the
   * Borders menu and Border Painter.
   */
  import { BUILT_IN_TABLE_STYLES, listStyles, type TableBorder, type TableBorderEdges, type TableLook } from '@office-kit/docx';
  import { commands, tableStylePreviews, type TablePreviewCell } from '@office-kit/docx-editor';
  import Button from '../Button.svelte';
  import Group from '../Group.svelte';
  import SplitButton from '../SplitButton.svelte';
  import RibbonIcon, { type IconName } from '../../RibbonIcon.svelte';
  import { getSession } from '../../session.svelte';
  import { t, type MessageKey } from '../../i18n/index.svelte';
  import BordersShadingDialog from '../table/BordersShadingDialog.svelte';
  import ColorMenu from '../table/ColorMenu.svelte';
  import TableStyleDialog from '../table/TableStyleDialog.svelte';
  import TableStyleTile from '../table/TableStyleTile.svelte';
  import { DIALOGS, LINE_STYLES, LINE_WEIGHTS } from '../table/lines';
  import { penCss, tableTool, toggleTableTool } from '../table/table-tool.svelte';

  const session = getSession();

  const OPTIONS: readonly { key: keyof TableLook; label: MessageKey }[] = [
    { key: 'headerRow', label: 'tbl.opt.headerRow' },
    { key: 'firstColumn', label: 'tbl.opt.firstColumn' },
    { key: 'totalRow', label: 'tbl.opt.totalRow' },
    { key: 'lastColumn', label: 'tbl.opt.lastColumn' },
    { key: 'bandedRows', label: 'tbl.opt.bandedRows' },
    { key: 'bandedColumns', label: 'tbl.opt.bandedColumns' },
  ];
  // Word's defaults for a new table (tblLook 04A0).
  const DEFAULT_LOOK: TableLook = { headerRow: true, totalRow: false, firstColumn: true, lastColumn: false, bandedRows: true, bandedColumns: false };
  // Tiles shown in the ribbon before the gallery's More menu.
  const INLINE_TILES = 7;

  const look = $derived((session.tick >= 0 && session.model && commands.currentTableLook(session.model)) || DEFAULT_LOOK);
  const lookKey = $derived(OPTIONS.map((o) => (look[o.key] ? 1 : 0)).join(''));
  const current = $derived(session.tick >= 0 && session.model ? commands.currentTableStyle(session.model) : undefined);

  const customStyles = $derived.by(() => {
    if (!session.model || session.version < 0) return [];
    const builtIn = new Set(BUILT_IN_TABLE_STYLES.map((s) => s.styleId));
    return listStyles(session.model.doc)
      .filter((s) => s.type === 'table' && s.styleId !== 'TableNormal' && !builtIn.has(s.styleId))
      .map((s) => ({ styleId: s.styleId, name: s.styleId }));
  });
  const galleryOpen = $derived(session.openMenu === 'tbl-styles');
  const sections = $derived([
    { label: 'tbl.styles.custom' as const, styles: customStyles },
    { label: 'tbl.styles.plain' as const, styles: BUILT_IN_TABLE_STYLES.filter((s) => s.category === 'plain') },
    { label: 'tbl.styles.grid' as const, styles: BUILT_IN_TABLE_STYLES.filter((s) => s.category === 'grid') },
    { label: 'tbl.styles.list' as const, styles: BUILT_IN_TABLE_STYLES.filter((s) => s.category === 'list') },
  ]);
  const inline = $derived(sections.flatMap((s) => s.styles).slice(0, INLINE_TILES));
  // Thumbnails depend on the document (its theme and its own table styles) and
  // the Table Style Options, not on the caret, so they are rebuilt only then;
  // the full gallery only while it is open.
  const previews = $derived.by(() => {
    const model = session.model;
    if (!model || session.version < 0) return new Map<string, TablePreviewCell[][]>();
    const ids = (galleryOpen ? sections.flatMap((s) => s.styles) : inline).map((s) => s.styleId);
    return tableStylePreviews(model.doc, ids, lookFromKey(lookKey));
  });

  function lookFromKey(key: string): TableLook {
    const entries = OPTIONS.map((o, i) => [o.key, key[i] === '1'] as const);
    return { ...DEFAULT_LOOK, ...Object.fromEntries(entries) };
  }

  function setOption(key: keyof TableLook, value: boolean): void {
    const change: Partial<Record<keyof TableLook, boolean>> = {};
    change[key] = value;
    session.apply(commands.setTableLookCommand, change);
  }

  const setStyle = (styleId: string | undefined): void => {
    session.apply(commands.setTableStyleCommand, { styleId });
  };

  // --- Borders ----------------------------------------------------------------------

  const BORDER_ITEMS: readonly { edges: TableBorderEdges; icon: IconName; label: MessageKey }[] = [
    { edges: 'bottom', icon: 'tblBorderBottom', label: 'tbl.border.bottom' },
    { edges: 'top', icon: 'tblBorderTop', label: 'tbl.border.top' },
    { edges: 'left', icon: 'tblBorderLeft', label: 'tbl.border.left' },
    { edges: 'right', icon: 'tblBorderRight', label: 'tbl.border.right' },
    { edges: 'none', icon: 'tblBorderNone', label: 'tbl.border.none' },
    { edges: 'all', icon: 'tblBorderAll', label: 'tbl.border.all' },
    { edges: 'outside', icon: 'tblBorderOutside', label: 'tbl.border.outside' },
    { edges: 'inside', icon: 'tblBorderInside', label: 'tbl.border.inside' },
    { edges: 'insideH', icon: 'tblBorderInsideH', label: 'tbl.border.insideH' },
    { edges: 'insideV', icon: 'tblBorderInsideV', label: 'tbl.border.insideV' },
    { edges: 'tl2br', icon: 'tblBorderDiagDown', label: 'tbl.border.tl2br' },
    { edges: 'tr2bl', icon: 'tblBorderDiagUp', label: 'tbl.border.tr2bl' },
  ];
  const lastBorder = $derived(BORDER_ITEMS.find((b) => b.edges === tableTool.borders) ?? BORDER_ITEMS[0]);

  function applyBorders(edges: TableBorderEdges): void {
    tableTool.borders = edges;
    session.apply(commands.rangeBordersCommand, edges === 'none' ? { edges } : { edges, border: tableTool.pen });
  }

  // Border Styles: Word's theme border presets. Picking one loads the pen and
  // turns on Border Painter, as in Word.
  const BORDER_PRESETS: readonly TableBorder[] = [
    { style: 'single', size: 4, color: 'auto' },
    { style: 'single', size: 4, color: '4472C4' },
    { style: 'single', size: 4, color: 'ED7D31' },
    { style: 'single', size: 4, color: '70AD47' },
    { style: 'single', size: 12, color: 'auto' },
    { style: 'single', size: 12, color: '4472C4' },
    { style: 'double', size: 4, color: 'auto' },
    { style: 'double', size: 4, color: '4472C4' },
    { style: 'single', size: 24, color: 'auto' },
    { style: 'dotted', size: 4, color: 'auto' },
    { style: 'dashed', size: 4, color: 'auto' },
    { style: 'thickThinSmallGap', size: 18, color: 'auto' },
  ];

  /** Changing the pen arms Border Painter, as Word does. */
  function setPen(change: Partial<TableBorder>): void {
    tableTool.pen = { ...tableTool.pen, ...change };
    tableTool.mode = 'painter';
    session.openMenu = null;
  }

  const shadingSwatch = $derived(tableTool.shading === 'auto' ? 'transparent' : `#${tableTool.shading}`);
  const penSwatch = $derived(tableTool.pen.color === 'auto' ? '#000' : `#${tableTool.pen.color}`);
</script>

<Group label={t('tbl.group.options')}>
  <div class="tbl-options">
    {#each OPTIONS as o (o.key)}
      <label class="check-item"
        ><input type="checkbox" checked={look[o.key]} onchange={(e) => setOption(o.key, e.currentTarget.checked)} />{t(o.label)}</label
      >
    {/each}
  </div>
</Group>

<Group label={t('tbl.group.styles')} class="gallery-group">
  <div class="gallery tbl-gallery">
    {#each inline as s (s.styleId)}
      <TableStyleTile name={s.name} cells={previews.get(s.styleId)} selected={current === s.styleId} onclick={() => setStyle(s.styleId)} />
    {/each}
  </div>
  <SplitButton id="tbl-styles" tip={t('tbl.styles.more')}>
    {#snippet menu()}
      <div class="tbl-gallery-menu">
        {#each sections as section (section.label)}
          {#if section.styles.length > 0}
            <div class="menu-head">{t(section.label)}</div>
            <div class="tbl-gallery-grid">
              {#each section.styles as s (s.styleId)}
                <TableStyleTile name={s.name} cells={previews.get(s.styleId)} selected={current === s.styleId} onclick={() => setStyle(s.styleId)} />
              {/each}
            </div>
          {/if}
        {/each}
      </div>
      <hr />
      <button class="mi" role="menuitem" disabled={!current} onclick={() => session.openDialog(DIALOGS.modifyStyle)}>{t('tbl.styles.modify')}</button>
      <button class="mi" role="menuitem" onclick={() => setStyle(undefined)}>{t('tbl.styles.clear')}</button>
      <button class="mi" role="menuitem" onclick={() => session.openDialog(DIALOGS.newStyle)}><RibbonIcon name="tblNewStyle" size={16} />{t('tbl.styles.new')}</button>
    {/snippet}
  </SplitButton>
</Group>

<Group label={t('tbl.shading')}>
  <SplitButton id="tbl-shading" size="large" tip={t('tbl.shading')} onclick={() => session.apply(commands.shadeCellsCommand, { fill: tableTool.shading })}>
    {#snippet face()}<span class="tbl-face"><RibbonIcon name="tblShading" size={32} /><i style="background:{shadingSwatch}"></i></span>{/snippet}
    {#snippet menu()}
      <ColorMenu autoLabel="tbl.color.none" onpick={(fill) => { tableTool.shading = fill; session.apply(commands.shadeCellsCommand, { fill }); }} />
    {/snippet}
  </SplitButton>
</Group>

<Group label={t('tbl.group.borders')}>
  <SplitButton id="tbl-border-styles" size="large" tip={t('tbl.borderStyles')} icon="tblBorderStyles">
    {#snippet menu()}
      <div class="menu-head">{t('tbl.borderStyles')}</div>
      <div class="tbl-presets">
        {#each BORDER_PRESETS as preset, i (i)}
          <button class="tbl-preset" role="menuitem" title={`${preset.style} ${preset.size / 8} pt`} onclick={() => setPen(preset)}><i style="border-top:{penCss(preset)}"></i></button>
        {/each}
      </div>
    {/snippet}
  </SplitButton>
  <div class="rows">
    <div class="row">
      <select class="combo tbl-pen" aria-label={t('tbl.lineStyle')} title={t('tbl.lineStyle')} value={tableTool.pen.style} onchange={(e) => setPen({ style: e.currentTarget.value })}>
        {#each LINE_STYLES as l (l.style)}<option value={l.style}>{t(l.label)}</option>{/each}
      </select>
    </div>
    <div class="row">
      <select class="combo tbl-pen" aria-label={t('tbl.lineWeight')} title={t('tbl.lineWeight')} value={tableTool.pen.size} onchange={(e) => setPen({ size: Number(e.currentTarget.value) })}>
        {#each LINE_WEIGHTS as w (w.size)}<option value={w.size}>{w.label}</option>{/each}
      </select>
    </div>
  </div>
  <SplitButton id="tbl-pen-color" size="large" tip={t('tbl.penColor')} label={t('tbl.penColor')}>
    {#snippet face()}<span class="tbl-face"><RibbonIcon name="tblPen" size={32} /><i style="background:{penSwatch}"></i></span>{/snippet}
    {#snippet menu()}<ColorMenu autoLabel="tbl.color.automatic" onpick={(color) => setPen({ color })} />{/snippet}
  </SplitButton>
  <SplitButton id="tbl-borders" size="large" tip={t('tbl.borders')} icon={lastBorder?.icon ?? 'tblBorders'} onclick={() => applyBorders(tableTool.borders)}>
    {#snippet menu()}
      {#each BORDER_ITEMS as b (b.edges)}
        <button class="mi" role="menuitem" onclick={() => applyBorders(b.edges)}><RibbonIcon name={b.icon} size={16} />{t(b.label)}</button>
      {/each}
      <hr />
      <button class="mi" role="menuitem" onclick={() => session.openDialog(DIALOGS.bordersShading)}><RibbonIcon name="tblBordersShading" size={16} />{t('tbl.border.dialog')}</button>
    {/snippet}
  </SplitButton>
  <Button size="large" icon="tblBorderPainter" tip={t('tbl.borderPainter')} on={tableTool.mode === 'painter'} onclick={() => toggleTableTool('painter')} />
</Group>

<BordersShadingDialog />
<TableStyleDialog mode="modify" dialogId={DIALOGS.modifyStyle} />
<TableStyleDialog mode="new" dialogId={DIALOGS.newStyle} />

<style>
  .tbl-options {
    display: grid;
    grid-template-columns: auto auto;
    column-gap: 10px;
  }
  .tbl-gallery {
    gap: 2px;
    padding: 2px 4px;
  }
  .tbl-gallery-menu {
    max-height: 60vh;
    overflow: auto;
  }
  .tbl-gallery-grid {
    display: grid;
    grid-template-columns: repeat(7, 62px);
    gap: 2px;
    padding: 2px 4px;
  }
  .tbl-face {
    position: relative;
    display: inline-flex;
    flex-direction: column;
    align-items: center;
  }
  .tbl-face i {
    display: block;
    width: 24px;
    height: 4px;
    margin-top: -3px;
    border: 0.5px solid rgba(0, 0, 0, 0.15);
  }
  .tbl-pen {
    width: 120px;
  }
  .tbl-presets {
    display: grid;
    grid-template-columns: repeat(3, 56px);
    gap: 4px;
    padding: 4px;
  }
  .tbl-preset {
    display: flex;
    align-items: center;
    height: 22px;
    padding: 0 6px;
    border: 1px solid #dedede;
    border-radius: 3px;
    background: #fff;
    cursor: pointer;
  }
  .tbl-preset:hover {
    border-color: var(--accent);
  }
  .tbl-preset i {
    display: block;
    width: 100%;
  }
</style>
