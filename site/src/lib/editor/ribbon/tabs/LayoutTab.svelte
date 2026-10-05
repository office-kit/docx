<script lang="ts">
  /**
   * Word's Layout tab: Text Direction, Margins, Orientation, Size, Columns,
   * Breaks, Line Numbers, Hyphenation, the Indent / Spacing boxes for the
   * selected paragraphs, and Arrange. Menus check the caret section's current
   * value, as Word's do.
   */
  import { commands, createStyleResolver, paragraphAt } from '@office-kit/docx-editor';
  import { getDocumentSetting, type LineNumbering, type PageMargins, type SectionStart, type SectionTextDirection } from '@office-kit/docx';
  import RibbonIcon from '../../RibbonIcon.svelte';
  import SplitButton from '../SplitButton.svelte';
  import Group from '../Group.svelte';
  import ArrangeGroup from '../ArrangeGroup.svelte';
  import '../../design-layout/design-layout.css';
  import { getSession } from '../../session.svelte';
  import { locale, t, type MessageKey } from '../../i18n/index.svelte';
  import LengthField from '../../design-layout/LengthField.svelte';
  import { formatLength, lengthUnitFor } from '../../design-layout/units';
  import { columnPreset, columnPresetOf, MARGIN_PRESETS, matchesPaper, PAPER_PRESETS, type ColumnPresetId, type MarginPreset, type PaperPreset } from '../../design-layout/presets';
  import { layoutState, type PageSetupTab } from '../../design-layout/layout-state.svelte';
  import PageSetupDialog from '../../design-layout/PageSetupDialog.svelte';
  import ColumnsDialog from '../../design-layout/ColumnsDialog.svelte';
  import LineNumbersDialog from '../../design-layout/LineNumbersDialog.svelte';
  import HyphenationDialog from '../../design-layout/HyphenationDialog.svelte';
  import TextDirectionDialog from '../../design-layout/TextDirectionDialog.svelte';

  const session = getSession();
  const unit = $derived(lengthUnitFor(locale()));
  // A fresh wrapper per edit: the model object itself never changes identity,
  // so deriving from it directly would never recompute after an edit.
  const live = $derived(session.tick >= 0 && session.model ? { model: session.model } : undefined);
  const section = $derived(live ? commands.currentSectionProperties(live.model) : undefined);
  const mirrored = $derived(live ? getDocumentSetting(live.model.doc, 'mirrorMargins').present : false);
  const autoHyphenation = $derived(live ? getDocumentSetting(live.model.doc, 'autoHyphenation').present : false);
  const textWidth = $derived(section ? section.pageSize.widthTwips - section.margins.left - section.margins.right - section.margins.gutter : 0);
  const columnsPreset = $derived(section ? columnPresetOf(section.columns) : undefined);
  const caretFormat = $derived.by(() => {
    const model = live?.model;
    const para = model?.selection ? paragraphAt(model.doc, model.selection.focus) : undefined;
    return para && model ? createStyleResolver(model.doc).paragraph(para) : undefined;
  });
  const canFormat = $derived(!!live?.model.selection);
  // Word's spinner range for paragraph indents and spacing.
  const MAX_INDENT = 31680;
  const MAX_SPACING = 31680;

  const DIRECTIONS: ReadonlyArray<{ value: SectionTextDirection; key: MessageKey }> = [
    { value: 'lrTb', key: 'lay.td.horizontal' },
    { value: 'tbRl', key: 'lay.td.rotate90' },
    { value: 'btLr', key: 'lay.td.rotate270' },
  ];
  const COLUMN_ITEMS: ReadonlyArray<{ id: ColumnPresetId; key: MessageKey }> = [
    { id: 'one', key: 'lay.cols.one' },
    { id: 'two', key: 'lay.cols.two' },
    { id: 'three', key: 'lay.cols.three' },
    { id: 'left', key: 'lay.cols.left' },
    { id: 'right', key: 'lay.cols.right' },
  ];
  const PAGE_BREAKS: ReadonlyArray<{ kind: 'page' | 'column' | 'textWrapping'; key: MessageKey; desc: MessageKey }> = [
    { kind: 'page', key: 'lay.br.page', desc: 'lay.br.pageDesc' },
    { kind: 'column', key: 'lay.br.column', desc: 'lay.br.columnDesc' },
    { kind: 'textWrapping', key: 'lay.br.textWrapping', desc: 'lay.br.textWrappingDesc' },
  ];
  const SECTION_BREAKS: ReadonlyArray<{ type: SectionStart; key: MessageKey; desc: MessageKey }> = [
    { type: 'nextPage', key: 'lay.br.nextPage', desc: 'lay.br.nextPageDesc' },
    { type: 'continuous', key: 'lay.br.continuous', desc: 'lay.br.continuousDesc' },
    { type: 'evenPage', key: 'lay.br.evenPage', desc: 'lay.br.evenPageDesc' },
    { type: 'oddPage', key: 'lay.br.oddPage', desc: 'lay.br.oddPageDesc' },
  ];
  const LINE_NUMBER_ITEMS: ReadonlyArray<{ restart: LineNumbering['restart']; key: MessageKey }> = [
    { restart: 'continuous', key: 'lay.ln.continuous' },
    { restart: 'newPage', key: 'lay.ln.restartPage' },
    { restart: 'newSection', key: 'lay.ln.restartSection' },
  ];

  function sameMargins(a: PageMargins, b: PageMargins | undefined): boolean {
    return !!b && a.top === b.top && a.bottom === b.bottom && a.left === b.left && a.right === b.right;
  }

  function marginDesc(m: PageMargins, mirror: boolean): string {
    const f = (v: number): string => formatLength(v, unit);
    return `${t('lay.top')}: ${f(m.top)}  ${t('lay.bottom')}: ${f(m.bottom)}  ${t(mirror ? 'lay.inside' : 'lay.left')}: ${f(m.left)}  ${t(mirror ? 'lay.outside' : 'lay.right')}: ${f(m.right)}`;
  }

  function applyMargins(m: PageMargins, mirror: boolean): void {
    if (!section) return;
    // The presets keep the section's own header / footer distances and gutter, as Word does.
    const { header, footer, gutter } = section.margins;
    session.apply(commands.pageSetupCommand, { target: 'section', margins: { ...m, header, footer, gutter }, settings: { mirrorMargins: mirror } });
  }

  function applyPaper(p: PaperPreset): void {
    const landscape = section?.pageSize.orientation === 'landscape';
    session.apply(commands.setPageSizeCommand, {
      size: {
        widthTwips: landscape ? p.heightTwips : p.widthTwips,
        heightTwips: landscape ? p.widthTwips : p.heightTwips,
        orientation: landscape ? 'landscape' : 'portrait',
        paperCode: p.code,
      },
    });
  }

  function lineNumbers(restart: LineNumbering['restart']): void {
    const current = section?.lineNumbering;
    session.apply(commands.lineNumbersCommand, { lineNumbering: { start: current?.start ?? 1, countBy: current?.countBy ?? 1, restart, ...(current?.distanceTwips === undefined ? {} : { distanceTwips: current.distanceTwips }) } });
  }

  function openPageSetup(tab: PageSetupTab): void {
    layoutState.pageSetupTab = tab;
    session.openDialog('layout.pageSetup');
  }

  function isMarginPreset(p: MarginPreset): boolean {
    return sameMargins(p.margins, section?.margins) && p.mirror === mirrored;
  }
</script>

<Group label={t('lay.group.pageSetup')}>
  <SplitButton id="layout.textDirection" size="large" tip={t('lay.textDirection')} disabled={!section}>
    {#snippet face()}<RibbonIcon name="layTextDirection" size={32} />{/snippet}
    {#snippet menu()}
      {#each DIRECTIONS as d (d.value)}
        {@const on = section?.textDirection === d.value}
        <button class="mi check" class:checked={on} role="menuitemradio" aria-checked={on} onclick={() => session.apply(commands.textDirectionCommand, { direction: d.value })}>{t(d.key)}</button>
      {/each}
      <hr />
      <button class="mi" role="menuitem" onclick={() => session.openDialog('layout.textDirection')}>{t('lay.td.options')}</button>
    {/snippet}
  </SplitButton>
  <SplitButton id="layout.margins" size="large" tip={t('lay.margins')} disabled={!section}>
    {#snippet face()}<RibbonIcon name="margins" size={32} />{/snippet}
    {#snippet menu()}
      {#if layoutState.lastCustomMargins}
        {@const last = layoutState.lastCustomMargins}
        <button class="mi check described" class:checked={sameMargins(last, section?.margins)} role="menuitemradio" aria-checked={sameMargins(last, section?.margins)} onclick={() => applyMargins(last, mirrored)}>
          <b>{t('lay.margins.last')}</b><small>{marginDesc(last, mirrored)}</small>
        </button>
      {/if}
      {#each MARGIN_PRESETS as p (p.key)}
        {@const on = isMarginPreset(p)}
        <button class="mi check described" class:checked={on} role="menuitemradio" aria-checked={on} onclick={() => applyMargins(p.margins, p.mirror)}>
          <b>{t(p.key)}</b><small>{marginDesc(p.margins, p.mirror)}</small>
        </button>
      {/each}
      <hr />
      <button class="mi" role="menuitem" onclick={() => openPageSetup('margins')}>{t('lay.margins.custom')}</button>
    {/snippet}
  </SplitButton>
  <SplitButton id="layout.orientation" size="large" tip={t('lay.orientation')} disabled={!section}>
    {#snippet face()}<RibbonIcon name="orientation" size={32} />{/snippet}
    {#snippet menu()}
      {#each ['portrait', 'landscape'] as const as o (o)}
        {@const on = section?.pageSize.orientation === o}
        <button class="mi check" class:checked={on} role="menuitemradio" aria-checked={on} onclick={() => session.apply(commands.setOrientationCommand, { orientation: o })}>{t(o === 'portrait' ? 'lay.portrait' : 'lay.landscape')}</button>
      {/each}
    {/snippet}
  </SplitButton>
  <SplitButton id="layout.size" size="large" tip={t('lay.size')} disabled={!section}>
    {#snippet face()}<RibbonIcon name="pageSize" size={32} />{/snippet}
    {#snippet menu()}
      {#each PAPER_PRESETS as p (p.name)}
        {@const on = !!section && matchesPaper(p, section.pageSize.widthTwips, section.pageSize.heightTwips)}
        <button class="mi check described" class:checked={on} role="menuitemradio" aria-checked={on} onclick={() => applyPaper(p)}>
          <b>{p.name}</b><small>{formatLength(p.widthTwips, unit)} × {formatLength(p.heightTwips, unit)}</small>
        </button>
      {/each}
      <hr />
      <button class="mi" role="menuitem" onclick={() => openPageSetup('paper')}>{t('lay.size.more')}</button>
    {/snippet}
  </SplitButton>
  <SplitButton id="layout.columns" size="large" tip={t('lay.columns')} disabled={!section}>
    {#snippet face()}<RibbonIcon name="layColumns" size={32} />{/snippet}
    {#snippet menu()}
      {#each COLUMN_ITEMS as c (c.id)}
        {@const on = columnsPreset === c.id}
        <button class="mi check" class:checked={on} role="menuitemradio" aria-checked={on} onclick={() => section && session.apply(commands.columnsCommand, { columns: { ...columnPreset(c.id, textWidth), separator: section.columns.separator } })}>
          <span class="col-icon col-icon-{c.id}"></span>{t(c.key)}
        </button>
      {/each}
      <hr />
      <button class="mi" role="menuitem" onclick={() => session.openDialog('layout.columns')}>{t('lay.cols.more')}</button>
    {/snippet}
  </SplitButton>
  <div class="col">
    <SplitButton id="layout.breaks" size="mid" icon="breaks" tip={t('lay.breaks')} disabled={!canFormat}>
      {#snippet menu()}
        <div class="menu-head">{t('lay.br.pageBreaks')}</div>
        {#each PAGE_BREAKS as b (b.kind)}
          <button class="mi described" role="menuitem" onclick={() => session.apply(commands.insertBreakCommand, { kind: b.kind })}><b>{t(b.key)}</b><small>{t(b.desc)}</small></button>
        {/each}
        <div class="menu-head">{t('lay.br.sectionBreaks')}</div>
        {#each SECTION_BREAKS as b (b.type)}
          <button class="mi described" role="menuitem" onclick={() => session.apply(commands.insertSectionBreakCommand, { type: b.type })}><b>{t(b.key)}</b><small>{t(b.desc)}</small></button>
        {/each}
      {/snippet}
    </SplitButton>
    <SplitButton id="layout.lineNumbers" size="mid" icon="lineNumbers" tip={t('lay.lineNumbers')} disabled={!section}>
      {#snippet menu()}
        <button class="mi check" class:checked={!section?.lineNumbering} role="menuitemradio" aria-checked={!section?.lineNumbering} onclick={() => session.apply(commands.lineNumbersCommand, { lineNumbering: null })}>{t('lay.none')}</button>
        {#each LINE_NUMBER_ITEMS as item (item.restart)}
          {@const on = section?.lineNumbering?.restart === item.restart}
          <button class="mi check" class:checked={on} role="menuitemradio" aria-checked={on} onclick={() => lineNumbers(item.restart)}>{t(item.key)}</button>
        {/each}
        <button class="mi check" class:checked={session.active(commands.suppressLineNumbersCommand)} role="menuitemcheckbox" aria-checked={session.active(commands.suppressLineNumbersCommand)} disabled={!canFormat} onclick={() => session.apply(commands.suppressLineNumbersCommand, undefined)}>{t('lay.ln.suppress')}</button>
        <hr />
        <button class="mi" role="menuitem" onclick={() => session.openDialog('layout.lineNumbers')}>{t('lay.ln.options')}</button>
      {/snippet}
    </SplitButton>
    <SplitButton id="layout.hyphenation" size="mid" icon="hyphenation" tip={t('lay.hyphenation')} disabled={!live}>
      {#snippet menu()}
        <button class="mi check" class:checked={!autoHyphenation} role="menuitemradio" aria-checked={!autoHyphenation} onclick={() => session.apply(commands.hyphenationCommand, { automatic: false })}>{t('lay.none')}</button>
        <button class="mi check" class:checked={autoHyphenation} role="menuitemradio" aria-checked={autoHyphenation} onclick={() => session.apply(commands.hyphenationCommand, { automatic: true })}>{t('lay.hy.automatic')}</button>
        <hr />
        <button class="mi" role="menuitem" onclick={() => session.openDialog('layout.hyphenation')}>{t('lay.hy.options')}</button>
      {/snippet}
    </SplitButton>
  </div>
</Group>

<Group label={t('lay.group.paragraph')} class="layout-spin-group">
  <div class="spin-block">
    <span class="spin-head">{t('lay.indent')}</span>
    <label class="spin-row"><RibbonIcon name="layIndentLeft" size={16} /><span>{t('lay.left')}:</span>
      <LengthField value={caretFormat?.left ?? 0} {unit} label={t('lay.indentLeft')} min={-MAX_INDENT} max={MAX_INDENT} disabled={!canFormat} onchange={(v) => session.apply(commands.layoutIndentCommand, { left: v })} />
    </label>
    <label class="spin-row"><RibbonIcon name="layIndentRight" size={16} /><span>{t('lay.right')}:</span>
      <LengthField value={caretFormat?.right ?? 0} {unit} label={t('lay.indentRight')} min={-MAX_INDENT} max={MAX_INDENT} disabled={!canFormat} onchange={(v) => session.apply(commands.layoutIndentCommand, { right: v })} />
    </label>
  </div>
  <div class="spin-block">
    <span class="spin-head">{t('lay.spacing')}</span>
    <label class="spin-row"><RibbonIcon name="laySpacingBefore" size={16} /><span>{t('lay.before')}:</span>
      <LengthField value={caretFormat?.before ?? 0} unit="pt" label={t('lay.spacingBefore')} max={MAX_SPACING} disabled={!canFormat} onchange={(v) => session.apply(commands.layoutSpacingCommand, { before: v })} />
    </label>
    <label class="spin-row"><RibbonIcon name="laySpacingAfter" size={16} /><span>{t('lay.after')}:</span>
      <LengthField value={caretFormat?.after ?? 0} unit="pt" label={t('lay.spacingAfter')} max={MAX_SPACING} disabled={!canFormat} onchange={(v) => session.apply(commands.layoutSpacingCommand, { after: v })} />
    </label>
  </div>
</Group>

<ArrangeGroup />

<PageSetupDialog />
<ColumnsDialog />
<LineNumbersDialog />
<HyphenationDialog />
<TextDirectionDialog />
