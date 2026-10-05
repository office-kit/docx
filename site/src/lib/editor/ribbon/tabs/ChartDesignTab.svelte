<script lang="ts">
  /**
   * Word's Chart Design tab for the selected chart: Chart Layouts (Add Chart
   * Element), Chart Styles (Change Colors), Data and Type. Word puts Arrange
   * and Size on the chart's Format tab; with no separate Format tab here they
   * follow on this one.
   */
  import { commands } from '@office-kit/docx-editor';
  import type { ChartSpec, LegendPosition } from '@office-kit/docx';
  import ArrangeGroup from '../ArrangeGroup.svelte';
  import Button from '../Button.svelte';
  import Group from '../Group.svelte';
  import SplitButton from '../SplitButton.svelte';
  import ChartTypeDialog from '../../picture/ChartTypeDialog.svelte';
  import EditDataDialog from '../../picture/EditDataDialog.svelte';
  import { getSession } from '../../session.svelte';
  import { t, type MessageKey } from '../../i18n/index.svelte';
  import { applyToSelected, EMU_PER_CM, selected, selectedChart } from '../../picture/state';
  import { CHART_PALETTES } from '../../picture/presets';

  const session = getSession();
  const chart = $derived(selectedChart(session));
  const info = $derived(selected(session)?.info);
  const LEGENDS: ReadonlyArray<readonly [LegendPosition, MessageKey]> = [
    ['none', 'chart.none'],
    ['right', 'chart.legendRight'],
    ['top', 'chart.legendTop'],
    ['left', 'chart.legendLeft'],
    ['bottom', 'chart.legendBottom'],
  ];

  function edit(f: (spec: ChartSpec) => ChartSpec): void {
    if (chart) applyToSelected(session, commands.editChartCommand, { spec: f(chart.spec) });
  }

  function setTitle(on: boolean): void {
    edit(({ title, ...rest }) => (on ? { ...rest, title: title ?? t('chart.title') } : rest));
  }

  /** Switch Row/Column: categories become series and series become categories. */
  function switchRowColumn(spec: ChartSpec): ChartSpec {
    return {
      ...spec,
      categories: spec.series.map((s) => s.name),
      series: spec.categories.map((name, i) => ({ name, values: spec.series.map((s) => s.values[i] ?? null) })),
    };
  }

  function setPalette(colors: readonly string[]): void {
    edit((spec) => ({
      ...spec,
      series: spec.series.map((s, i) => ({ ...s, color: colors[i % colors.length] ?? s.color ?? '4472C4' })),
    }));
  }

  const cm = (emu: number): number => Math.round((emu / EMU_PER_CM) * 100) / 100;
  function setSize(dim: 'h' | 'w', value: number): void {
    if (!info || !(value > 0)) return;
    const emu = value * EMU_PER_CM;
    applyToSelected(session, commands.resizeImageCommand, {
      cxEmu: dim === 'w' ? emu : info.widthEmu,
      cyEmu: dim === 'h' ? emu : info.heightEmu,
    });
  }
</script>

<Group label={t('chart.layouts')}>
  <SplitButton id="chart.addElement" size="large" icon="addChartElement" tip={t('chart.addElement')} disabled={!chart}>
    {#snippet menu()}
      <div class="menu-head">{t('chart.title')}</div>
      <button class="mi check" class:checked={!chart?.spec.title} onclick={() => setTitle(false)}>{t('chart.none')}</button>
      <button class="mi check" class:checked={!!chart?.spec.title} onclick={() => setTitle(true)}>{t('chart.aboveChart')}</button>
      <div class="menu-head">{t('chart.dataLabels')}</div>
      <button class="mi check" class:checked={!chart?.spec.dataLabels} onclick={() => edit((s) => ({ ...s, dataLabels: false }))}>{t('chart.none')}</button>
      <button class="mi check" class:checked={!!chart?.spec.dataLabels} onclick={() => edit((s) => ({ ...s, dataLabels: true }))}>{t('chart.show')}</button>
      <div class="menu-head">{t('chart.legend')}</div>
      {#each LEGENDS as [pos, key] (pos)}
        <button class="mi check" class:checked={chart?.spec.legend === pos} onclick={() => edit((s) => ({ ...s, legend: pos }))}>{t(key)}</button>
      {/each}
    {/snippet}
  </SplitButton>
</Group>
<Group label={t('chart.styles')}>
  <SplitButton id="chart.colors" size="large" icon="changeColors" tip={t('chart.changeColors')} disabled={!chart}>
    {#snippet menu()}
      {#each CHART_PALETTES as p (p.id)}
        <button class="mi" title={p.id} aria-label={p.id} onclick={() => setPalette(p.colors)}>
          {#each p.colors as c (c)}<i class="swatch" style="background:#{c}"></i>{/each}
        </button>
      {/each}
    {/snippet}
  </SplitButton>
</Group>
<Group label={t('chart.data')}>
  <Button size="large" icon="switchRowCol" tip={t('chart.switchRowCol')} disabled={!chart} onclick={() => edit(switchRowColumn)} />
  <Button size="large" icon="editData" tip={t('chart.editData')} disabled={!chart} onclick={() => session.openDialog('chart.data')} />
</Group>
<Group label={t('chart.type')}>
  <Button size="large" icon="changeChartType" tip={t('chart.changeType')} disabled={!chart} onclick={() => session.openDialog('chart.changeType')} />
</Group>
<ArrangeGroup compact />
<Group label={t('pic.size')}>
  <div class="rows chart-size">
    <label class="field">{t('pic.height')} <input type="number" min="0.01" step="0.1" value={info ? cm(info.heightEmu) : ''} disabled={!info} onchange={(e) => setSize('h', e.currentTarget.valueAsNumber)} /> cm</label>
    <label class="field">{t('pic.width')} <input type="number" min="0.01" step="0.1" value={info ? cm(info.widthEmu) : ''} disabled={!info} onchange={(e) => setSize('w', e.currentTarget.valueAsNumber)} /> cm</label>
  </div>
</Group>
<EditDataDialog />
<ChartTypeDialog mode="change" />

<style>
  .swatch {
    display: inline-block;
    width: 14px;
    height: 14px;
  }
  .chart-size input {
    width: 56px;
  }
</style>
