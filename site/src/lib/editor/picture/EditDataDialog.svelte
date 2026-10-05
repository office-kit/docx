<script lang="ts">
  /**
   * Chart Design ▸ Edit Data: the chart's sheet as a grid (categories down,
   * series across, like the embedded workbook Word opens). OK rewrites the
   * chart part's caches and the embedded workbook together.
   */
  import { untrack } from 'svelte';
  import { commands } from '@office-kit/docx-editor';
  import Dialog from '../Dialog.svelte';
  import { getSession } from '../session.svelte';
  import { t } from '../i18n/index.svelte';
  import { applyToSelected, selectedChart } from './state';

  const session = getSession();
  let title = $state('');
  let categories = $state<string[]>([]);
  let series = $state<Array<{ name: string; values: string[] }>>([]);
  let error = $state('');

  $effect(() => {
    if (session.dialog !== 'chart.data') return;
    const chart = untrack(() => selectedChart(session));
    if (!chart) return;
    title = chart.spec.title ?? '';
    categories = [...chart.spec.categories];
    series = chart.spec.series.map((s) => ({ name: s.name, values: s.values.map((v) => (v === null ? '' : String(v))) }));
    error = '';
  });

  function addCategory(): void {
    categories.push(`${t('chart.category')} ${categories.length + 1}`);
    for (const s of series) s.values.push('');
  }
  function removeCategory(i: number): void {
    categories.splice(i, 1);
    for (const s of series) s.values.splice(i, 1);
  }
  function addSeries(): void {
    series.push({ name: `${t('chart.series')} ${series.length + 1}`, values: categories.map(() => '') });
  }
  function removeSeries(i: number): void {
    series.splice(i, 1);
  }

  function ok(): boolean {
    const chart = selectedChart(session);
    if (!chart) return true;
    const colors = chart.spec.series.map((s) => s.color);
    const parsed = series.map((s, i) => {
      const values = s.values.map((v) => (v.trim() === '' ? null : Number(v)));
      const color = colors[i];
      return color ? { name: s.name, values, color } : { name: s.name, values };
    });
    if (parsed.some((s) => s.values.some((v) => v !== null && !Number.isFinite(v)))) {
      error = t('chart.numbersOnly');
      return false;
    }
    if (!parsed.length || !categories.length) {
      error = t('chart.needData');
      return false;
    }
    const { title: _title, ...rest } = chart.spec;
    applyToSelected(session, commands.editChartCommand, {
      spec: {
        ...rest,
        ...(title.trim() && { title: title.trim() }),
        categories: [...categories],
        series: parsed,
      },
    });
    return true;
  }
</script>

<Dialog id="chart.data" title={t('chart.editData')} onok={ok}>
  <label class="field">{t('chart.title')} <input bind:value={title} /></label>
  <div class="sheet-wrap">
    <table class="sheet">
      <thead>
        <tr>
          <th></th>
          {#each series as s, j (j)}
            <th><input bind:value={s.name} aria-label={t('chart.series')} /><button type="button" title={t('chart.remove')} aria-label={t('chart.remove')} onclick={() => removeSeries(j)}>×</button></th>
          {/each}
          <th><button type="button" onclick={addSeries}>+ {t('chart.series')}</button></th>
        </tr>
      </thead>
      <tbody>
        {#each categories as _c, i (i)}
          <tr>
            <th><input bind:value={categories[i]} aria-label={t('chart.category')} /><button type="button" title={t('chart.remove')} aria-label={t('chart.remove')} onclick={() => removeCategory(i)}>×</button></th>
            {#each series as s, j (j)}<td><input bind:value={s.values[i]} inputmode="decimal" aria-label="{s.name} {categories[i]}" /></td>{/each}
          </tr>
        {/each}
        <tr><th><button type="button" onclick={addCategory}>+ {t('chart.category')}</button></th></tr>
      </tbody>
    </table>
  </div>
  {#if error}<p class="pane-note" role="alert">{error}</p>{/if}
</Dialog>

<style>
  .sheet-wrap {
    max-height: 50vh;
    overflow: auto;
    margin-top: 8px;
  }
  .sheet {
    border-collapse: collapse;
    background: #fff;
  }
  .sheet th,
  .sheet td {
    border: 1px solid #d4d4d4;
    padding: 0;
  }
  .sheet th {
    background: #f3f3f3;
    white-space: nowrap;
  }
  .sheet input {
    width: 90px;
    border: none;
    padding: 3px 4px;
    background: transparent;
    font: inherit;
  }
  .sheet td input {
    text-align: right;
  }
  .sheet button {
    border: none;
    background: none;
    cursor: pointer;
  }
</style>
