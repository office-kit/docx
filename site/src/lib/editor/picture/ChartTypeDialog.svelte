<script lang="ts">
  /**
   * Insert Chart (`chart.insert`) and Change Chart Type (`chart.changeType`):
   * chart families on the left, subtypes with a live preview on the right.
   */
  import { chartSvg, commands } from '@office-kit/docx-editor';
  import Dialog from '../Dialog.svelte';
  import { getSession } from '../session.svelte';
  import { t } from '../i18n/index.svelte';
  import { CHART_CHOICES, type ChartChoice, sampleChart, withChartType } from './charts';
  import { applyToSelected, selectedChart } from './state';

  type Props = { mode: 'insert' | 'change' };
  const { mode }: Props = $props();
  const session = getSession();
  const id = $derived(mode === 'insert' ? 'chart.insert' : 'chart.changeType');
  const PREVIEW_W = 320;
  const PREVIEW_H = 200;

  let family = $state(CHART_CHOICES[0]?.family ?? 'column');
  const familyChoices = $derived(CHART_CHOICES.find((f) => f.family === family)?.choices ?? []);
  let choice = $state<ChartChoice | undefined>(CHART_CHOICES[0]?.choices[0]);
  const current = $derived(mode === 'change' ? selectedChart(session) : undefined);
  const spec = $derived(choice ? (current ? withChartType(current.spec, choice) : sampleChart(choice)) : undefined);

  function pickFamily(f: (typeof CHART_CHOICES)[number]): void {
    family = f.family;
    choice = f.choices[0];
  }

  function ok(): void {
    if (!spec) return;
    if (mode === 'insert') session.apply(commands.insertChartCommand, { spec });
    else applyToSelected(session, commands.editChartCommand, { spec });
  }
</script>

<Dialog {id} title={mode === 'insert' ? t('chart.insertTitle') : t('chart.changeType')} onok={ok}>
  <div class="ct">
    <ul class="ct-families" role="listbox" aria-label={t('chart.types')}>
      {#each CHART_CHOICES as f (f.family)}
        <li><button type="button" role="option" aria-selected={family === f.family} class:active={family === f.family} onclick={() => pickFamily(f)}>{t(`chart.fam.${f.family}`)}</button></li>
      {/each}
    </ul>
    <div class="ct-main">
      <div class="ct-subtypes">
        {#each familyChoices as c (c.id)}
          <button type="button" class:active={choice?.id === c.id} onclick={() => (choice = c)}>{t(`chart.${c.id}`)}</button>
        {/each}
      </div>
      {#if spec}<div class="ct-preview">{@html chartSvg(spec, PREVIEW_W, PREVIEW_H)}</div>{/if}
    </div>
  </div>
</Dialog>

<style>
  .ct {
    display: flex;
    gap: 12px;
    min-height: 260px;
  }
  .ct-families {
    margin: 0;
    padding: 0;
    list-style: none;
    min-width: 120px;
  }
  .ct-families button,
  .ct-subtypes button {
    width: 100%;
    padding: 4px 8px;
    border: 1px solid transparent;
    border-radius: 4px;
    background: none;
    text-align: left;
  }
  .ct-subtypes {
    display: flex;
    flex-wrap: wrap;
    gap: 4px;
  }
  .ct-subtypes button {
    width: auto;
    border-color: var(--chrome-line);
  }
  .ct-families button.active,
  .ct-subtypes button.active {
    border-color: #f0a30a;
    background: #fff5e0;
  }
  .ct-preview {
    margin-top: 8px;
    background: #fff;
    border: 1px solid var(--chrome-line);
  }
</style>
