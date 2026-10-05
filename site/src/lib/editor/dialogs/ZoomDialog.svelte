<script lang="ts">
  /** View ▸ Zoom: fixed percentages, fitted widths, whole / many pages, or a percent. */
  import Dialog from '../Dialog.svelte';
  import { getSession, MAX_ZOOM, MIN_ZOOM } from '../session.svelte';
  import { t, type MessageKey } from '../i18n/index.svelte';
  import { fitZoom, type ZoomFit } from '../review-actions';

  const session = getSession();

  type Choice = '200' | '100' | '75' | ZoomFit | 'percent';
  const FIXED: readonly ('200' | '100' | '75')[] = ['200', '100', '75'];
  const FITTED: readonly [ZoomFit, MessageKey][] = [
    ['pageWidth', 'view.pageWidth'],
    ['textWidth', 'view.textWidth'],
    ['wholePage', 'view.wholePage'],
    ['multiplePages', 'view.multiplePages'],
  ];

  let choice = $state<Choice>('percent');
  let percent = $state(100);

  $effect(() => {
    if (session.dialog !== 'zoom') return;
    percent = Math.round(session.zoom * 100);
    choice = FIXED.find((p) => Number(p) === percent) ?? 'percent';
  });

  function pickFixed(p: (typeof FIXED)[number]): void {
    choice = p;
    percent = Number(p);
  }

  function pickFit(fit: ZoomFit): void {
    choice = fit;
    const zoom = fitZoom(session, fit);
    if (zoom !== undefined) percent = Math.round(zoom * 100);
  }

  function apply(): void {
    session.prefs.pagesAcross = choice === 'multiplePages' ? 2 : 1;
    session.setZoom(percent / 100);
  }
</script>

<Dialog id="zoom" title={t('view.zoom')} onok={apply}>
  <fieldset>
    <legend>{t('view.zoomTo')}</legend>
    <div class="choices">
      {#each FIXED as p (p)}
        <label class="check-item"><input type="radio" name="zoom" checked={choice === p} onchange={() => pickFixed(p)} />{p}%</label>
      {/each}
      {#each FITTED as [fit, key] (fit)}
        <label class="check-item"><input type="radio" name="zoom" checked={choice === fit} onchange={() => pickFit(fit)} />{t(key)}</label>
      {/each}
    </div>
  </fieldset>
  <label class="field">
    {t('view.percent')}
    <input type="number" min={MIN_ZOOM * 100} max={MAX_ZOOM * 100} bind:value={percent} oninput={() => (choice = 'percent')} />%
  </label>
</Dialog>

<style>
  .choices { display: grid; grid-template-columns: repeat(2, auto); gap: 4px 24px; }
</style>
