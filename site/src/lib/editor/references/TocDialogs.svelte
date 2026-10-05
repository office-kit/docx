<script lang="ts">
  /** Custom Table of Contents and Insert Table of Figures. */
  import { captionLabels, type TabLeader } from '@office-kit/docx';
  import { commands } from '@office-kit/docx-editor';
  import Dialog from '../Dialog.svelte';
  import { getSession } from '../session.svelte';
  import { t } from '../i18n/index.svelte';
  import { pageProvider } from './state.svelte';
  import { LEADERS } from './options';

  const session = getSession();

  let pageNumbers = $state(true);
  let rightAlign = $state(true);
  let leader = $state<TabLeader>('dot');
  let hyperlinks = $state(true);
  let showLevels = $state(3);
  let figureLabel = $state('Figure');
  let includeLabel = $state(true);

  const labels = $derived(session.tick >= 0 && session.model ? captionLabels(session.model.doc) : []);
  const MAX_TOC_LEVEL = 9;

  function common() {
    const pageOf = pageProvider(session);
    return {
      pageNumbers,
      rightAlignPageNumbers: rightAlign,
      tabLeader: rightAlign ? leader : 'none',
      hyperlinks,
      ...(pageOf ? { pageOf } : {}),
    } as const;
  }
</script>

{#snippet pageNumberOptions()}
  <label><input type="checkbox" bind:checked={pageNumbers} /> {t('ref.showPageNumbers')}</label>
  <label><input type="checkbox" bind:checked={rightAlign} disabled={!pageNumbers} /> {t('ref.rightAlign')}</label>
  <label class="field">{t('ref.tabLeader')}
    <select bind:value={leader} disabled={!pageNumbers || !rightAlign}>
      {#each LEADERS as l (l.value)}<option value={l.value}>{l.sample || t('ref.none')}</option>{/each}
    </select>
  </label>
  <label><input type="checkbox" bind:checked={hyperlinks} /> {t('ref.useHyperlinks')}</label>
{/snippet}

<Dialog id="references.toc" title={t('ref.customTocTitle')} onok={() => { session.apply(commands.insertTocCommand, { ...common(), levels: { from: 1, to: showLevels } }); }}>
  {@render pageNumberOptions()}
  <label class="field">{t('ref.showLevels')} <input type="number" min="1" max={MAX_TOC_LEVEL} bind:value={showLevels} /></label>
</Dialog>

<Dialog id="references.tof" title={t('ref.tableOfFigures')} onok={() => { session.apply(commands.insertTocCommand, { ...common(), captionLabel: figureLabel, includeLabelAndNumber: includeLabel }); }}>
  {@render pageNumberOptions()}
  <label class="field">{t('ref.captionLabel')}
    <select bind:value={figureLabel}>
      {#each labels as label (label)}<option value={label}>{label}</option>{/each}
    </select>
  </label>
  <label><input type="checkbox" bind:checked={includeLabel} /> {t('ref.includeLabel')}</label>
</Dialog>
