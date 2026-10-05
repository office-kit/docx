<script lang="ts">
  /** Mark Index Entry, Index, Mark Citation and Table of Authorities. */
  import { untrack } from 'svelte';
  import { AUTHORITY_CATEGORIES, type TabLeader } from '@office-kit/docx';
  import { commands } from '@office-kit/docx-editor';
  import Dialog from '../Dialog.svelte';
  import { getSession } from '../session.svelte';
  import { t } from '../i18n/index.svelte';
  import { LEADERS } from './options';
  import { pageProvider } from './state.svelte';

  const session = getSession();
  const MAX_INDEX_COLUMNS = 4;

  // Mark Entry and Mark Citation start from the selected text, as in Word.
  $effect(() => {
    const id = session.dialog;
    if (id !== 'references.markEntry' && id !== 'references.markCitation') return;
    untrack(() => {
      const selected = session.model ? commands.selectedText(session.model).trim() : '';
      if (id === 'references.markEntry') main = selected;
      else {
        longCitation = selected;
        shortCitation = selected;
      }
    });
  });

  // --- Mark Entry --------------------------------------------------------------------
  let main = $state('');
  let subentry = $state('');
  let pageOption = $state<'current' | 'crossReference' | 'range'>('current');
  let crossReference = $state('See ');
  let bookmark = $state('');
  let bold = $state(false);
  let italic = $state(false);

  function entry() {
    return {
      main: main.trim(),
      ...(subentry.trim() ? { subentry: subentry.trim() } : {}),
      ...(pageOption === 'crossReference' ? { crossReference } : {}),
      ...(pageOption === 'range' && bookmark ? { pageRangeBookmark: bookmark } : {}),
      bold,
      italic,
    };
  }

  function mark(markAll: boolean): boolean {
    if (!main.trim()) return false;
    const n = session.apply(commands.markIndexEntryCommand, { entry: entry(), ...(markAll ? { markAll: main.trim() } : {}) });
    if (n !== undefined) session.status = t('ref.markedEntries').replace('{n}', String(n));
    return true;
  }

  // --- Index ---------------------------------------------------------------------------
  let indexType = $state<'indented' | 'runIn'>('indented');
  let columns = $state(2);
  let rightAlign = $state(false);
  let indexLeader = $state<TabLeader>('dot');

  function insertIndex(): void {
    const pageOf = pageProvider(session);
    session.apply(commands.insertIndexCommand, {
      type: indexType,
      columns,
      rightAlignPageNumbers: rightAlign,
      tabLeader: indexLeader,
      ...(pageOf ? { pageOf } : {}),
    });
  }

  // --- Mark Citation -------------------------------------------------------------------
  let longCitation = $state('');
  let shortCitation = $state('');
  let category = $state(1);

  function markCitation(): boolean {
    if (!longCitation.trim()) return false;
    session.apply(commands.markCitationCommand, {
      longCitation: longCitation.trim(),
      ...(shortCitation.trim() ? { shortCitation: shortCitation.trim() } : {}),
      category,
    });
    return true;
  }

  // --- Table of Authorities -----------------------------------------------------------
  let toaCategory = $state<number | 'all'>('all');
  let passim = $state(true);
  let toaLeader = $state<TabLeader>('dot');

  function insertToa(): void {
    const pageOf = pageProvider(session);
    session.apply(commands.insertToaCommand, { category: toaCategory, passim, tabLeader: toaLeader, ...(pageOf ? { pageOf } : {}) });
  }
</script>

{#snippet leaderSelect(get: () => TabLeader, set: (v: TabLeader) => void, disabled: boolean)}
  <label class="field">{t('ref.tabLeader')}
    <select bind:value={get, set} {disabled}>
      {#each LEADERS as l (l.value)}<option value={l.value}>{l.sample || t('ref.none')}</option>{/each}
    </select>
  </label>
{/snippet}

<Dialog id="references.markEntry" title={t('ref.markEntryTitle')} okLabel={t('ref.mark')} onok={() => mark(false)} okDisabled={!main.trim()}>
  <label class="field">{t('ref.mainEntry')} <input bind:value={main} size="30" /></label>
  <label class="field">{t('ref.subentry')} <input bind:value={subentry} size="30" /></label>
  <fieldset>
    <legend>{t('ref.options')}</legend>
    <label><input type="radio" bind:group={pageOption} value="crossReference" /> {t('ref.crossRef')}</label>
    <input bind:value={crossReference} disabled={pageOption !== 'crossReference'} aria-label={t('ref.crossRef')} />
    <label><input type="radio" bind:group={pageOption} value="current" /> {t('ref.currentPage')}</label>
    <label><input type="radio" bind:group={pageOption} value="range" /> {t('ref.pageRange')}</label>
    <input bind:value={bookmark} disabled={pageOption !== 'range'} placeholder={t('ref.bookmark')} aria-label={t('ref.bookmark')} />
  </fieldset>
  <fieldset>
    <legend>{t('ref.pageNumberFormat')}</legend>
    <label><input type="checkbox" bind:checked={bold} /> {t('ref.bold')}</label>
    <label><input type="checkbox" bind:checked={italic} /> {t('ref.italic')}</label>
  </fieldset>
  <button type="button" class="push" onclick={() => { if (mark(true)) session.dialog = null; }} disabled={!main.trim()}>{t('ref.markAll')}</button>
</Dialog>

<Dialog id="references.index" title={t('ref.indexTitle')} onok={insertIndex}>
  <fieldset>
    <legend>{t('ref.type')}</legend>
    <label><input type="radio" bind:group={indexType} value="indented" /> {t('ref.indented')}</label>
    <label><input type="radio" bind:group={indexType} value="runIn" /> {t('ref.runIn')}</label>
  </fieldset>
  <label class="field">{t('ref.columns')} <input type="number" min="1" max={MAX_INDEX_COLUMNS} bind:value={columns} /></label>
  <label><input type="checkbox" bind:checked={rightAlign} /> {t('ref.rightAlign')}</label>
  {@render leaderSelect(() => indexLeader, (v) => (indexLeader = v), !rightAlign)}
</Dialog>

<Dialog id="references.markCitation" title={t('ref.markCitationTitle')} okLabel={t('ref.mark')} onok={markCitation} okDisabled={!longCitation.trim()}>
  <label class="field">{t('ref.longCitation')} <textarea bind:value={longCitation} rows="3" cols="36"></textarea></label>
  <label class="field">{t('ref.category')}
    <select bind:value={category}>
      {#each AUTHORITY_CATEGORIES as c, i (c)}<option value={i + 1}>{t(`ref.cat.${c}`)}</option>{/each}
    </select>
  </label>
  <label class="field">{t('ref.shortCitation')} <input bind:value={shortCitation} size="36" /></label>
</Dialog>

<Dialog id="references.toa" title={t('ref.toaTitle')} onok={insertToa}>
  <label class="field">{t('ref.category')}
    <select bind:value={toaCategory}>
      <option value="all">{t('ref.all')}</option>
      {#each AUTHORITY_CATEGORIES as c, i (c)}<option value={i + 1}>{t(`ref.cat.${c}`)}</option>{/each}
    </select>
  </label>
  <label><input type="checkbox" bind:checked={passim} /> {t('ref.passim')}</label>
  {@render leaderSelect(() => toaLeader, (v) => (toaLeader = v), false)}
</Dialog>
