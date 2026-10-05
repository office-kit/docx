<script lang="ts">
  /**
   * Review ▸ Word Count. Lines are not listed: they depend on line layout,
   * which only the browser knows, and Word recomputes them on every save.
   */
  import { wordCount } from '@office-kit/docx';
  import Dialog from '../Dialog.svelte';
  import { getSession } from '../session.svelte';
  import { t } from '../i18n/index.svelte';

  const session = getSession();
  let includeNotes = $state(false);
  const counts = $derived(
    session.dialog === 'wordCount' && session.tick >= 0 && session.model
      ? wordCount(session.model.doc, { includeTextboxesAndNotes: includeNotes })
      : undefined,
  );
</script>

<Dialog id="wordCount" title={t('review.wordCount')} onok={() => {}} okLabel={t('group.close')}>
  {#if counts}
    <div class="stats-head">{t('review.wcStatistics')}</div>
    <table class="stats">
      <tbody>
        <tr><th>{t('review.wcPages')}</th><td>{session.pageCount.toLocaleString()}</td></tr>
        <tr><th>{t('review.wcWords')}</th><td>{counts.words.toLocaleString()}</td></tr>
        <tr><th>{t('review.wcChars')}</th><td>{counts.characters.toLocaleString()}</td></tr>
        <tr><th>{t('review.wcCharsSpaces')}</th><td>{counts.charactersWithSpaces.toLocaleString()}</td></tr>
        <tr><th>{t('review.wcParagraphs')}</th><td>{counts.paragraphs.toLocaleString()}</td></tr>
      </tbody>
    </table>
    <label class="check-item"><input type="checkbox" bind:checked={includeNotes} />{t('review.wcInclude')}</label>
  {/if}
</Dialog>

<style>
  .stats-head { font-weight: 600; }
  .stats { border-collapse: collapse; min-width: 300px; }
  .stats th { text-align: left; font-weight: normal; padding: 2px 24px 2px 0; }
  .stats td { text-align: right; padding: 2px 0; font-variant-numeric: tabular-nums; }
</style>
