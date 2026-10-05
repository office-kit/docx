<script lang="ts">
  /**
   * Track Changes Options: how insertions, deletions and changed lines are
   * drawn, and the user name and initials new changes and comments carry.
   * These are preferences of this editor, not of the document.
   */
  import { setReviewer } from '@office-kit/docx-editor';
  import Dialog from '../Dialog.svelte';
  import { getSession } from '../session.svelte';
  import { t, type MessageKey } from '../i18n/index.svelte';
  import type { ChangedLines, DeletionMark, InsertionMark } from '../review-prefs.svelte';

  const session = getSession();
  const prefs = session.prefs;

  const INSERTIONS: readonly [InsertionMark, MessageKey][] = [
    ['underline', 'review.mark.underline'],
    ['doubleUnderline', 'review.mark.doubleUnderline'],
    ['bold', 'review.mark.bold'],
    ['italic', 'review.mark.italic'],
    ['colorOnly', 'review.mark.colorOnly'],
  ];
  const DELETIONS: readonly [DeletionMark, MessageKey][] = [
    ['strikethrough', 'review.mark.strikethrough'],
    ['doubleStrikethrough', 'review.mark.doubleStrikethrough'],
    ['hidden', 'review.mark.hidden'],
  ];
  const LINES: readonly [ChangedLines, MessageKey][] = [
    ['outside', 'review.line.outside'],
    ['left', 'review.line.left'],
    ['right', 'review.line.right'],
    ['none', 'review.line.none'],
  ];

  let insertion = $state<InsertionMark>('underline');
  let deletion = $state<DeletionMark>('strikethrough');
  let lines = $state<ChangedLines>('outside');
  let userName = $state('');
  let initials = $state('');

  $effect(() => {
    if (session.dialog !== 'trackChangesOptions') return;
    insertion = prefs.insertionMark;
    deletion = prefs.deletionMark;
    lines = prefs.changedLines;
    userName = prefs.userName;
    initials = prefs.userInitials;
  });

  function apply(): void {
    prefs.insertionMark = insertion;
    prefs.deletionMark = deletion;
    prefs.changedLines = lines;
    prefs.setUser(userName, initials);
    if (session.model) setReviewer(session.model, { author: prefs.userName, initials: prefs.userInitials });
  }
</script>

<Dialog id="trackChangesOptions" title={t('review.trackOptions').replace('…', '')} onok={apply}>
  <fieldset>
    <legend>{t('review.markupInsDel')}</legend>
    <div class="grid2">
      <label for="tco-ins">{t('review.tcoInsertions')}</label>
      <select id="tco-ins" bind:value={insertion}>{#each INSERTIONS as [v, k] (v)}<option value={v}>{t(k)}</option>{/each}</select>
      <label for="tco-del">{t('review.tcoDeletions')}</label>
      <select id="tco-del" bind:value={deletion}>{#each DELETIONS as [v, k] (v)}<option value={v}>{t(k)}</option>{/each}</select>
      <label for="tco-lines">{t('review.tcoChangedLines')}</label>
      <select id="tco-lines" bind:value={lines}>{#each LINES as [v, k] (v)}<option value={v}>{t(k)}</option>{/each}</select>
    </div>
  </fieldset>
  <fieldset>
    <legend>{t('review.reviewers')}</legend>
    <div class="grid2">
      <label for="tco-user">{t('review.tcoUser')}</label>
      <input id="tco-user" bind:value={userName} />
      <label for="tco-initials">{t('review.tcoInitials')}</label>
      <input id="tco-initials" bind:value={initials} maxlength="9" />
    </div>
  </fieldset>
</Dialog>

<style>
  .grid2 { display: grid; grid-template-columns: auto 1fr; gap: 6px 10px; align-items: center; }
  .grid2 select, .grid2 input { height: 22px; font: inherit; }
</style>
