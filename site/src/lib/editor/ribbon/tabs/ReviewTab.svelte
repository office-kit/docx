<script lang="ts">
  /**
   * Word for Mac's Review tab: Proofing, Accessibility, Language, Comments,
   * Tracking, Changes, Compare and Protect. Editor, Thesaurus, Read Aloud,
   * Translate, Resolve / Reply and Block Authors are left out: they need a
   * cloud service, an extension namespace (w15 comment threads) or a
   * SharePoint server.
   */
  import { commands, revisionIdsAtSelection } from '@office-kit/docx-editor';
  import { documentProtection, revisions, writeProtection } from '@office-kit/docx';
  import Button from '../Button.svelte';
  import Group from '../Group.svelte';
  import SplitButton from '../SplitButton.svelte';
  import { getSession, type MarkupMode } from '../../session.svelte';
  import { t, type MessageKey } from '../../i18n/index.svelte';
  import { commentAtCaret, goToAdjacent, newComment, revisionIdsShown } from '../../review-actions';

  const session = getSession();
  const prefs = session.prefs;
  const hasSelection = $derived(session.tick >= 0 && !!session.model?.selection);
  const hasRevisions = $derived(session.tick >= 0 && !!session.model && revisions(session.model.doc).length > 0);
  const protectedDoc = $derived(session.tick >= 0 && !!session.model && !!documentProtection(session.model.doc)?.enforced);
  const readOnlyRecommended = $derived(
    session.tick >= 0 && !!session.model && !!writeProtection(session.model.doc)?.recommended,
  );
  const changeHere = $derived(session.tick >= 0 && !!session.model && revisionIdsAtSelection(session.model).length > 0);

  const MARKUP: readonly [MarkupMode, MessageKey][] = [
    ['simple', 'review.markupSimple'],
    ['all', 'review.markupAll'],
    ['none', 'review.markupNone'],
    ['original', 'review.markupOriginal'],
  ];

  function toggleSpelling(): void {
    session.spellcheck = !session.spellcheck;
    session.status = session.spellcheck ? '' : t('review.spellingOff');
  }

  function deleteComment(): void {
    const id = commentAtCaret(session);
    if (id !== undefined) session.apply(commands.deleteCommentCommand, { id });
  }

  /** Accept / Reject This Change; "and Move to Next" then jumps to the next change. */
  function decide(accept: boolean, moveNext: boolean): void {
    session.apply(accept ? commands.acceptRevisionsCommand : commands.rejectRevisionsCommand, {});
    if (moveNext) goToAdjacent(session, 'revision', 1);
  }

  function decideShown(accept: boolean): void {
    if (!session.model) return;
    const ids = revisionIdsShown(session);
    session.apply(accept ? commands.acceptRevisionsCommand : commands.rejectRevisionsCommand, { ids });
  }

  /** The face of Accept / Reject: this change when there is one, otherwise the next. */
  function decideFace(accept: boolean): void {
    if (changeHere) decide(accept, true);
    else goToAdjacent(session, 'revision', 1);
  }

  function toggleReadOnly(): void {
    session.apply(commands.setWriteProtectionCommand, readOnlyRecommended ? undefined : { recommended: true });
  }
</script>

<Group label={t('group.proofing')}>
  <Button size="large" icon="reviewSpelling" tip={t('review.spelling')} on={session.spellcheck} onclick={toggleSpelling} />
  <Button size="large" icon="reviewWordCount" tip={t('review.wordCount')} onclick={() => session.openDialog('wordCount')} />
</Group>
<Group label={t('group.accessibility')}>
  <Button size="large" icon="reviewAccessibility" tip={t('review.checkAccessibility')} on={session.pane.right === 'accessibility'} onclick={() => session.togglePane('right', 'accessibility')} />
</Group>
<Group label={t('group.language')}>
  <SplitButton id="review-language" size="large" icon="reviewLanguage" tip={t('review.language')}>
    {#snippet menu()}
      <button class="mi" onclick={() => session.openDialog('language')} disabled={!hasSelection}>{t('review.setLanguage')}</button>
    {/snippet}
  </SplitButton>
</Group>
<Group label={t('group.comments')}>
  <Button size="large" icon="reviewNewComment" tip={t('review.newComment')} onclick={() => newComment(session)} disabled={!hasSelection} />
  <SplitButton id="review-delete-comment" size="large" icon="reviewDeleteComment" tip={t('review.deleteComment')} onclick={deleteComment}>
    {#snippet menu()}
      <button class="mi" onclick={deleteComment}>{t('review.deleteComment')}</button>
      <button class="mi" onclick={() => session.apply(commands.deleteAllCommentsCommand, undefined)}>{t('review.deleteAllComments')}</button>
    {/snippet}
  </SplitButton>
  <div class="rows">
    <Button size="mid" icon="reviewPrevComment" tip={t('review.previousComment')} onclick={() => goToAdjacent(session, 'comment', -1)} />
    <Button size="mid" icon="reviewNextComment" tip={t('review.nextComment')} onclick={() => goToAdjacent(session, 'comment', 1)} />
  </div>
  <Button size="large" icon="reviewShowComments" tip={t('review.showComments')} on={session.pane.right === 'comments'} onclick={() => session.togglePane('right', 'comments')} />
</Group>
<Group label={t('group.tracking')}>
  <Button
    size="large"
    icon="reviewTrackChanges"
    tip={session.enabled(commands.toggleTrackChangesCommand) ? t('review.trackChanges') : t('review.trackingForced')}
    label={t('review.trackChanges')}
    on={session.active(commands.toggleTrackChangesCommand)}
    disabled={!session.enabled(commands.toggleTrackChangesCommand)}
    onclick={() => session.apply(commands.toggleTrackChangesCommand, undefined)}
  />
  <div class="rows">
    <label class="field" title={t('review.displayForReview')}>
      <select value={session.markup} onchange={(e) => (session.markup = e.currentTarget.value as MarkupMode)} aria-label={t('review.displayForReview')}>
        {#each MARKUP as [mode, key] (mode)}<option value={mode}>{t(key)}</option>{/each}
      </select>
    </label>
    <SplitButton id="review-markup" size="mid" icon="reviewMarkupOptions" tip={t('review.showMarkup')}>
      {#snippet menu()}
        <button class="mi check" class:checked={prefs.showComments} onclick={() => (prefs.showComments = !prefs.showComments)}>{t('review.markupComments')}</button>
        <button class="mi check" class:checked={prefs.showInsertionsDeletions} onclick={() => (prefs.showInsertionsDeletions = !prefs.showInsertionsDeletions)}>{t('review.markupInsDel')}</button>
        <button class="mi check" class:checked={prefs.showFormatting} onclick={() => (prefs.showFormatting = !prefs.showFormatting)}>{t('review.markupFormatting')}</button>
        <hr />
        <button class="mi" onclick={() => session.openDialog('trackChangesOptions')}>{t('review.trackOptions')}</button>
      {/snippet}
    </SplitButton>
    <Button size="mid" icon="reviewReviewingPane" tip={t('review.reviewingPane')} on={session.pane.left === 'reviewing'} onclick={() => session.togglePane('left', 'reviewing')} />
  </div>
</Group>
<Group label={t('group.changes')}>
  <SplitButton id="review-accept" size="large" icon="reviewAccept" tip={t('review.accept')} onclick={() => decideFace(true)} disabled={!hasRevisions}>
    {#snippet menu()}
      <button class="mi" onclick={() => decide(true, true)} disabled={!changeHere}>{t('review.acceptAndNext')}</button>
      <button class="mi" onclick={() => decide(true, false)} disabled={!changeHere}>{t('review.acceptThis')}</button>
      <button class="mi" onclick={() => decideShown(true)}>{t('review.acceptShown')}</button>
      <button class="mi" onclick={() => session.apply(commands.acceptAllRevisionsCommand, undefined)}>{t('review.acceptAll')}</button>
      <button class="mi" onclick={() => session.apply(commands.acceptAllAndStopCommand, undefined)}>{t('review.acceptAllStop')}</button>
    {/snippet}
  </SplitButton>
  <SplitButton id="review-reject" size="large" icon="reviewReject" tip={t('review.reject')} onclick={() => decideFace(false)} disabled={!hasRevisions}>
    {#snippet menu()}
      <button class="mi" onclick={() => decide(false, true)} disabled={!changeHere}>{t('review.rejectAndNext')}</button>
      <button class="mi" onclick={() => decide(false, false)} disabled={!changeHere}>{t('review.rejectThis')}</button>
      <button class="mi" onclick={() => decideShown(false)}>{t('review.rejectShown')}</button>
      <button class="mi" onclick={() => session.apply(commands.rejectAllRevisionsCommand, undefined)}>{t('review.rejectAll')}</button>
      <button class="mi" onclick={() => session.apply(commands.rejectAllAndStopCommand, undefined)}>{t('review.rejectAllStop')}</button>
    {/snippet}
  </SplitButton>
  <div class="rows">
    <Button size="mid" icon="reviewPrevChange" tip={t('review.previousChange')} onclick={() => goToAdjacent(session, 'revision', -1)} disabled={!hasRevisions} />
    <Button size="mid" icon="reviewNextChange" tip={t('review.nextChange')} onclick={() => goToAdjacent(session, 'revision', 1)} disabled={!hasRevisions} />
  </div>
</Group>
<Group label={t('group.compare')}>
  <SplitButton id="review-compare" size="large" icon="reviewCompare" tip={t('review.compare')}>
    {#snippet menu()}
      <button class="mi" onclick={() => session.openDialog('compare')}>{t('review.compareDocuments')}</button>
    {/snippet}
  </SplitButton>
</Group>
<Group label={t('group.protect')}>
  <Button size="large" icon="reviewProtect" tip={t('review.restrictEditing')} on={session.pane.right === 'restrictEditing' || protectedDoc} onclick={() => session.togglePane('right', 'restrictEditing')} />
  <Button size="large" icon="reviewReadOnly" tip={t('review.alwaysReadOnly')} on={readOnlyRecommended} onclick={toggleReadOnly} />
</Group>
