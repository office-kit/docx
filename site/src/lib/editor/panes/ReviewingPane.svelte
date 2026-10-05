<script lang="ts">
  /**
   * Word's Reviewing pane (docked left, as Word's vertical pane is): a count
   * of the document's revisions and every revision with its author, kind and
   * text. Clicking one moves to it; each can be accepted or rejected.
   */
  import { revisions, type RevisionKind } from '@office-kit/docx';
  import { commands } from '@office-kit/docx-editor';
  import RibbonIcon from '../RibbonIcon.svelte';
  import { getSession } from '../session.svelte';
  import { t, type MessageKey } from '../i18n/index.svelte';
  import { goToMark } from '../review-actions';

  const session = getSession();
  const list = $derived(session.tick >= 0 && session.model ? revisions(session.model.doc) : []);

  const KIND: Readonly<Record<RevisionKind, MessageKey>> = {
    insert: 'review.kindInsert',
    delete: 'review.kindDelete',
    paragraphInsert: 'review.kindParagraphInsert',
    paragraphDelete: 'review.kindParagraphDelete',
    format: 'review.kindFormat',
    rowInsert: 'review.kindRowInsert',
    rowDelete: 'review.kindRowDelete',
  };

  // Excerpts are cut like Word's pane, which shows the start of long changes.
  const EXCERPT_LENGTH = 80;
  const excerpt = (text: string): string => (text.length > EXCERPT_LENGTH ? `${text.slice(0, EXCERPT_LENGTH)}…` : text);
</script>

<div class="pane-head">
  <span>{t('review.paneTitle')}</span>
  <button class="pane-close" onclick={() => (session.pane.left = null)} aria-label={t('group.close')}><RibbonIcon name="close" size={14} /></button>
</div>
<p class="pane-note">{list.length ? t('review.revisionCount').replace('{n}', String(list.length)) : t('review.noRevisions')}</p>
{#each list as r (r.id)}
  <div class="rev">
    <button class="rev-main" onclick={() => goToMark(session, 'revision', r.id)}>
      <span class="rev-head"><strong>{r.author ?? ''}</strong> {t(KIND[r.kind])}</span>
      {#if r.date}<small>{new Date(r.date).toLocaleString()}</small>{/if}
      {#if r.text}<span class="rev-text" class:del={r.kind === 'delete' || r.kind === 'paragraphDelete' || r.kind === 'rowDelete'}>{excerpt(r.text)}</span>{/if}
    </button>
    <span class="rev-actions">
      <button onclick={() => session.apply(commands.acceptRevisionsCommand, { ids: [r.id] })} title={t('review.acceptThis')} aria-label={t('review.acceptThis')}><RibbonIcon name="accept" size={14} /></button>
      <button onclick={() => session.apply(commands.rejectRevisionsCommand, { ids: [r.id] })} title={t('review.rejectThis')} aria-label={t('review.rejectThis')}><RibbonIcon name="reject" size={14} /></button>
    </span>
  </div>
{/each}

<style>
  .rev { display: flex; gap: 4px; padding: 6px; border-bottom: 1px solid var(--chrome-line); }
  .rev-main { flex: 1; display: flex; flex-direction: column; gap: 2px; border: none; background: none; text-align: left; cursor: pointer; padding: 0; min-width: 0; }
  .rev-main small { color: var(--muted); }
  .rev-text { overflow-wrap: anywhere; }
  .rev-text.del { text-decoration: line-through; }
  .rev-actions { display: flex; flex-direction: column; gap: 2px; }
  .rev-actions button { border: none; background: none; cursor: pointer; padding: 3px; border-radius: 4px; }
  .rev-actions button:hover { background: var(--hover); }
</style>
