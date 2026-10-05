<script lang="ts">
  /**
   * Word's Comments pane: every comment as a card, in document order. A card
   * is edited in place; clicking one moves the caret to its anchor.
   */
  import { comments } from '@office-kit/docx';
  import { commands } from '@office-kit/docx-editor';
  import RibbonIcon from '../RibbonIcon.svelte';
  import { getSession } from '../session.svelte';
  import { t } from '../i18n/index.svelte';
  import { goToMark, newComment } from '../review-actions';

  const session = getSession();
  const prefs = session.prefs;
  const list = $derived(session.tick >= 0 && session.model ? comments(session.model.doc) : []);

  function select(id: number): void {
    prefs.activeComment = id;
    goToMark(session, 'comment', String(id));
  }

  function save(id: number, previous: string, e: Event): void {
    const text = (e.currentTarget as HTMLTextAreaElement).value;
    if (text !== previous) session.apply(commands.editCommentCommand, { id, text });
  }

  /** Put the cursor in the card of a comment just created (Review ▸ New Comment). */
  function focusWhen(node: HTMLTextAreaElement, active: boolean): void {
    if (active) node.focus();
  }

  function when(iso: string | undefined): string {
    return iso ? new Date(iso).toLocaleString() : '';
  }
</script>

<div class="pane-head">
  <span>{t('review.commentsTitle')}</span>
  <button class="pane-close" onclick={() => (session.pane.right = null)} aria-label={t('group.close')}><RibbonIcon name="close" size={14} /></button>
</div>
<button class="pane-button" onclick={() => newComment(session)} disabled={!session.model?.selection}>{t('review.newComment')}</button>
{#if list.length === 0}
  <p class="pane-note">{t('review.noComments')}</p>
{/if}
{#each list as c (c.id)}
  <div class="card" class:active={prefs.activeComment === c.id} role="group" aria-label={c.author}>
    <button class="who" onclick={() => select(c.id)}>
      <span class="avatar" aria-hidden="true">{c.initials ?? c.author.slice(0, 1)}</span>
      <span class="meta"><strong>{c.author}</strong><small>{when(c.date)}</small></span>
    </button>
    <textarea
      rows="2"
      value={c.text}
      placeholder={t('review.commentPlaceholder')}
      aria-label={c.author}
      onfocus={() => (prefs.activeComment = c.id)}
      onchange={(e) => save(c.id, c.text, e)}
      use:focusWhen={prefs.activeComment === c.id && c.text === ''}
    ></textarea>
    <button class="del" onclick={() => session.apply(commands.deleteCommentCommand, { id: c.id })} title={t('review.deleteComment')} aria-label={t('review.deleteComment')}><RibbonIcon name="close" size={12} /></button>
  </div>
{/each}

<style>
  .card { position: relative; display: flex; flex-direction: column; gap: 6px; padding: 8px; border: 1px solid var(--chrome-line); border-radius: 8px; background: #fff; }
  .card.active { border-color: var(--accent); box-shadow: 0 0 0 1px var(--accent); }
  .who { display: flex; align-items: center; gap: 8px; border: none; background: none; padding: 0; text-align: left; cursor: pointer; }
  .avatar { display: inline-flex; align-items: center; justify-content: center; width: 24px; height: 24px; border-radius: 50%; background: #5b7fbd; color: #fff; font-size: 11px; }
  .meta { display: flex; flex-direction: column; }
  .meta small { color: var(--muted); }
  textarea { resize: vertical; border: 1px solid var(--control-line); border-radius: 4px; font: inherit; padding: 4px; }
  .del { position: absolute; top: 6px; right: 6px; border: none; background: none; cursor: pointer; color: var(--icon); border-radius: 4px; padding: 3px; }
  .del:hover { background: var(--hover); }
</style>
