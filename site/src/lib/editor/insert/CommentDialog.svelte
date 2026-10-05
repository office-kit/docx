<script lang="ts">
  /**
   * Insert ▸ Comment: the comment text plus the author Word takes from the
   * user's name (remembered in this browser), with initials and a timestamp.
   */
  import { commands } from '@office-kit/docx-editor';
  import Dialog from '../Dialog.svelte';
  import { getSession } from '../session.svelte';
  import { t } from '../i18n/index.svelte';
  import { DIALOG, insertUi } from './state.svelte';

  const session = getSession();
  let author = $state('');
  let text = $state('');

  $effect(() => {
    if (session.dialog !== DIALOG.comment) return;
    author = insertUi.userName;
    text = '';
  });

  function ok(): boolean {
    const name = author.trim();
    if (name !== insertUi.userName) insertUi.rememberUserName(name);
    session.apply(commands.addCommentCommand, {
      author: name,
      initials: insertUi.initials(),
      text,
      // Word stores comment times without fractional seconds.
      date: new Date().toISOString().replace(/\.\d+Z$/, 'Z'),
    });
    return session.status === '';
  }
</script>

<Dialog id={DIALOG.comment} title={t('ins.comment.new')} onok={ok} okDisabled={!author.trim() || !text.trim()}>
  <label class="stack">{t('ins.comment.author')}<input type="text" bind:value={author} /></label>
  <label class="stack">{t('ins.comment.text')}<textarea rows="4" bind:value={text}></textarea></label>
  {#if session.status}<p class="error-note">{session.status}</p>{/if}
</Dialog>
