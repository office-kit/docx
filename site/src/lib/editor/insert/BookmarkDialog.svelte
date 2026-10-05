<script lang="ts">
  /**
   * Insert ▸ Bookmark: add a bookmark around the selection, delete one, or go
   * to one; the list sorts by name or location and can include Word's hidden
   * bookmarks (`_Ref…`, `_Toc…`).
   */
  import { isValidBookmarkName } from '@office-kit/docx';
  import { caretAt, commands, documentBookmarks } from '@office-kit/docx-editor';
  import Dialog from '../Dialog.svelte';
  import { getSession } from '../session.svelte';
  import { t } from '../i18n/index.svelte';
  import { DIALOG } from './state.svelte';

  const session = getSession();
  let name = $state('');
  let sortBy = $state<'name' | 'location'>('name');
  let showHidden = $state(false);

  const all = $derived(
    session.dialog === DIALOG.bookmark && session.version >= 0 && session.model ? documentBookmarks(session.model.doc) : [],
  );
  const listed = $derived.by(() => {
    const visible = all.filter((b) => showHidden || !b.hidden);
    return sortBy === 'name' ? visible.toSorted((a, b) => a.name.localeCompare(b.name)) : visible;
  });
  const existing = $derived(all.find((b) => b.name === name));
  const valid = $derived(isValidBookmarkName(name));

  $effect(() => {
    if (session.dialog === DIALOG.bookmark) name = '';
  });

  function add(): void {
    if (!valid) return;
    // Word's Add on an existing name moves that bookmark to the selection.
    if (existing) session.apply(commands.deleteBookmarkCommand, { name });
    session.apply(commands.addBookmarkAtSelectionCommand, { name });
    if (session.status === '') session.dialog = null;
  }

  function remove(): void {
    if (existing) session.apply(commands.deleteBookmarkCommand, { name });
  }

  function goTo(): void {
    const model = session.model;
    if (!existing || !model) return;
    model.setSelection(caretAt({ block: existing.block, inline: 0, offset: 0 }));
    session.changed();
  }
</script>

<Dialog id={DIALOG.bookmark} title={t('ref.bookmark')}>
  <label class="stack">{t('ins.bm.name')}<input type="text" bind:value={name} maxlength="40" /></label>
  <div class="listbox" role="listbox" aria-label={t('ins.bm.name')}>
    {#each listed as b (b.name)}
      <button type="button" role="option" aria-selected={b.name === name} class:selected={b.name === name} onclick={() => (name = b.name)} ondblclick={goTo}>{b.name}</button>
    {/each}
  </div>
  <div class="dialog-row">
    <fieldset>
      <legend>{t('ins.bm.sortBy')}</legend>
      <label><input type="radio" name="bm-sort" value="name" bind:group={sortBy} /> {t('ins.bm.byName')}</label>
      <label><input type="radio" name="bm-sort" value="location" bind:group={sortBy} /> {t('ins.bm.byLocation')}</label>
    </fieldset>
    <label><input type="checkbox" bind:checked={showHidden} /> {t('ins.bm.hidden')}</label>
  </div>
  {#if name && !valid}<p class="error-note">{t('ins.bm.invalid')}</p>{/if}
  {#if session.status}<p class="error-note">{session.status}</p>{/if}
  <div class="dialog-row">
    <button type="button" class="push" disabled={!valid} onclick={add}>{t('ins.bm.add')}</button>
    <button type="button" class="push" disabled={!existing} onclick={remove}>{t('ins.bm.delete')}</button>
    <button type="button" class="push" disabled={!existing} onclick={goTo}>{t('ins.bm.goTo')}</button>
  </div>
</Dialog>
