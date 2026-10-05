<script lang="ts">
  /**
   * Review ▸ Restrict Editing: formatting and editing restrictions, everyone's
   * exceptions, enforcing with an optional password, Stop Protection, and the
   * password to modify (`w:writeProtection`).
   */
  import { documentProtection, editableRanges, writeProtection, type ProtectionEdit } from '@office-kit/docx';
  import { commands } from '@office-kit/docx-editor';
  import RibbonIcon from '../RibbonIcon.svelte';
  import { getSession } from '../session.svelte';
  import { t, type MessageKey } from '../i18n/index.svelte';

  const session = getSession();
  const protection = $derived(session.tick >= 0 && session.model ? documentProtection(session.model.doc) : undefined);
  const ranges = $derived(session.tick >= 0 && session.model ? editableRanges(session.model.doc) : []);
  const write = $derived(session.tick >= 0 && session.model ? writeProtection(session.model.doc) : undefined);
  const enforced = $derived(!!protection?.enforced);

  const EDITS: readonly [Exclude<ProtectionEdit, 'none'>, MessageKey][] = [
    ['readOnly', 'review.rsReadOnly'],
    ['trackedChanges', 'review.rsTracked'],
    ['comments', 'review.rsComments'],
    ['forms', 'review.rsForms'],
  ];

  let limitFormatting = $state(false);
  let limitEditing = $state(true);
  let edit = $state<Exclude<ProtectionEdit, 'none'>>('readOnly');
  let starting = $state(false);
  let password = $state('');
  let confirm = $state('');
  let stopPassword = $state('');
  let modifyPassword = $state('');
  let message = $state('');

  // Exceptions only make sense where the rest of the document is locked.
  const exceptionsApply = $derived(limitEditing && (edit === 'readOnly' || edit === 'comments'));

  function start(): void {
    if (password !== confirm) {
      message = t('review.rsMismatch');
      return;
    }
    session.apply(commands.protectDocumentCommand, {
      edit: limitEditing ? edit : 'none',
      formatting: limitFormatting,
      ...(password ? { password } : {}),
    });
    starting = false;
    password = confirm = message = '';
  }

  function stop(): void {
    session.apply(commands.unprotectDocumentCommand, { password: stopPassword });
    stopPassword = '';
  }

  function setModifyPassword(): void {
    session.apply(commands.setWriteProtectionCommand, {
      recommended: !!write?.recommended,
      ...(modifyPassword ? { password: modifyPassword } : {}),
    });
    modifyPassword = '';
  }
</script>

<div class="pane-head">
  <span>{t('review.restrictEditing')}</span>
  <button class="pane-close" onclick={() => (session.pane.right = null)} aria-label={t('group.close')}><RibbonIcon name="close" size={14} /></button>
</div>

{#if enforced}
  <p class="pane-note">{t('review.rsEnforced')}</p>
  {#if protection?.hasPassword}
    <label class="rs-field">{t('review.rsPasswordToStop')}<input type="password" class="pane-input" bind:value={stopPassword} /></label>
  {/if}
  <button class="pane-button" onclick={stop}>{t('review.rsStop')}</button>
{:else}
  <section>
    <h3>{t('review.rsFormatting')}</h3>
    <label class="check-item"><input type="checkbox" bind:checked={limitFormatting} />{t('review.rsLimitFormatting')}</label>
  </section>
  <section>
    <h3>{t('review.rsEditing')}</h3>
    <label class="check-item"><input type="checkbox" bind:checked={limitEditing} />{t('review.rsAllowOnly')}</label>
    <select class="pane-input" bind:value={edit} disabled={!limitEditing} aria-label={t('review.rsAllowOnly')}>
      {#each EDITS as [value, key] (value)}<option {value}>{t(key)}</option>{/each}
    </select>
  </section>
  {#if exceptionsApply}
    <section>
      <h3>{t('review.rsExceptions')}</h3>
      <p class="pane-note">{t('review.rsExceptionsHelp')}</p>
      <button class="pane-button" onclick={() => session.apply(commands.addEditableRangeCommand, { group: 'everyone' })} disabled={!session.enabled(commands.addEditableRangeCommand)}>{t('review.rsEveryone')}</button>
    </section>
  {/if}
  <section>
    <h3>{t('review.rsStart')}</h3>
    {#if starting}
      <label class="rs-field">{t('review.rsPassword')}<input type="password" class="pane-input" bind:value={password} /></label>
      <label class="rs-field">{t('review.rsConfirm')}<input type="password" class="pane-input" bind:value={confirm} /></label>
      {#if message}<p class="pane-note" role="alert">{message}</p>{/if}
      <div class="rs-actions">
        <button class="pane-button" onclick={() => { starting = false; message = ''; }}>{t('dialog.cancel')}</button>
        <button class="pane-button" onclick={start}>{t('dialog.ok')}</button>
      </div>
    {:else}
      <button class="pane-button" onclick={() => (starting = true)} disabled={!limitEditing && !limitFormatting}>{t('review.rsYesStart')}</button>
    {/if}
  </section>
{/if}

{#if ranges.length}
  <section>
    <h3>{t('review.rsExceptions')}</h3>
    {#each ranges as range (range.id)}
      <div class="rs-range">
        <span>{range.editor ?? t('review.rsEveryone')} · ¶{range.block + 1}</span>
        <button class="pane-button" onclick={() => session.apply(commands.removeEditableRangeCommand, { id: range.id })} disabled={enforced}>{t('review.rsRemoveException')}</button>
      </div>
    {/each}
  </section>
{/if}

<section>
  <h3>{t('review.protect')}</h3>
  <label class="check-item">
    <input type="checkbox" checked={!!write?.recommended} onchange={(e) => session.apply(commands.setWriteProtectionCommand, e.currentTarget.checked ? { recommended: true } : undefined)} />{t('review.alwaysReadOnly')}
  </label>
  <label class="rs-field">{t('review.rsModifyPassword')}<input type="password" class="pane-input" bind:value={modifyPassword} /></label>
  <button class="pane-button" onclick={setModifyPassword} disabled={!modifyPassword}>{t('dialog.ok')}</button>
</section>

<style>
  section { display: flex; flex-direction: column; gap: 6px; }
  h3 { margin: 6px 0 0; font-size: 13px; }
  .rs-field { display: flex; flex-direction: column; gap: 4px; }
  .rs-actions { display: flex; gap: 6px; justify-content: flex-end; }
  .rs-actions button { min-width: 72px; }
  .rs-range { display: flex; align-items: center; justify-content: space-between; gap: 6px; }
</style>
