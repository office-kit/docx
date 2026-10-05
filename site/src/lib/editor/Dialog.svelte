<script lang="ts">
  /**
   * A Word-style modal dialog (Font, Paragraph, Insert Table …), shown while
   * `session.dialog === id`. OK runs `onok`; returning `false` keeps the dialog
   * open (e.g. invalid input that the dialog reports itself).
   */
  import type { Snippet } from 'svelte';
  import { getSession } from './session.svelte';
  import { t } from './i18n/index.svelte';

  type Props = {
    id: string;
    title: string;
    children: Snippet;
    onok?: () => boolean | void;
    okLabel?: string;
    okDisabled?: boolean;
  };
  const { id, title, children, onok, okLabel, okDisabled = false }: Props = $props();
  const session = getSession();

  function close(): void {
    session.dialog = null;
  }

  function ok(): void {
    if (onok?.() === false) return;
    close();
  }

  function onkeydown(e: KeyboardEvent): void {
    if (e.key === 'Escape') close();
  }
</script>

{#if session.dialog === id}
  <div class="dialog-backdrop" role="presentation" onkeydown={onkeydown}>
    <div class="dialog" role="dialog" aria-modal="true" aria-label={title}>
     <form onsubmit={(e) => { e.preventDefault(); ok(); }}>
      <div class="dialog-title">{title}</div>
      <div class="dialog-body">{@render children()}</div>
      <div class="dialog-actions">
        <button type="button" onclick={close}>{t('dialog.cancel')}</button>
        {#if onok}<button type="submit" class="default" disabled={okDisabled}>{okLabel ?? t('dialog.ok')}</button>{/if}
      </div>
     </form>
    </div>
  </div>
{/if}
