<script lang="ts">
  /**
   * Word's Selection Pane: every floating and inline object, top of the stack
   * first; click to select, the eye to show / hide, double-click to rename.
   */
  import RibbonIcon from '../RibbonIcon.svelte';
  import { getSession } from '../session.svelte';
  import { t } from '../i18n/index.svelte';
  import { OBJECT_LISTERS, type PaneObject } from '../ribbon/arrange';

  const session = getSession();
  const rows = $derived(OBJECT_LISTERS.flatMap((list) => list(session)));
  let editing = $state<string | null>(null);

  function rename(row: PaneObject, name: string): void {
    editing = null;
    const trimmed = name.trim();
    if (trimmed && trimmed !== row.name) row.rename?.(trimmed);
  }
</script>

<div class="pane-head">
  <span>{t('arr.selectionPane')}</span>
  <button class="pane-close" onclick={() => (session.pane.right = null)} aria-label={t('pic.closePane')}><RibbonIcon name="close" size={14} /></button>
</div>
<div class="sel-actions">
  <button class="pane-button" disabled={!rows.length} onclick={() => rows.forEach((r) => r.hidden && r.setHidden?.(false))}>{t('selPane.showAll')}</button>
  <button class="pane-button" disabled={!rows.length} onclick={() => rows.forEach((r) => !r.hidden && r.setHidden?.(true))}>{t('selPane.hideAll')}</button>
</div>
{#if rows.length}
  <ul class="sel-list">
    {#each rows as row (row.key)}
      <li class:hidden={row.hidden}>
        {#if editing === row.key}
          <!-- svelte-ignore a11y_autofocus -->
          <input class="pane-input" value={row.name} autofocus aria-label={t('selPane.rename')}
            onkeydown={(e) => { if (e.key === 'Enter') rename(row, e.currentTarget.value); else if (e.key === 'Escape') editing = null; }}
            onblur={(e) => rename(row, e.currentTarget.value)} />
        {:else}
          <button class="sel-name" onclick={() => row.select()} ondblclick={() => row.rename && (editing = row.key)}>{row.name}</button>
        {/if}
        <button class="sel-eye" disabled={!row.setHidden} aria-pressed={!row.hidden} title={row.hidden ? t('selPane.show') : t('selPane.hide')} aria-label={row.hidden ? t('selPane.show') : t('selPane.hide')}
          onclick={() => row.setHidden?.(!row.hidden)}>
          <svg viewBox="0 0 20 20" width="16" height="16" fill="none" stroke="currentColor" aria-hidden="true">
            <path d="M2 10s3-5 8-5 8 5 8 5-3 5-8 5-8-5-8-5z" />
            {#if row.hidden}<path d="M3 17L17 3" />{:else}<circle cx="10" cy="10" r="2.5" />{/if}
          </svg>
        </button>
      </li>
    {/each}
  </ul>
{:else}
  <p class="pane-note">{t('selPane.empty')}</p>
{/if}

<style>
  .sel-actions {
    display: flex;
    gap: 6px;
  }
  .sel-actions .pane-button {
    flex: 1;
  }
  .sel-list {
    margin: 0;
    padding: 0;
    list-style: none;
  }
  .sel-list li {
    display: flex;
    align-items: center;
    gap: 4px;
  }
  .sel-list li.hidden .sel-name {
    color: var(--muted);
  }
  .sel-name {
    flex: 1;
    padding: 4px 6px;
    border: none;
    background: none;
    text-align: left;
    cursor: pointer;
  }
  .sel-name:hover {
    background: rgba(0, 0, 0, 0.06);
  }
  .sel-eye {
    border: none;
    background: none;
    cursor: pointer;
  }
</style>
