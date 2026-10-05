<script lang="ts">
  /** Word's Navigation pane, docked left: search the document and replace. */
  import { commands } from '@office-kit/docx-editor';
  import RibbonIcon from '../RibbonIcon.svelte';
  import { getSession } from '../session.svelte';
  import { t } from '../i18n/index.svelte';

  const session = getSession();
  let replaceValue = $state('');
  let findStatus = $state('');

  function replaceAll(): void {
    if (!session.search) return;
    const n = session.apply(commands.replaceAllCommand, { query: session.search, replacement: replaceValue });
    if (n === undefined) return;
    findStatus = `Replaced ${n} occurrence${n === 1 ? '' : 's'}.`;
  }
</script>

<div class="pane-head">
  <span>{t('find.title')}</span>
  <button class="pane-close" onclick={() => (session.pane.left = null)} aria-label={t('find.close')}><RibbonIcon name="close" size={14} /></button>
</div>
<label class="pane-search">
  <RibbonIcon name="search" size={14} />
  <input placeholder={t('find.find')} bind:value={session.search} aria-label={t('find.find')} />
</label>
<input class="pane-input" placeholder={t('find.replaceWith')} bind:value={replaceValue} aria-label={t('find.replaceWith')} />
<button class="pane-button" onclick={replaceAll} disabled={!session.search}>{t('find.replaceAll')}</button>
<p class="pane-note" aria-live="polite">{findStatus}</p>
