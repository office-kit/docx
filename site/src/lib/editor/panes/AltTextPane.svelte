<script lang="ts">
  /** Word's Alt Text pane for the selected picture or chart (`wp:docPr/@descr`). */
  import { commands } from '@office-kit/docx-editor';
  import RibbonIcon from '../RibbonIcon.svelte';
  import { getSession } from '../session.svelte';
  import { t } from '../i18n/index.svelte';
  import { applyToSelected, selected } from '../picture/state';

  const session = getSession();
  const info = $derived(selected(session)?.info);
</script>

<div class="pane-head">
  <span>{t('pic.altText')}</span>
  <button class="pane-close" onclick={() => (session.pane.right = null)} aria-label={t('pic.closePane')}><RibbonIcon name="close" size={14} /></button>
</div>
{#if info}
  <p class="pane-note">{t('altText.help')}</p>
  <textarea
    class="pane-input alt-text"
    rows="6"
    value={info.description}
    aria-label={t('pic.altText')}
    onchange={(e) => applyToSelected(session, commands.setImageAltCommand, { descr: e.currentTarget.value })}
  ></textarea>
{:else}
  <p class="pane-note">{t('altText.none')}</p>
{/if}

<style>
  .alt-text {
    height: auto;
    padding: 6px 8px;
    resize: vertical;
  }
</style>
