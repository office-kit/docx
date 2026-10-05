<script lang="ts">
  /**
   * Word for Mac's Citations pane: the citation style, the document's sources
   * (double-click or Insert to cite one), and the master list's sources that
   * can be added to the document.
   */
  import { bibliographySources, type BibliographySource } from '@office-kit/docx';
  import { commands } from '@office-kit/docx-editor';
  import RibbonIcon from '../RibbonIcon.svelte';
  import { getSession } from '../session.svelte';
  import { t } from '../i18n/index.svelte';
  import { masterList, sourceEditor } from '../references/state.svelte';

  const session = getSession();
  let query = $state('');

  const current = $derived(session.tick >= 0 && session.model ? bibliographySources(session.model.doc) : []);
  const currentTags = $derived(new Set(current.map((s) => s.tag)));
  const matches = (s: BibliographySource): boolean => {
    const q = query.trim().toLowerCase();
    return !q || label(s).toLowerCase().includes(q) || s.tag.toLowerCase().includes(q);
  };
  const label = (s: BibliographySource): string =>
    [s.corporateAuthor ?? s.authors?.map((a) => a.last).join(', '), s.fields?.Year && `(${s.fields.Year})`, s.fields?.Title]
      .filter(Boolean)
      .join(' ') || s.tag;

  function cite(tag: string): void {
    session.apply(commands.insertCitationCommand, { tag });
  }

  function addToDocument(source: BibliographySource): void {
    session.apply(commands.setSourcesCommand, { sources: [...current, source] });
  }

  function newSource(): void {
    sourceEditor.editing = null;
    sourceEditor.citeAfterSave = false;
    session.openDialog('references.source');
  }
</script>

<div class="pane-head">
  <span>{t('ref.citations')}</span>
  <button class="pane-close" onclick={() => (session.pane.right = null)} aria-label={t('ref.closePane')}><RibbonIcon name="close" size={14} /></button>
</div>
<label class="pane-search">
  <RibbonIcon name="search" size={14} />
  <input placeholder={t('ref.searchSources')} bind:value={query} aria-label={t('ref.searchSources')} />
</label>
<div class="menu-head">{t('ref.currentList')}</div>
{#each current.filter(matches) as s (s.tag)}
  <button class="mi source" onclick={() => cite(s.tag)} disabled={!session.model?.selection} title={t('ref.insertCitation')}>{label(s)}</button>
{:else}
  <p class="pane-note">{t('ref.noSources')}</p>
{/each}
{#if masterList.sources.some((s) => !currentTags.has(s.tag) && matches(s))}
  <div class="menu-head">{t('ref.masterList')}</div>
  {#each masterList.sources.filter((s) => !currentTags.has(s.tag) && matches(s)) as s (s.tag)}
    <button class="mi source" onclick={() => addToDocument(s)} title={t('ref.copyRight')}>{label(s)}</button>
  {/each}
{/if}
<button class="pane-button" onclick={newSource}>{t('ref.addNewSource')}</button>
<button class="pane-button" onclick={() => session.openDialog('references.sources')}>{t('ref.manageSources')}</button>

<style>
  .source {
    white-space: normal;
  }
</style>
