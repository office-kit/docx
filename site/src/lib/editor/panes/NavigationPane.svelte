<script lang="ts">
  /**
   * Word's Navigation pane, docked left: the Headings outline, and search with
   * Results (every result highlighted in the page), Find Next / Previous,
   * Replace / Replace All, and the Advanced Find and Replace dialog.
   */
  import { onDestroy } from 'svelte';
  import { bodyParagraphs, caretAt, commands, createStyleResolver, paragraphPlainText, type FindMatch } from '@office-kit/docx-editor';
  import RibbonIcon from '../RibbonIcon.svelte';
  import { getSession } from '../session.svelte';
  import { t } from '../i18n/index.svelte';
  import { homeState } from '../ribbon/tabs/home/state.svelte';
  import { clearPaint, findOptions, goTo, next, paintMatches, search } from '../ribbon/tabs/home/find-ui';

  const session = getSession();
  const home = homeState(session);
  // Characters of context Word shows around a result.
  const CONTEXT = 24;
  const OUTLINE_LEVELS = 9;

  let view = $state<'headings' | 'results'>('results');
  let showOptions = $state(false);
  let findStatus = $state('');

  const options = $derived(findOptions(session, home));
  const result = $derived(session.version >= 0 ? search(session, options) : { matches: [] });
  const matches = $derived(result.matches);
  const texts = $derived.by(() => {
    if (session.version < 0 || !session.model) return new Map<string, string>();
    return new Map(bodyParagraphs(session.model.doc).map(({ para, at }) => [key(at), paragraphPlainText(para)]));
  });
  const headings = $derived.by(() => {
    if (session.version < 0 || !session.model) return [];
    const styles = createStyleResolver(session.model.doc);
    return bodyParagraphs(session.model.doc).flatMap(({ para, at }) => {
      const level = styles.paragraph(para).outlineLevel;
      return level !== undefined && level < OUTLINE_LEVELS && !at.cell ? [{ at, level, text: paragraphPlainText(para) }] : [];
    });
  });

  function key(at: FindMatch['at']): string {
    return `${at.block}/${at.cell?.row ?? ''}/${at.cell?.col ?? ''}/${at.para ?? 0}`;
  }

  function context(m: FindMatch): { before: string; hit: string; after: string } {
    const text = texts.get(key(m.at)) ?? '';
    return { before: text.slice(Math.max(0, m.start - CONTEXT), m.start), hit: text.slice(m.start, m.end), after: text.slice(m.end, m.end + CONTEXT) };
  }

  // Repaint after the canvas re-renders the version this search ran on.
  $effect(() => {
    const found = matches;
    if (view !== 'results') {
      clearPaint();
      return;
    }
    requestAnimationFrame(() => paintMatches(session, found));
  });
  onDestroy(clearPaint);

  function replaceParams(): commands.ReplaceParams {
    const format = home.find.replaceFormat;
    return { find: options, replacement: home.find.replace, ...(Object.keys(format).length > 0 ? { format } : {}) };
  }

  function replaceOne(): void {
    session.apply(commands.replaceNextCommand, replaceParams());
  }

  function replaceAll(): void {
    const n = session.apply(commands.findReplaceAllCommand, replaceParams());
    if (n !== undefined) findStatus = t('home.find.replaced').replace('{n}', String(n));
  }

  function goToHeading(at: FindMatch['at']): void {
    session.model?.setSelection(caretAt({ ...at, inline: 0, offset: 0 }));
    session.changed();
  }
</script>

<div class="pane-head">
  <span>{t('find.title')}</span>
  <button class="pane-close" onclick={() => (session.pane.left = null)} aria-label={t('find.close')}><RibbonIcon name="close" size={14} /></button>
</div>
<div class="pane-search-row">
  <label class="pane-search">
    <RibbonIcon name="search" size={14} />
    <input placeholder={t('find.find')} bind:value={session.search} aria-label={t('find.find')} onkeydown={(e) => { if (e.key === 'Enter') next(session, matches, e.shiftKey ? -1 : 1); }} />
  </label>
  <button class="pane-icon" onclick={() => (showOptions = !showOptions)} aria-label={t('home.find.options')} aria-expanded={showOptions} title={t('home.find.options')}><RibbonIcon name="findOptions" size={14} /></button>
</div>
{#if showOptions}
  <div class="pane-options">
    <label class="field"><input type="checkbox" bind:checked={home.find.matchCase} />{t('home.find.matchCase')}</label>
    <label class="field"><input type="checkbox" bind:checked={home.find.wholeWords} disabled={home.find.wildcards} />{t('home.find.wholeWords')}</label>
    <label class="field"><input type="checkbox" bind:checked={home.find.wildcards} />{t('home.find.wildcards')}</label>
    <button class="pane-link" onclick={() => session.openDialog('home.find')}>{t('home.find.advanced')}</button>
  </div>
{/if}
<input class="pane-input" placeholder={t('find.replaceWith')} bind:value={home.find.replace} aria-label={t('find.replaceWith')} />
<div class="pane-buttons">
  <button class="pane-button" onclick={replaceOne} disabled={matches.length === 0}>{t('home.find.replace')}</button>
  <button class="pane-button" onclick={replaceAll} disabled={matches.length === 0}>{t('find.replaceAll')}</button>
</div>
<div class="pane-tabs" role="tablist">
  <button role="tab" aria-selected={view === 'headings'} class:active={view === 'headings'} onclick={() => (view = 'headings')}>{t('home.find.headings')}</button>
  <button role="tab" aria-selected={view === 'results'} class:active={view === 'results'} onclick={() => (view = 'results')}>{t('home.find.results')}</button>
</div>
{#if view === 'results'}
  <div class="pane-results-head">
    <span aria-live="polite">{session.search || Object.keys(home.find.format).length > 0 ? t('home.find.count').replace('{n}', String(matches.length)) : ''}</span>
    <span>
      <button class="pane-icon" onclick={() => next(session, matches, -1)} disabled={matches.length === 0} aria-label={t('home.find.previous')} title={t('home.find.previous')}><RibbonIcon name="findPrevious" size={14} /></button>
      <button class="pane-icon" onclick={() => next(session, matches, 1)} disabled={matches.length === 0} aria-label={t('home.find.next')} title={t('home.find.next')}><RibbonIcon name="findNext" size={14} /></button>
    </span>
  </div>
  <ul class="pane-list">
    {#each matches as m (`${key(m.at)}:${m.start}`)}
      {@const c = context(m)}
      <li><button onclick={() => goTo(session, m)}>{c.before}<mark>{c.hit}</mark>{c.after}</button></li>
    {/each}
  </ul>
{:else}
  <ul class="pane-list">
    {#each headings as h (key(h.at))}
      <li><button style="padding-left:{8 + h.level * 12}px" onclick={() => goToHeading(h.at)}>{h.text || '—'}</button></li>
    {/each}
  </ul>
{/if}
<p class="pane-note" aria-live="polite">{result.error ?? findStatus}</p>
