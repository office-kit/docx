<script lang="ts">
  /**
   * Word's status bar: page, word count, proofing language, accessibility and
   * Track Changes on the left; Focus, the view buttons and zoom on the right.
   * It also hosts the Review / View dialogs, which the ribbon and the status
   * bar both open, outside the bar so Focus (which hides the bar) keeps them.
   */
  import { commands, runAtPath } from '@office-kit/docx-editor';
  import { checkAccessibility, getRunLanguage, writeProtection } from '@office-kit/docx';
  import RibbonIcon, { type IconName } from './RibbonIcon.svelte';
  import { getSession, MIN_ZOOM, MAX_ZOOM, type ViewMode } from './session.svelte';
  import { locale, t, type MessageKey } from './i18n/index.svelte';
  import { DEFAULT_LANGUAGE, languageName } from './languages';
  import CompareDialog from './dialogs/CompareDialog.svelte';
  import LanguageDialog from './dialogs/LanguageDialog.svelte';
  import PropertiesDialog from './dialogs/PropertiesDialog.svelte';
  import TrackChangesOptionsDialog from './dialogs/TrackChangesOptionsDialog.svelte';
  import WordCountDialog from './dialogs/WordCountDialog.svelte';
  import ZoomDialog from './dialogs/ZoomDialog.svelte';

  const session = getSession();
  const prefs = session.prefs;

  const VIEWS: readonly [ViewMode, IconName, MessageKey][] = [
    ['print', 'viewPrintLayout', 'view.printLayout'],
    ['web', 'viewWebLayout', 'view.webLayout'],
    ['outline', 'viewOutline', 'view.outline'],
    ['draft', 'viewDraft', 'view.draft'],
  ];

  const language = $derived.by(() => {
    const model = session.model;
    if (session.tick < 0 || !model?.selection) return DEFAULT_LANGUAGE;
    const run = runAtPath(model.doc, model.selection.focus);
    return (run && getRunLanguage(run).latin) ?? DEFAULT_LANGUAGE;
  });
  const accessible = $derived(
    !prefs.keepAccessibilityRunning || session.tick < 0 || !session.model
      ? undefined
      : checkAccessibility(session.model.doc).every((i) => i.severity === 'tip'),
  );
  const pageText = $derived(
    t('view.pageOf').replace('{n}', String(session.currentPage)).replace('{m}', String(session.pageCount)),
  );

  // Always Open Read-Only: Word asks on open; here the request is shown once per document.
  let notifiedFor: object | null = null;
  $effect(() => {
    const model = session.model;
    if (!model || model === notifiedFor) return;
    notifiedFor = model;
    if (writeProtection(model.doc)?.recommended) session.status = t('review.readOnlyNotice');
  });

  function onKey(e: KeyboardEvent): void {
    if (e.key === 'Escape' && prefs.focus) prefs.focus = false;
  }
</script>

<svelte:window onkeydown={onKey} />

<footer class="statusbar">
  <span>{pageText}</span>
  <button class="sb" onclick={() => session.openDialog('wordCount')} title={`${session.charCount} ${t('status.chars')}`}>{session.wordCount.toLocaleString()} {t('status.words')}</button>
  <button class="sb" onclick={() => (session.spellcheck = !session.spellcheck)} title={t('review.spelling')} aria-label={t('review.spelling')} aria-pressed={session.spellcheck}>
    <RibbonIcon name="statusProofing" size={16} />
  </button>
  <button class="sb" onclick={() => session.openDialog('language')} title={t('review.setLanguage')} disabled={!session.model?.selection}>{languageName(language, locale())}</button>
  {#if accessible !== undefined}
    <button class="sb" onclick={() => session.togglePane('right', 'accessibility')}>
      <RibbonIcon name="reviewAccessibility" size={16} />{accessible ? t('review.a11yGood') : t('review.a11yInvestigate')}
    </button>
  {/if}
  {#if session.active(commands.toggleTrackChangesCommand)}
    <button class="sb on" onclick={() => session.apply(commands.toggleTrackChangesCommand, undefined)} disabled={!session.enabled(commands.toggleTrackChangesCommand)} title={t('review.trackChanges')}>
      <RibbonIcon name="statusTracking" size={16} />{t('review.trackChanges')}
    </button>
  {/if}
  {#if session.status}<span class="status-msg" role="status">{session.status}</span>{/if}
  <div class="spacer"></div>
  <button class="sb" onclick={() => (prefs.focus = true)} title={t('view.focus')}><RibbonIcon name="viewFocus" size={16} />{t('view.focus')}</button>
  <span class="views">
    {#each VIEWS as [mode, icon, key] (mode)}
      <button class="sb" class:on={session.viewMode === mode} onclick={() => session.setView(mode)} title={t(key)} aria-label={t(key)} aria-pressed={session.viewMode === mode}><RibbonIcon name={icon} size={16} /></button>
    {/each}
  </span>
  <div class="zoom">
    <button onclick={() => session.setZoom(session.zoom - 0.1)} aria-label={t('view.zoomOut')}>−</button>
    <input
      type="range"
      min={MIN_ZOOM * 100}
      max={MAX_ZOOM * 100}
      step="10"
      value={Math.round(session.zoom * 100)}
      oninput={(e) => session.setZoom(Number((e.currentTarget as HTMLInputElement).value) / 100)}
      aria-label={t('view.zoom')}
    />
    <button onclick={() => session.setZoom(session.zoom + 0.1)} aria-label={t('view.zoomIn')}>+</button>
    <button class="zoom-val" onclick={() => session.openDialog('zoom')} title={t('view.zoom')}>{Math.round(session.zoom * 100)}%</button>
  </div>
</footer>

{#if prefs.focus}
  <button class="focus-exit" onclick={() => (prefs.focus = false)}>{t('view.exitFocus')}</button>
{/if}

<WordCountDialog />
<LanguageDialog />
<TrackChangesOptionsDialog />
<CompareDialog />
<ZoomDialog />
<PropertiesDialog />

<style>
  .statusbar {
    display: flex;
    align-items: center;
    gap: 10px;
    height: 26px;
    padding: 0 14px;
    border-top: 1px solid var(--status-line);
    background: var(--status-bg);
    font-size: 13px;
    color: var(--text);
    white-space: nowrap;
  }
  .sb {
    display: inline-flex;
    align-items: center;
    gap: 4px;
    height: 22px;
    padding: 0 4px;
    border: none;
    border-radius: 3px;
    background: none;
    cursor: pointer;
  }
  .sb:hover:not(:disabled) { background: var(--hover); }
  .sb.on { background: var(--pressed); }
  .views { display: inline-flex; gap: 2px; }
  .status-msg { color: #a4262c; overflow: hidden; text-overflow: ellipsis; }
  .spacer { flex: 1; }
  .zoom { display: flex; align-items: center; gap: 6px; }
  .zoom button { min-width: 18px; height: 18px; border: none; background: none; cursor: pointer; font-size: 14px; line-height: 1; border-radius: 3px; color: inherit; }
  .zoom button:hover { background: var(--hover); }
  .zoom input[type='range'] { width: 130px; accent-color: #8a8a8a; }
  .zoom-val { min-width: 40px; text-align: right; font-size: 13px !important; }
  .focus-exit {
    position: fixed;
    top: 12px;
    right: 16px;
    z-index: 50;
    padding: 4px 12px;
    border: 1px solid #666;
    border-radius: 6px;
    background: #2b2b2b;
    color: #fff;
    cursor: pointer;
  }
</style>
