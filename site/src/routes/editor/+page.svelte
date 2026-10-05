<script lang="ts">
  import { onMount } from 'svelte';
  import { createEditor, openEditor, editorFor } from '@office-kit/docx-editor';
  import {
    createDocx,
    appendHeading,
    appendParagraph,
    addTable,
    toUint8Array,
    PAGE_SIZE_A4,
    MARGINS_NORMAL,
    setPageSize,
    setPageMargins,
    ensureHeadingStyles,
  } from '@office-kit/docx';
  import EditorCanvas from '$lib/editor/EditorCanvas.svelte';
  import RibbonIcon from '$lib/editor/RibbonIcon.svelte';
  import StatusBar from '$lib/editor/StatusBar.svelte';
  import InsertDialogs from '$lib/editor/insert/InsertDialogs.svelte';
  import HomeDialogs from '$lib/editor/ribbon/tabs/home/HomeDialogs.svelte';
  import Ruler from '$lib/editor/Ruler.svelte';
  import { TABS } from '$lib/editor/ribbon/tabs';
  import { PANES } from '$lib/editor/panes/registry';
  import { EditorSession, setSession, type PaneSide } from '$lib/editor/session.svelte';
  import { t } from '$lib/editor/i18n/index.svelte';
  import '$lib/editor/ribbon.css';
  import '$lib/editor/review-view.css';

  const session = new EditorSession();
  setSession(session);

  let tab = $state('home');
  let appEl = $state<HTMLDivElement | null>(null);
  // Distance from the top of the document to the editor: the site header sits
  // above it, and the editor fills the rest of the window like Word's.
  let appTop = $state(0);

  const visibleTabs = $derived(TABS.filter((x) => !x.when || (session.tick >= 0 && x.when(session))));
  // A contextual tab that stops applying (the caret left the table) falls back
  // to Home, as Word does.
  const current = $derived(visibleTabs.find((x) => x.id === tab) ?? visibleTabs[0] ?? TABS[0]);
  // Opening a header or footer brings up its contextual tab, as in Word.
  $effect(() => {
    if (session.headerFooter) tab = 'headerFooter';
  });
  // Selecting a picture, chart, SmartArt or shape brings up its format tab, as in Word.
  const OBJECT_TAB: Readonly<Record<string, string>> = {
    picture: 'pictureFormat',
    chart: 'chartDesign',
    smartArt: 'smartArtDesign',
    shape: 'shapeFormat',
    textBox: 'shapeFormat',
    ink: 'shapeFormat',
  };
  $effect(() => {
    const kind = session.selectedObject?.kind;
    const objectTab = kind && OBJECT_TAB[kind];
    if (objectTab) tab = objectTab;
  });
  // Switching to the Outline view brings up its Outlining tab, as in Word.
  $effect(() => {
    if (session.viewMode === 'outline') tab = 'outlining';
  });

  onMount(() => {
    const measure = (): void => {
      if (appEl) appTop = appEl.getBoundingClientRect().top + window.scrollY;
    };
    measure();
    window.addEventListener('resize', measure);
    session.load(editorFor(sampleDoc()), t('doc.newName'));
    // Dev-only handle so the editor can be driven/inspected from the console
    // (and by the e2e verification). Stripped from production builds.
    if (import.meta.env.DEV) Object.assign(window, { wkEditorModel: () => session.model, wkEditorSession: session });
    return () => window.removeEventListener('resize', measure);
  });

  function sampleDoc() {
    const doc = createDocx({ paragraphs: [] });
    setPageSize(doc, PAGE_SIZE_A4);
    setPageMargins(doc, MARGINS_NORMAL);
    ensureHeadingStyles(doc);
    appendHeading(doc, 'Welcome to the word-kit editor', 1);
    appendParagraph(
      doc,
      'This is a Word-like editor. Every edit routes through @office-kit/docx, so what you save is a real .docx. Type here, use the ribbon above, then Save.',
    );
    appendHeading(doc, 'Try it', 2);
    appendParagraph(doc, 'Select text and click Bold, or change alignment and color.');
    addTable(doc, [
      ['Feature', 'Status'],
      ['Text formatting', 'Editable'],
      ['Tables', 'Editable'],
    ]);
    return doc;
  }

  function onGlobalKey(e: KeyboardEvent): void {
    const mod = e.ctrlKey || e.metaKey;
    if (e.key === 'Escape') session.openMenu = null;
    if (mod && (e.key.toLowerCase() === 'h' || e.key.toLowerCase() === 'f')) {
      e.preventDefault();
      session.pane.left = 'navigation';
    }
  }

  /** Close an open drop-down when the pointer goes down outside it. */
  function onGlobalPointer(e: PointerEvent): void {
    if (session.openMenu && !(e.target as Element).closest('.split')) session.openMenu = null;
  }

  /**
   * Ribbon buttons must not take focus from the page: Word keeps the text
   * selection while you click formatting buttons, and a focused button would
   * collapse the canvas selection before the command runs.
   */
  function keepSelection(e: MouseEvent): void {
    if ((e.target as Element).closest('button')) e.preventDefault();
  }

  async function openFile(e: Event): Promise<void> {
    const input = e.target as HTMLInputElement;
    const file = input.files?.[0];
    if (!file) return;
    const bytes = new Uint8Array(await file.arrayBuffer());
    // Clear the input so re-opening the same file fires `change` again.
    input.value = '';
    try {
      session.load(openEditor(bytes), file.name);
    } catch (err) {
      session.status = `Failed to open: ${(err as Error).message}`;
    }
  }

  function download(): void {
    if (!session.model) return;
    const bytes = toUint8Array(session.model.doc);
    // Copy into a plain ArrayBuffer so the Blob part type is unambiguous
    // (Uint8Array<ArrayBufferLike> is not assignable to BlobPart under strict DOM libs).
    const buffer = bytes.slice().buffer;
    const blob = new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = session.fileName.endsWith('.docx') ? session.fileName : `${session.fileName}.docx`;
    a.click();
    URL.revokeObjectURL(url);
    session.status = `Downloaded ${a.download}.`;
  }
</script>

<svelte:head><title>word-kit editor</title></svelte:head>
<svelte:window onkeydown={onGlobalKey} onpointerdown={onGlobalPointer} />

{#snippet pane(side: PaneSide)}
  {@const id = session.pane[side]}
  {@const Pane = id ? PANES[id] : undefined}
  {#if Pane}<aside class="pane {side}"><Pane /></aside>{/if}
{/snippet}

<div class="wk-app" class:marks={session.showMarks} class:field-codes={session.showFieldCodes} {...session.displayAttrs} bind:this={appEl} style="--app-top: {appTop}px">
  <!-- Title bar: Quick Access Toolbar, document name, search (Word for Mac). -->
  <header class="titlebar" onmousedown={keepSelection} role="toolbar" tabindex="-1">
    <div class="qat">
      <button class="qb" onclick={() => session.load(createEditor(), t('doc.newName'))} title={t('action.new')} aria-label={t('action.new')}><RibbonIcon name="newDoc" size={18} /></button>
      <label class="qb" title={t('action.open')} aria-label={t('action.open')}><RibbonIcon name="open" size={18} /><input type="file" accept=".docx" onchange={openFile} hidden /></label>
      <button class="qb" onclick={download} title={t('action.save')} aria-label={t('action.save')}><RibbonIcon name="save" size={18} /></button>
      <button class="qb" onclick={() => session.undo()} disabled={session.tick < 0 || !session.model?.canUndo()} title={`${t('action.undo')} (⌘Z)`} aria-label={t('action.undo')}><RibbonIcon name="undo" size={18} /></button>
      <button class="qb" onclick={() => session.redo()} disabled={session.tick < 0 || !session.model?.canRedo()} title={`${t('action.redo')} (⌘Y)`} aria-label={t('action.redo')}><RibbonIcon name="redo" size={18} /></button>
    </div>
    <input class="docname" bind:value={session.fileName} aria-label="File name" spellcheck="false" />
    <label class="search">
      <RibbonIcon name="search" size={16} />
      <input
        placeholder={t('search.placeholder')}
        bind:value={session.search}
        onfocus={() => (session.pane.left = 'navigation')}
        aria-label={t('search.placeholder')}
      />
    </label>
  </header>

  <!-- Ribbon tabs, with Comments / Download on the right like Word's Comments / Share. -->
  <div class="tabs" onmousedown={keepSelection} role="toolbar" tabindex="-1">
    <div class="tablist" role="tablist">
      {#each visibleTabs as x (x.id)}
        <button class="tab" class:contextual={!!x.when} class:active={current.id === x.id} role="tab" aria-selected={current.id === x.id} onclick={() => (tab = x.id)}>{t(x.label)}</button>
      {/each}
    </div>
    <div class="tab-actions">
      <button class="pill" onclick={() => session.togglePane('right', 'comments')} class:on={session.pane.right === 'comments'}><RibbonIcon name="comment" size={16} />{t('action.comments')}</button>
      <button class="pill" onclick={download}><RibbonIcon name="share" size={16} />{t('action.download')}</button>
    </div>
  </div>

  <!-- Ribbon body. Word for Mac shows no group captions; groups are split by rules. -->
  <div class="ribbon" onmousedown={keepSelection} role="toolbar" tabindex="-1">
    <current.component />
  </div>

  <!-- Editing surface -->
  <div class="workspace">
    {@render pane('left')}
    <div class="surface">
      {#if session.showRuler && session.model}<Ruler />{/if}
      {#if session.model}
        <EditorCanvas model={session.model} />
      {:else}
        <p class="loading">Loading editor…</p>
      {/if}
    </div>
    {@render pane('right')}
  </div>

  <StatusBar />
  <InsertDialogs />
  <HomeDialogs />
</div>

<style>
  /* Word for Mac (light appearance) chrome. */
  .wk-app {
    /* Measured from Word for Mac 16 (light appearance). */
    --chrome: #e8e8e8;
    --chrome-line: #cccccc;
    --control-line: #c8c8c8;
    --hover: #dadada;
    --pressed: #c8c8c8;
    --text: #262626;
    --icon: #6b6b6b;
    --muted: #6e6e6e;
    --accent: #365695;
    --canvas-bg: #ececec;
    --status-bg: #f5f5f5;
    --status-line: #c1c1c1;
    display: flex;
    flex-direction: column;
    height: calc(100dvh - var(--app-top));
    background: var(--chrome);
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', system-ui, sans-serif;
    font-size: 13px;
    color: var(--text);
    /* The site root is `color-scheme: dark`, which gives native selects and
       inputs light default text; on this light UI their values were unreadable. */
    color-scheme: light;
  }


  .titlebar {
    display: grid;
    grid-template-columns: 1fr auto 1fr;
    align-items: center;
    gap: 12px;
    height: 38px;
    padding: 0 12px;
    outline: none;
  }
  .qat { display: flex; align-items: center; gap: 2px; }
  .qb {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 28px;
    height: 26px;
    border: none;
    border-radius: 5px;
    background: none;
    cursor: pointer;
  }
  .qb:hover:not(:disabled) { background: var(--hover); }
  .qb:disabled { opacity: 0.35; cursor: default; }
  .docname {
    justify-self: center;
    width: 26ch;
    border: 1px solid transparent;
    border-radius: 4px;
    background: none;
    padding: 2px 6px;
    font: inherit;
    font-weight: 600;
    text-align: center;
    color: var(--text);
  }
  .docname:hover { border-color: var(--chrome-line); }
  .docname:focus { background: #fff; border-color: var(--accent); outline: none; }
  /* Word draws its search field without a box until it is focused. */
  .search {
    justify-self: end;
    display: flex;
    align-items: center;
    gap: 6px;
    width: min(240px, 100%);
    height: 26px;
    padding: 0 8px;
    border: 1px solid transparent;
    border-radius: 6px;
    color: var(--muted);
  }
  .search:focus-within { background: #fff; border-color: var(--control-line); }
  .search input::placeholder { color: var(--muted); }
  .search input { flex: 1; min-width: 0; border: none; outline: none; background: none; font: inherit; color: var(--text); }

  .tabs { outline: none; display: flex; align-items: flex-end; justify-content: space-between; padding: 0 8px; }
  .tablist { display: flex; gap: 4px; overflow-x: auto; }
  .tab {
    position: relative;
    padding: 6px 10px 8px;
    border: none;
    background: none;
    cursor: pointer;
    font-size: 14px;
    color: var(--text);
    white-space: nowrap;
  }
  .tab:hover { color: #000; }
  .tab.active { font-weight: 600; }
  /* Contextual tabs (Table Design, Picture Format …) are drawn in the accent color. */
  .tab.contextual { color: #5b7fbd; }
  .tab.active::after {
    content: '';
    position: absolute;
    left: 10px;
    right: 10px;
    bottom: 2px;
    height: 3px;
    border-radius: 2px;
    background: var(--accent);
  }
  .tab-actions { display: flex; gap: 8px; padding-bottom: 4px; }
  .pill {
    display: inline-flex;
    align-items: center;
    gap: 6px;
    height: 28px;
    padding: 0 12px;
    border: 1px solid var(--control-line);
    border-radius: 6px;
    background: #fff;
    cursor: pointer;
    font-size: 14px;
  }
  .pill:hover:not(:disabled) { background: var(--hover); }
  .pill.on { background: var(--pressed); }
  .pill:disabled { opacity: 0.45; cursor: default; }

  .workspace { flex: 1; display: flex; min-height: 0; background: var(--canvas-bg); }
  .surface { flex: 1; overflow: auto; padding: 24px 24px 48px; }
  .loading { text-align: center; color: var(--muted); }

  /* Show/Hide ¶: a pilcrow at the end of every paragraph, as Word draws it. */
  .wk-app.marks :global(.wk-p)::after { content: '¶'; color: #7a7a7a; font-weight: normal; font-style: normal; }
</style>
