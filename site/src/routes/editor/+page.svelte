<script lang="ts">
  import { onMount } from 'svelte';
  import {
    createEditor,
    openEditor,
    runCommand,
    commands,
    runAtPath,
    paragraphAt,
    orderSelection,
    paragraphsInRange,
    runsInRange,
    type EditorModel,
    type Command,
  } from '@office-kit/docx-editor';
  import {
    createDocx,
    appendHeading,
    appendParagraph,
    addTable,
    toUint8Array,
    text as documentText,
    getRunFormat,
    type RunFormatting,
    getParagraphStyle,
    PAGE_SIZE_A4,
    PAGE_SIZE_LETTER,
    MARGINS_NORMAL,
    setPageSize,
    setPageMargins,
    ensureHeadingStyles,
    HIGHLIGHT_COLORS,
  } from '@office-kit/docx';
  import { editorFor } from '@office-kit/docx-editor';
  import EditorCanvas from '$lib/editor/EditorCanvas.svelte';
  import RawXmlInspector from '$lib/editor/RawXmlInspector.svelte';
  import { t, locale, setLocale, LOCALES, type LocaleId } from '$lib/editor/i18n.svelte';

  type TabId = 'home' | 'insert' | 'layout' | 'references' | 'review';

  let model = $state<EditorModel | null>(null);
  let version = $state(0);
  let tick = $state(0);
  let tab = $state<TabId>('home');
  let fileName = $state('Untitled.docx');
  let status = $state('');
  let showXml = $state(false);
  let zoom = $state(1);
  let wordCount = $state(0);
  let charCount = $state(0);

  // Find & replace.
  let showFind = $state(false);
  let findQuery = $state('');
  let replaceValue = $state('');
  let findStatus = $state('');

  /**
   * What a ribbon control shows for the current selection. `inherited` means
   * the text has no direct value: it comes from styles / document defaults,
   * which the public API does not resolve, so the UI says so instead of
   * guessing a font or size. Nothing here is ever written back to the
   * document; only an explicit user change runs a command.
   */
  type FieldState<T> =
    | { kind: 'none' }
    | { kind: 'mixed' }
    | { kind: 'inherited' }
    | { kind: 'value'; value: T };

  const NONE = { kind: 'none' } as const;
  let fontState = $state<FieldState<string>>(NONE);
  let sizeState = $state<FieldState<number>>(NONE);
  let colorState = $state<FieldState<string>>(NONE);
  let highlightState = $state<FieldState<string>>(NONE);
  let styleState = $state<FieldState<string>>(NONE);
  // The color the A button applies (like Word's split button).
  // `<input type=color>` only accepts #rrggbb, so this is always valid hex;
  // the selection's own color is shown separately, in the label.
  let fontColorPen = $state('#111111');

  const FONT_CHOICES = ['Calibri', 'Arial', 'Times New Roman', 'Georgia', 'Courier New', 'Yu Gothic'];
  const STYLE_CHOICES = ['Heading1', 'Heading2', 'Heading3'] as const;
  const HEX6 = /^[0-9a-fA-F]{6}$/;
  // Value of the disabled "No selection" / "Mixed" placeholder option. It can
  // never be chosen, and is not a style id or font name.
  const STATE_OPTION = '\u2014state';

  onMount(() => {
    loadModel(editorFor(sampleDoc()));
    // Dev-only handle so the editor can be driven/inspected from the console
    // (and by the e2e verification). Stripped from production builds.
    if (import.meta.env.DEV)
      (window as unknown as { wkEditorModel?: () => EditorModel | null }).wkEditorModel = () => model;
  });

  function sampleDoc() {
    const doc = createDocx({ paragraphs: [] });
    setPageSize(doc, PAGE_SIZE_A4);
    setPageMargins(doc, MARGINS_NORMAL);
    ensureHeadingStyles(doc);
    appendHeading(doc, 'Welcome to the word-kit editor', 1);
    appendParagraph(
      doc,
      'This is a Word-like editor. Every edit routes through @office-kit/docx, so what you save is a real .docx. Type here, use the ribbon above, then Download.',
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

  /**
   * Run a command and force a structural re-render. Commands are atomic: a
   * failure (e.g. rejected input) leaves the document untouched, so it is
   * reported in the status bar rather than crashing the page.
   */
  function apply<P, R>(cmd: Command<P, R>, params: P): R | undefined {
    if (!model) return undefined;
    let result: R | undefined;
    try {
      result = runCommand(model, cmd, params);
    } catch (err) {
      status = `${cmd.label} failed: ${(err as Error).message}`;
      return undefined;
    } finally {
      version++;
      tick++;
      refreshCounts();
      syncToolbarFromCaret();
    }
    status = `${cmd.label} applied.`;
    return result;
  }

  /** Reset every piece of chrome derived from the previous document. */
  function loadModel(next: EditorModel): void {
    model = next;
    version++;
    tick++;
    refreshCounts();
    findStatus = '';
    syncToolbarFromCaret();
  }

  function active(cmd: Command<unknown>): boolean {
    // `tick` is read so the button state recomputes on selection/edit changes.
    return tick >= 0 && !!model && (cmd.isActive?.(model) ?? false);
  }

  function enabled(cmd: Command<unknown>): boolean {
    return tick >= 0 && !!model && (cmd.isEnabled?.(model) ?? true);
  }

  /** Enabled for the current selection (some commands have no isEnabled). */
  function canFormat(cmd: Command<unknown>): boolean {
    return enabled(cmd) && !!model?.selection;
  }

  function refreshCounts(): void {
    if (!model) return;
    const t = documentText(model.doc);
    charCount = t.length;
    const words = t.trim().match(/\S+/g);
    wordCount = words ? words.length : 0;
  }

  /** Collapse per-run (or per-paragraph) direct values into one control state. */
  function fieldOf<T>(values: ReadonlyArray<T | undefined>): FieldState<T> {
    if (values.length === 0) return NONE;
    const first = values[0];
    if (values.some((v) => v !== first)) return { kind: 'mixed' };
    return first === undefined ? { kind: 'inherited' } : { kind: 'value', value: first };
  }

  /**
   * Reflect the selection's direct run formatting and paragraph styles in the
   * ribbon. A caret reads the run it is in; a range reads every run it covers,
   * so a differently formatted run in the middle makes the field "mixed".
   */
  function syncToolbarFromCaret(): void {
    const sel = model?.selection;
    if (!model || !sel) {
      fontState = sizeState = colorState = highlightState = styleState = NONE;
      return;
    }
    const doc = model.doc;
    const ordered = orderSelection(sel);
    const runs = ordered.collapsed
      ? [runAtPath(doc, sel.focus)].filter((run) => run !== undefined)
      : runsInRange(doc, ordered);
    const found = runs.map((run) => getRunFormat(run));
    // A caret in a paragraph with no runs still has a selection: its text
    // would take inherited formatting, so report that rather than "none".
    const noDirectFormat: RunFormatting = {};
    const formats = found.length > 0 ? found : [noDirectFormat];
    fontState = fieldOf(formats.map((f) => f.font));
    sizeState = fieldOf(formats.map((f) => (f.fontSizeHalfPoints ? f.fontSizeHalfPoints / 2 : undefined)));
    // `auto` is a real OOXML color (let the consumer pick); show it as such.
    colorState = fieldOf(formats.map((f) => f.color));
    highlightState = fieldOf(formats.map((f) => f.highlight));
    if (colorState.kind === 'value' && HEX6.test(colorState.value)) fontColorPen = `#${colorState.value}`;
    // paragraphsInRange takes every cell of a table in range, so a caret uses
    // only its own paragraph.
    const caretPara = ordered.collapsed ? paragraphAt(doc, sel.focus) : undefined;
    const paras = ordered.collapsed ? (caretPara ? [caretPara] : []) : paragraphsInRange(doc, ordered);
    styleState = fieldOf(paras.map((p) => getParagraphStyle(p)));
  }

  function isListedHighlight(value: string): boolean {
    return HIGHLIGHT_COLORS.some((c) => c === value);
  }

  /** Short text for a non-value state (placeholder / label). */
  function stateText(state: FieldState<unknown>): string {
    if (state.kind === 'none') return t('state.none');
    if (state.kind === 'mixed') return t('state.mixed');
    return t('state.inherited');
  }

  /** Label for a color state; `named` maps special values (`auto`, `none`) to text. */
  function colorLabel(state: FieldState<string>, named: Record<string, string>): string {
    if (state.kind !== 'value') return stateText(state);
    const special = named[state.value];
    if (special !== undefined) return special;
    return HEX6.test(state.value) ? `#${state.value.toUpperCase()}` : state.value;
  }

  function onFontSizeChange(e: Event): void {
    const input = e.currentTarget as HTMLInputElement;
    const points = Number(input.value);
    // Word accepts 1–1638 pt in half-point steps; reject anything else without
    // touching the document, and put the control back to the selection's state.
    if (input.value === '' || !Number.isFinite(points) || points < 1 || points > 1638) {
      input.value = sizeState.kind === 'value' ? String(sizeState.value) : '';
      return;
    }
    apply(commands.setFontSizeCommand, { points });
  }

  function onCanvasSelection(): void {
    tick++;
    syncToolbarFromCaret();
  }

  /** Canvas-originated edit (typing / Enter / paste / undo): refresh chrome only. */
  function onCanvasEdit(): void {
    tick++;
    refreshCounts();
    syncToolbarFromCaret();
    status = t('status.editing');
  }

  function undo(): void {
    if (!model) return;
    model.undo();
    version++;
    tick++;
    refreshCounts();
    syncToolbarFromCaret();
    status = 'Undo.';
  }

  function redo(): void {
    if (!model) return;
    model.redo();
    version++;
    tick++;
    refreshCounts();
    syncToolbarFromCaret();
    status = 'Redo.';
  }

  function doReplaceAll(): void {
    if (!findQuery) return;
    const n = apply(commands.replaceAllCommand, { query: findQuery, replacement: replaceValue });
    if (n === undefined) return;
    findStatus = `Replaced ${n} occurrence${n === 1 ? '' : 's'}.`;
  }

  function setZoom(z: number): void {
    zoom = Math.max(0.5, Math.min(2, z));
  }

  function onGlobalKey(e: KeyboardEvent): void {
    const mod = e.ctrlKey || e.metaKey;
    if (mod && (e.key.toLowerCase() === 'h' || e.key.toLowerCase() === 'f')) {
      e.preventDefault();
      showFind = e.key.toLowerCase() === 'h' ? !showFind : true;
    }
  }

  async function openFile(e: Event): Promise<void> {
    const input = e.target as HTMLInputElement;
    const file = input.files?.[0];
    if (!file) return;
    const bytes = new Uint8Array(await file.arrayBuffer());
    // Clear the input so re-opening the same file fires `change` again.
    input.value = '';
    try {
      loadModel(openEditor(bytes));
      fileName = file.name;
      status = `Opened ${file.name}.`;
    } catch (err) {
      status = `Failed to open: ${(err as Error).message}`;
    }
  }

  function download(): void {
    if (!model) return;
    const bytes = toUint8Array(model.doc);
    // Copy into a plain ArrayBuffer so the Blob part type is unambiguous
    // (Uint8Array<ArrayBufferLike> is not assignable to BlobPart under strict DOM libs).
    const buffer = bytes.slice().buffer;
    const blob = new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName.endsWith('.docx') ? fileName : `${fileName}.docx`;
    a.click();
    URL.revokeObjectURL(url);
    status = `Downloaded ${a.download}.`;
  }

  function newDoc(): void {
    loadModel(createEditor());
    fileName = 'Untitled.docx';
    status = 'New document.';
  }

  function insertTable(): void {
    const spec = prompt('Table size (rows x cols)', '3x3');
    if (!spec) return;
    // Invalid sizes are rejected by the command itself and shown in the status bar.
    const [rows = Number.NaN, cols = Number.NaN] = spec
      .split(/[x×,]/)
      .map((n) => Number(n.trim()));
    apply(commands.insertTableCommand, { rows, cols });
  }

  function insertLink(): void {
    const url = prompt('Link URL', 'https://');
    if (!url) return;
    const text = prompt('Link text', url) ?? url;
    apply(commands.insertHyperlinkCommand, { url, text });
  }

  function insertHeading(level: number): void {
    const text = prompt(`Heading ${level} text`, 'Heading');
    if (text != null) apply(commands.insertHeadingCommand, { text, level });
  }

  function addComment(): void {
    const text = prompt('Comment');
    if (text) apply(commands.addCommentCommand, { author: 'You', initials: 'Y', text });
  }

  async function insertImage(e: Event): Promise<void> {
    const input = e.target as HTMLInputElement;
    const file = input.files?.[0];
    if (!file || !model) return;
    const bytes = new Uint8Array(await file.arrayBuffer());
    apply(commands.insertImageCommand, { bytes, options: { widthEmu: 2743200, heightEmu: 2057400 } });
    input.value = '';
  }

  const TAB_IDS: TabId[] = ['home', 'insert', 'layout', 'references', 'review'];
</script>

{#snippet alignIcon(d: string)}
  <!-- Drawn, not a font glyph: ⯇/⯈ are missing from common UI fonts. -->
  <svg viewBox="0 0 16 16" width="16" height="16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" aria-hidden="true"><path {d} /></svg>
{/snippet}

<svelte:head><title>word-kit editor</title></svelte:head>
<svelte:window onkeydown={onGlobalKey} />

<div class="app">
  <!-- Title bar -->
  <div class="titlebar">
    <span class="brand">word-kit</span>
    <input class="filename" bind:value={fileName} aria-label="File name" />
    <button onclick={undo} disabled={tick < 0 || !model?.canUndo()} title={t('action.undo') + ' (Ctrl+Z)'}>↶</button>
    <button onclick={redo} disabled={tick < 0 || !model?.canRedo()} title={t('action.redo') + ' (Ctrl+Y)'}>↷</button>
    <div class="spacer"></div>
    <select
      class="lang"
      value={locale()}
      onchange={(e) => setLocale((e.currentTarget as HTMLSelectElement).value as LocaleId)}
      title={t('language')}
      aria-label={t('language')}
    >
      {#each Object.entries(LOCALES) as [id, name] (id)}
        <option value={id}>{name}</option>
      {/each}
    </select>
    <button onclick={newDoc}>{t('action.new')}</button>
    <label class="btn-file">
      {t('action.open')}
      <input type="file" accept=".docx" onchange={openFile} hidden />
    </label>
    <button class:primary={showFind} onclick={() => (showFind = !showFind)} title="Ctrl+H">🔍 {t('action.find')}</button>
    <button class:primary={showXml} onclick={() => (showXml = !showXml)} title={t('xml.title')}>&lt;/&gt; {t('action.xml')}</button>
    <button class="primary" onclick={download}>{t('action.download')}</button>
  </div>

  {#if showFind}
    <div class="findbar">
      <input placeholder={t('find.find')} bind:value={findQuery} aria-label={t('find.find')} />
      <input placeholder={t('find.replaceWith')} bind:value={replaceValue} aria-label={t('find.replaceWith')} />
      <button onclick={doReplaceAll} disabled={!findQuery}>{t('find.replaceAll')}</button>
      <span class="find-status">{findStatus}</span>
      <button class="find-close" onclick={() => (showFind = false)} aria-label={t('find.close')}>✕</button>
    </div>
  {/if}

  <!-- Ribbon tabs -->
  <div class="tabs">
    {#each TAB_IDS as id (id)}
      <button class="tab" class:active={tab === id} onclick={() => (tab = id)}>{t(`tab.${id}`)}</button>
    {/each}
  </div>

  <!-- Ribbon body -->
  <div class="ribbon">
    {#if tab === 'home'}
      <div class="group">
        <div class="group-body">
          <!-- No bind:value: the control mirrors the selection; only a user
               choice runs a command (nothing is written for a mere caret move). -->
          <select
            value={fontState.kind === 'value' ? fontState.value : ''}
            disabled={!canFormat(commands.setFontCommand)}
            onchange={(e) => {
              const font = (e.currentTarget as HTMLSelectElement).value;
              if (font) apply(commands.setFontCommand, { font });
            }}
            aria-label={t('group.font')}
            title={fontState.kind === 'value' ? fontState.value : stateText(fontState)}
          >
            {#if fontState.kind !== 'value'}
              <option value="" disabled>{stateText(fontState)}</option>
            {:else if !FONT_CHOICES.includes(fontState.value)}
              <option value={fontState.value}>{fontState.value}</option>
            {/if}
            {#each FONT_CHOICES as font (font)}<option value={font}>{font}</option>{/each}
          </select>
          <input
            type="number"
            min="1"
            max="1638"
            step="0.5"
            value={sizeState.kind === 'value' ? sizeState.value : ''}
            placeholder={sizeState.kind === 'value' ? '' : stateText(sizeState)}
            disabled={!canFormat(commands.setFontSizeCommand)}
            onchange={onFontSizeChange}
            aria-label="Font size"
            title={sizeState.kind === 'value' ? `${sizeState.value} pt` : stateText(sizeState)}
            class="num"
          />
        </div>
        <div class="group-label">{t('group.font')}</div>
      </div>

      <div class="group">
        <div class="group-body">
          <button class:on={active(commands.toggleBoldCommand)} disabled={!enabled(commands.toggleBoldCommand)} onclick={() => apply(commands.toggleBoldCommand, undefined)}><b>B</b></button>
          <button class:on={active(commands.toggleItalicCommand)} disabled={!enabled(commands.toggleItalicCommand)} onclick={() => apply(commands.toggleItalicCommand, undefined)}><i>I</i></button>
          <button class:on={active(commands.toggleUnderlineCommand)} disabled={!enabled(commands.toggleUnderlineCommand)} onclick={() => apply(commands.toggleUnderlineCommand, undefined)}><u>U</u></button>
          <button class:on={active(commands.toggleStrikeCommand)} onclick={() => apply(commands.toggleStrikeCommand, undefined)}><s>S</s></button>
          <!-- The swatch is the color the button applies (always valid #rrggbb);
               the selection's own color — auto / inherited / mixed included —
               is in the label, so those meanings are never lost to a fake hex. -->
          <label class="color" title={`Font color: ${colorLabel(colorState, { auto: t('color.auto') })}`}>A<input
              type="color"
              bind:value={fontColorPen}
              onchange={() => apply(commands.setColorCommand, { color: fontColorPen })}
              aria-label={`Font color (${colorLabel(colorState, { auto: t('color.auto') })})`}
            /></label>
          <!-- w:highlight takes only ST_HighlightColor names, so this is a named
               list, not a color picker. A value outside the list (from another
               tool) is shown but cannot be chosen, so it is never re-written. -->
          <select
            value={highlightState.kind === 'value' ? highlightState.value : ''}
            disabled={!canFormat(commands.setHighlightCommand)}
            onchange={(e) => {
              const picked = (e.currentTarget as HTMLSelectElement).value;
              const color = HIGHLIGHT_COLORS.find((c) => c === picked);
              if (color) apply(commands.setHighlightCommand, { color });
            }}
            aria-label={`Highlight (${colorLabel(highlightState, { none: t('highlight.none') })})`}
            title={`Highlight: ${colorLabel(highlightState, { none: t('highlight.none') })}`}
          >
            {#if highlightState.kind !== 'value'}
              <option value="" disabled>{stateText(highlightState)}</option>
            {:else if !isListedHighlight(highlightState.value)}
              <option value={highlightState.value} disabled>{highlightState.value}</option>
            {/if}
            {#each HIGHLIGHT_COLORS as color (color)}
              <option value={color}>{color === 'none' ? t('highlight.none') : color}</option>
            {/each}
          </select>
          <button onclick={() => apply(commands.clearFormattingCommand, undefined)} title="Clear formatting">⌫</button>
        </div>
        <div class="group-label">{t('group.text')}</div>
      </div>

      <div class="group">
        <div class="group-body">
          <button class:on={active(commands.alignLeftCommand)} onclick={() => apply(commands.alignLeftCommand, undefined)} title="Align left" aria-label="Align left">{@render alignIcon('M2 3h12M2 6.5h8M2 10h12M2 13.5h8')}</button>
          <button class:on={active(commands.alignCenterCommand)} onclick={() => apply(commands.alignCenterCommand, undefined)} title="Center" aria-label="Center">{@render alignIcon('M2 3h12M4 6.5h8M2 10h12M4 13.5h8')}</button>
          <button class:on={active(commands.alignRightCommand)} onclick={() => apply(commands.alignRightCommand, undefined)} title="Align right" aria-label="Align right">{@render alignIcon('M2 3h12M6 6.5h8M2 10h12M6 13.5h8')}</button>
          <button class:on={active(commands.alignJustifyCommand)} onclick={() => apply(commands.alignJustifyCommand, undefined)} title="Justify" aria-label="Justify">{@render alignIcon('M2 3h12M2 6.5h12M2 10h12M2 13.5h12')}</button>
          <button onclick={() => apply(commands.applyListCommand, { kind: 'bullet' })} title="Bullet list">• —</button>
          <button onclick={() => apply(commands.applyListCommand, { kind: 'numbered' })} title="Numbered list">1. —</button>
          <button onclick={() => apply(commands.setIndentCommand, { left: 720 })} title="Indent">⇥</button>
        </div>
        <div class="group-label">{t('group.paragraph')}</div>
      </div>

      <div class="group">
        <div class="group-body">
          <!-- "inherited" = no pStyle, i.e. the default (Normal) paragraph style;
               choosing Normal removes pStyle, as before. -->
          <select
            value={styleState.kind === 'value' ? styleState.value : styleState.kind === 'inherited' ? '' : STATE_OPTION}
            disabled={!canFormat(commands.setParagraphStyleCommand)}
            onchange={(e) => {
              const v = (e.currentTarget as HTMLSelectElement).value;
              apply(commands.setParagraphStyleCommand, { styleId: v || undefined });
            }}
            aria-label="Paragraph style"
            title={styleState.kind === 'value' ? styleState.value : stateText(styleState)}
          >
            {#if styleState.kind === 'none' || styleState.kind === 'mixed'}
              <option value={STATE_OPTION} disabled>{stateText(styleState)}</option>
            {:else if styleState.kind === 'value' && !(STYLE_CHOICES as readonly string[]).includes(styleState.value)}
              <option value={styleState.value}>{styleState.value}</option>
            {/if}
            <option value="">{t('style.normal')}</option>
            <option value="Heading1">{t('style.heading1')}</option>
            <option value="Heading2">{t('style.heading2')}</option>
            <option value="Heading3">{t('style.heading3')}</option>
          </select>
        </div>
        <div class="group-label">{t('group.styles')}</div>
      </div>
    {:else if tab === 'insert'}
      <div class="group">
        <div class="group-body">
          <button onclick={() => apply(commands.insertParagraphCommand, { text: '' })}>¶ {t('ins.paragraph')}</button>
          <button onclick={insertTable}>▦ {t('ins.table')}</button>
          <button onclick={() => apply(commands.insertPageBreakCommand, undefined)}>⤓ {t('ins.pageBreak')}</button>
          <button onclick={insertLink}>🔗 {t('ins.link')}</button>
          <label class="btn-inline">
            🖼 {t('ins.picture')}
            <input type="file" accept="image/*" onchange={insertImage} hidden />
          </label>
        </div>
        <div class="group-label">{t('group.insert')}</div>
      </div>
      <div class="group">
        <div class="group-body">
          <button onclick={() => insertHeading(1)}>H1</button>
          <button onclick={() => insertHeading(2)}>H2</button>
          <button onclick={() => insertHeading(3)}>H3</button>
        </div>
        <div class="group-label">{t('group.headings')}</div>
      </div>
    {:else if tab === 'layout'}
      <div class="group">
        <div class="group-body">
          <button onclick={() => apply(commands.setPageSizeCommand, { size: PAGE_SIZE_A4 })}>A4</button>
          <button onclick={() => apply(commands.setPageSizeCommand, { size: PAGE_SIZE_LETTER })}>Letter</button>
          <button onclick={() => apply(commands.setOrientationCommand, { orientation: 'portrait' })}>Portrait</button>
          <button onclick={() => apply(commands.setOrientationCommand, { orientation: 'landscape' })}>Landscape</button>
        </div>
        <div class="group-label">{t('group.pageSetup')}</div>
      </div>
      <div class="group">
        <div class="group-body">
          <button onclick={() => apply(commands.insertSectionBreakCommand, { type: 'nextPage' })}>{t('layout.sectionBreak')}</button>
          <button onclick={() => apply(commands.addPageNumberFooterCommand, {})}>{t('layout.pageNumbers')}</button>
        </div>
        <div class="group-label">{t('group.sections')}</div>
      </div>
    {:else if tab === 'references'}
      <div class="group">
        <div class="group-body">
          <button onclick={() => apply(commands.insertTocCommand, {})}>{t('ref.toc')}</button>
          <button onclick={() => { const txt = prompt('Footnote text'); if (txt) apply(commands.addFootnoteCommand, { text: txt }); }}>{t('ref.footnote')}</button>
          <button onclick={() => { const txt = prompt('Endnote text'); if (txt) apply(commands.addEndnoteCommand, { text: txt }); }}>{t('ref.endnote')}</button>
          <button onclick={() => { const n = prompt('Bookmark name'); if (n) apply(commands.addBookmarkCommand, { name: n }); }}>{t('ref.bookmark')}</button>
        </div>
        <div class="group-label">{t('group.references')}</div>
      </div>
    {:else if tab === 'review'}
      <div class="group">
        <div class="group-body">
          <button onclick={addComment}>💬 {t('review.newComment')}</button>
          <button onclick={() => apply(commands.acceptAllRevisionsCommand, undefined)}>✓ {t('review.acceptAll')}</button>
          <button onclick={() => apply(commands.rejectAllRevisionsCommand, undefined)}>✗ {t('review.rejectAll')}</button>
        </div>
        <div class="group-label">{t('group.tracking')}</div>
      </div>
    {/if}
  </div>

  <!-- Editing surface -->
  <div class="workspace">
    <div class="surface">
      {#if model}
        <EditorCanvas
          {model}
          {version}
          {zoom}
          onselectionchange={onCanvasSelection}
          onerror={(message) => (status = message)}
          onedit={onCanvasEdit}
        />
      {:else}
        <p class="loading">Loading editor…</p>
      {/if}
    </div>
    {#if showXml && model}
      <aside class="xml-panel">
        <div class="xml-head">{t('xml.title')}</div>
        <RawXmlInspector {model} onchange={() => { version++; tick++; status = 'Raw XML edited.'; }} />
      </aside>
    {/if}
  </div>

  <div class="statusbar">
    <span>{status || t('status.ready')}</span>
    <div class="spacer"></div>
    <span class="count">{wordCount} {t('status.words')} · {charCount} {t('status.chars')}</span>
    <div class="zoom">
      <button onclick={() => setZoom(zoom - 0.1)} aria-label="Zoom out">−</button>
      <span class="zoom-val">{Math.round(zoom * 100)}%</span>
      <button onclick={() => setZoom(zoom + 0.1)} aria-label="Zoom in">+</button>
    </div>
  </div>
</div>

<style>
  .app {
    display: flex;
    flex-direction: column;
    height: 100vh;
    background: #f3f2f1;
    font-family: 'Segoe UI', system-ui, sans-serif;
    color: #201f1e;
  }
  .titlebar {
    display: flex;
    align-items: center;
    gap: 8px;
    padding: 6px 12px;
    background: #2b579a;
    color: #fff;
  }
  .brand { font-weight: 700; letter-spacing: 0.02em; }
  .filename {
    background: rgba(255, 255, 255, 0.15);
    border: 1px solid transparent;
    color: #fff;
    padding: 3px 8px;
    border-radius: 3px;
    width: 180px;
  }
  .filename:focus { background: #fff; color: #201f1e; outline: none; }
  .spacer { flex: 1; }
  .titlebar button, .btn-file {
    background: rgba(255, 255, 255, 0.15);
    color: #fff;
    border: none;
    padding: 5px 12px;
    border-radius: 4px;
    cursor: pointer;
    font-size: 13px;
  }
  .titlebar button.primary { background: #fff; color: #2b579a; font-weight: 600; }
  .btn-file { display: inline-block; }
  .lang {
    background: rgba(255, 255, 255, 0.15);
    color: #fff;
    border: none;
    padding: 4px 6px;
    border-radius: 4px;
    font-size: 13px;
    cursor: pointer;
  }
  .lang option { color: #201f1e; }
  .tabs {
    display: flex;
    gap: 2px;
    background: #fff;
    padding: 0 8px;
    border-bottom: 1px solid #e1dfdd;
  }
  .tab {
    background: none;
    border: none;
    padding: 8px 16px;
    cursor: pointer;
    font-size: 13px;
    color: #444;
    border-bottom: 2px solid transparent;
  }
  .tab.active { color: #2b579a; border-bottom-color: #2b579a; font-weight: 600; }
  .ribbon {
    display: flex;
    gap: 2px;
    align-items: stretch;
    background: #faf9f8;
    padding: 6px 8px;
    border-bottom: 1px solid #e1dfdd;
    min-height: 72px;
    overflow-x: auto;
  }
  .group {
    display: flex;
    flex-direction: column;
    padding: 0 8px;
    border-right: 1px solid #edebe9;
  }
  .group-body {
    display: flex;
    flex-wrap: wrap;
    gap: 3px;
    align-items: center;
    flex: 1;
    max-width: 260px;
  }
  .group-label { text-align: center; font-size: 11px; color: #888; padding-top: 4px; }
  .ribbon button {
    min-width: 30px;
    height: 30px;
    padding: 0 8px;
    border: 1px solid transparent;
    background: transparent;
    border-radius: 4px;
    cursor: pointer;
    font-size: 14px;
    color: #201f1e;
  }
  .ribbon button:hover { background: #edebe9; border-color: #d2d0ce; }
  .ribbon button.on { background: #cfe0f4; border-color: #2b579a; }
  .ribbon button:disabled { opacity: 0.4; cursor: default; }
  .ribbon select, .ribbon .num {
    height: 28px;
    border: 1px solid #d2d0ce;
    border-radius: 4px;
    padding: 0 6px;
    background: #fff;
  }
  .ribbon .num { width: 52px; }
  .color { display: inline-flex; align-items: center; gap: 2px; cursor: pointer; font-size: 13px; }
  .color input[type='color'] { width: 20px; height: 20px; border: none; background: none; padding: 0; cursor: pointer; }
  .workspace {
    flex: 1;
    display: flex;
    min-height: 0;
  }
  .surface {
    flex: 1;
    overflow-y: auto;
    padding: 24px;
  }
  .xml-panel {
    width: 340px;
    border-left: 1px solid #e1dfdd;
    background: #fff;
    display: flex;
    flex-direction: column;
    min-height: 0;
  }
  .xml-head {
    font-size: 11px;
    font-weight: 600;
    color: #605e5c;
    padding: 6px 10px;
    border-bottom: 1px solid #edebe9;
    background: #faf9f8;
  }
  .btn-inline {
    display: inline-flex;
    align-items: center;
    height: 30px;
    padding: 0 8px;
    border: 1px solid transparent;
    border-radius: 4px;
    cursor: pointer;
    font-size: 14px;
  }
  .btn-inline:hover { background: #edebe9; border-color: #d2d0ce; }
  .loading { text-align: center; color: #888; }
  .statusbar {
    display: flex;
    align-items: center;
    gap: 12px;
    background: #2b579a;
    color: #fff;
    font-size: 12px;
    padding: 3px 12px;
  }
  .statusbar .spacer { flex: 1; }
  .statusbar .count { opacity: 0.9; }
  .zoom { display: flex; align-items: center; gap: 4px; }
  .zoom button {
    background: rgba(255, 255, 255, 0.15);
    color: #fff;
    border: none;
    width: 20px;
    height: 20px;
    border-radius: 3px;
    cursor: pointer;
    line-height: 1;
  }
  .zoom-val { min-width: 38px; text-align: center; }
  .findbar {
    display: flex;
    align-items: center;
    gap: 8px;
    padding: 6px 12px;
    background: #fff;
    border-bottom: 1px solid #e1dfdd;
  }
  .findbar input {
    height: 28px;
    border: 1px solid #d2d0ce;
    border-radius: 4px;
    padding: 0 8px;
    font-size: 13px;
  }
  .findbar button {
    height: 28px;
    padding: 0 12px;
    border: 1px solid #d2d0ce;
    border-radius: 4px;
    background: #f3f2f1;
    cursor: pointer;
    font-size: 13px;
  }
  .findbar button:disabled { opacity: 0.5; cursor: default; }
  .find-status { font-size: 12px; color: #605e5c; }
  .find-close { margin-left: auto; }
</style>
