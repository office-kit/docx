<script lang="ts">
  import { onMount } from 'svelte';
  import {
    createEditor,
    openEditor,
    runCommand,
    commands,
    highlightCss,
    runAtPath,
    paragraphAt,
    orderSelection,
    paragraphsInRange,
    runsInRange,
    caretAt,
    createStyleResolver,
    type ResolvedRunFormat,
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
    type RunFormatting,
    type HighlightColor,
    getParagraphStyle,
    PAGE_SIZE_A4,
    PAGE_SIZE_LETTER,
    MARGINS_NORMAL,
    setPageSize,
    setPageMargins,
    ensureHeadingStyles,
  } from '@office-kit/docx';
  import { editorFor } from '@office-kit/docx-editor';
  import EditorCanvas from '$lib/editor/EditorCanvas.svelte';
  import RawXmlInspector from '$lib/editor/RawXmlInspector.svelte';
  import RibbonIcon, { type IconName } from '$lib/editor/RibbonIcon.svelte';
  import { t, highlightName, locale, setLocale, LOCALES, type LocaleId, type MessageKey } from '$lib/editor/i18n.svelte';

  // Word's ribbon tabs, in Word's order. Draw and Design are left out: the
  // library has no ink or document-theme editing for them to drive.
  type TabId = 'home' | 'insert' | 'layout' | 'references' | 'mailings' | 'review' | 'view';
  const TAB_IDS: TabId[] = ['home', 'insert', 'layout', 'references', 'mailings', 'review', 'view'];

  let model = $state<EditorModel | null>(null);
  let version = $state(0);
  let tick = $state(0);
  let tab = $state<TabId>('home');
  let fileName = $state('Document1.docx');
  let status = $state('');
  let showXml = $state(false);
  let showMarks = $state(false);
  let zoom = $state(1);
  let wordCount = $state(0);
  let charCount = $state(0);

  // Find & replace lives in a Navigation pane on the left, as in Word.
  let showFind = $state(false);
  let findQuery = $state('');
  let replaceValue = $state('');
  let findStatus = $state('');

  /** Which drop-down menu is open (one at a time, like Word's ribbon). */
  type MenuId = 'underline' | 'highlight' | 'fontColor' | 'lineSpacing';
  let openMenu = $state<MenuId | null>(null);
  let fontInput = $state<HTMLInputElement | null>(null);
  let sizeInput = $state<HTMLInputElement | null>(null);
  let appEl = $state<HTMLDivElement | null>(null);
  // Distance from the top of the document to the editor: the site header sits
  // above it, and the editor fills the rest of the window like Word's.
  let appTop = $state(0);

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
  let vertAlignState = $state<FieldState<string>>(NONE);
  let styleState = $state<FieldState<string>>(NONE);
  // What the Font Color / Highlight buttons apply on a plain click: the last
  // color picked from their menus (Word's split-button behavior).
  let fontColorPen = $state('C00000');
  let highlightPen = $state<HighlightColor>('yellow');

  // Word's font size drop-down list.
  const FONT_SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];
  const FONT_CHOICES = ['Aptos', 'Arial', 'Calibri', 'Cambria', 'Courier New', 'Georgia', 'Times New Roman', 'Verdana', 'Yu Gothic', 'Yu Mincho'];
  // Word's "Standard Colors" row.
  const STANDARD_COLORS = ['C00000', 'FF0000', 'FFC000', 'FFFF00', '92D050', '00B050', '00B0F0', '0070C0', '002060', '7030A0'];
  // The highlight palette in Word's order; white is valid OOXML but Word does
  // not offer it.
  const HIGHLIGHT_MENU: HighlightColor[] = [
    'yellow', 'green', 'cyan', 'magenta', 'blue', 'red', 'darkBlue', 'darkCyan',
    'darkGreen', 'darkMagenta', 'darkRed', 'darkYellow', 'darkGray', 'lightGray', 'black',
  ];
  const LINE_SPACINGS = [1, 1.15, 1.5, 2, 2.5, 3];
  const UNDERLINES = [
    { style: 'single', key: 'underline.single' },
    { style: 'double', key: 'underline.double' },
    { style: 'thick', key: 'underline.thick' },
    { style: 'dotted', key: 'underline.dotted' },
    { style: 'wave', key: 'underline.wave' },
  ] as const satisfies ReadonlyArray<{ style: NonNullable<RunFormatting['underline']>; key: MessageKey }>;
  // The gallery tiles: styles every document gets from ensureHeadingStyles
  // (Normal is the absence of pStyle).
  const STYLE_GALLERY = [
    { id: undefined, key: 'style.normal' },
    { id: 'Heading1', key: 'style.heading1' },
    { id: 'Heading2', key: 'style.heading2' },
    { id: 'Heading3', key: 'style.heading3' },
  ] as const satisfies ReadonlyArray<{ id: string | undefined; key: MessageKey }>;
  const HEX6 = /^[0-9a-fA-F]{6}$/;
  const MIN_ZOOM = 0.1;
  const MAX_ZOOM = 5;

  const vertAlignCommand = commands.setVertAlignCommand;

  onMount(() => {
    const measure = (): void => {
      if (appEl) appTop = appEl.getBoundingClientRect().top + window.scrollY;
    };
    measure();
    window.addEventListener('resize', measure);
    loadModel(editorFor(sampleDoc()));
    // Dev-only handle so the editor can be driven/inspected from the console
    // (and by the e2e verification). Stripped from production builds.
    if (import.meta.env.DEV) Object.assign(window, { wkEditorModel: () => model });
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

  /**
   * Run a command and force a structural re-render. Commands are atomic: a
   * failure (e.g. rejected input) leaves the document untouched, so it is
   * reported in the status bar rather than crashing the page.
   */
  function apply<P, R>(cmd: Command<P, R>, params: P): R | undefined {
    openMenu = null;
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
    status = '';
    return result;
  }

  /** Reset every piece of chrome derived from the previous document. */
  function loadModel(next: EditorModel): void {
    // Word opens a document with the caret at its start.
    if (!next.selection && next.doc.document.body.blocks[0]?.kind === 'paragraph') {
      next.setSelection(caretAt({ block: 0, inline: 0, offset: 0 }));
    }
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
    return tick >= 0 && !!model && !!model.selection && (cmd.isEnabled?.(model) ?? true);
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
   * Reflect the selection's formatting in the ribbon the way Word does: the
   * effective value (styles and document defaults included), so a caret in a
   * Heading 1 reads its style's font and size. A caret reads the run it is in
   * (or the paragraph mark in an empty paragraph); a range reads every run it
   * covers, so a differently formatted run in the middle makes a field "mixed".
   */
  function syncToolbarFromCaret(): void {
    const sel = model?.selection;
    if (!model || !sel) {
      fontState = sizeState = vertAlignState = styleState = NONE;
      return;
    }
    const doc = model.doc;
    const styles = createStyleResolver(doc);
    const ordered = orderSelection(sel);
    // paragraphsInRange takes every cell of a table in range, so a caret uses
    // only its own paragraph.
    const caretPara = ordered.collapsed ? paragraphAt(doc, sel.focus) : undefined;
    const paras = ordered.collapsed ? (caretPara ? [caretPara] : []) : paragraphsInRange(doc, ordered);
    let formats: ResolvedRunFormat[];
    if (caretPara) {
      const run = runAtPath(doc, sel.focus);
      formats = [styles.run(caretPara, run)];
    } else {
      const owner = new Map(paras.flatMap((p) => p.children.filter((c) => c.kind === 'run').map((r) => [r, p] as const)));
      formats = runsInRange(doc, ordered).flatMap((run) => {
        const para = owner.get(run);
        return para ? [styles.run(para, run)] : [];
      });
    }
    fontState = fieldOf(formats.map((f) => (f.font === undefined ? undefined : `${f.font}${f.fontRole ? ` ${t(`font.${f.fontRole}`)}` : ''}`)));
    sizeState = fieldOf(formats.map((f) => (f.sizeHalfPoints === undefined ? undefined : f.sizeHalfPoints / 2)));
    vertAlignState = fieldOf(formats.map((f) => f.vertAlign ?? 'baseline'));
    styleState = fieldOf(paras.map((p) => getParagraphStyle(p)));
  }

  /** Text shown in the Font / Size boxes: Word leaves them blank when mixed. */
  function fieldText<T>(state: FieldState<T>): string {
    return state.kind === 'value' ? String(state.value) : '';
  }

  function fieldPlaceholder(state: FieldState<unknown>): string {
    if (state.kind === 'none') return '';
    if (state.kind === 'mixed') return '';
    return t('state.inherited');
  }

  function onFontChange(e: Event): void {
    const input = e.currentTarget as HTMLInputElement;
    // "Calibri (Body)" names the theme font; typing it back means that font.
    const font = input.value.replace(/\s*\([^)]*\)$/, '').trim();
    if (font) apply(commands.setFontCommand, { font });
    else input.value = fieldText(fontState);
  }

  function onFontSizeChange(e: Event): void {
    const input = e.currentTarget as HTMLInputElement;
    const points = Number(input.value);
    // Word accepts 1–1638 pt in half-point steps; reject anything else without
    // touching the document, and put the control back to the selection's state.
    if (input.value === '' || !Number.isFinite(points) || points < 1 || points > 1638) {
      input.value = fieldText(sizeState);
      return;
    }
    apply(commands.setFontSizeCommand, { points });
  }

  /**
   * Increase / Decrease Font Size step through Word's size list; past either
   * end Word moves in 10 pt steps (above 72) or 1 pt steps (below 8).
   */
  function stepFontSize(direction: 1 | -1): void {
    if (sizeState.kind !== 'value') return;
    const current = sizeState.value;
    const listed = direction > 0 ? FONT_SIZES.find((s) => s > current) : FONT_SIZES.findLast((s) => s < current);
    const fallback = direction > 0 ? Math.floor(current / 10) * 10 + 10 : Math.max(1, Math.ceil(current) - 1);
    apply(commands.setFontSizeCommand, { points: listed ?? fallback });
  }

  function toggleVertAlign(val: 'superscript' | 'subscript'): void {
    const on = vertAlignState.kind === 'value' && vertAlignState.value === val;
    apply(vertAlignCommand, { val: on ? undefined : val });
  }

  function lineSpacingNow(): number | undefined {
    return tick >= 0 && model ? commands.lineSpacingOf(model) : undefined;
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
    status = '';
  }

  function undo(): void {
    if (!model) return;
    model.undo();
    version++;
    tick++;
    refreshCounts();
    syncToolbarFromCaret();
  }

  function redo(): void {
    if (!model) return;
    model.redo();
    version++;
    tick++;
    refreshCounts();
    syncToolbarFromCaret();
  }

  function doReplaceAll(): void {
    if (!findQuery) return;
    const n = apply(commands.replaceAllCommand, { query: findQuery, replacement: replaceValue });
    if (n === undefined) return;
    findStatus = `Replaced ${n} occurrence${n === 1 ? '' : 's'}.`;
  }

  function setZoom(z: number): void {
    zoom = Math.max(MIN_ZOOM, Math.min(MAX_ZOOM, Math.round(z * 100) / 100));
  }

  function onGlobalKey(e: KeyboardEvent): void {
    const mod = e.ctrlKey || e.metaKey;
    if (e.key === 'Escape') openMenu = null;
    if (mod && (e.key.toLowerCase() === 'h' || e.key.toLowerCase() === 'f')) {
      e.preventDefault();
      showFind = true;
    }
  }

  /** Close an open drop-down when the pointer goes down outside it. */
  function onGlobalPointer(e: PointerEvent): void {
    if (openMenu && !(e.target as Element).closest('.split')) openMenu = null;
  }

  /**
   * Ribbon buttons must not take focus from the page: Word keeps the text
   * selection while you click formatting buttons, and a focused button would
   * collapse the canvas selection before the command runs.
   */
  function keepSelection(e: MouseEvent): void {
    if ((e.target as Element).closest('button')) e.preventDefault();
  }

  function toggleMenu(id: MenuId): void {
    openMenu = openMenu === id ? null : id;
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
      status = '';
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
    fileName = 'Document1.docx';
    status = '';
  }

  /** Cut / Copy go through the browser so the canvas' own clipboard handlers run. */
  function clipboard(action: 'cut' | 'copy'): void {
    document.execCommand(action);
  }

  async function paste(): Promise<void> {
    // Reading the clipboard needs the browser's permission; a refusal is shown
    // in the status bar and Cmd/Ctrl+V keeps working.
    try {
      const text = await navigator.clipboard.readText();
      if (text) apply(commands.insertTextCommand, { text });
    } catch (err) {
      status = `Paste failed: ${(err as Error).message}`;
    }
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

  function addComment(): void {
    const text = prompt('Comment');
    if (text) apply(commands.addCommentCommand, { author: 'You', initials: 'Y', text });
  }

  function insertMergeField(): void {
    const fieldName = prompt('Merge field name');
    if (fieldName) apply(commands.insertMergeFieldCommand, { fieldName });
  }

  async function insertImage(e: Event): Promise<void> {
    const input = e.target as HTMLInputElement;
    const file = input.files?.[0];
    if (!file || !model) return;
    const bytes = new Uint8Array(await file.arrayBuffer());
    apply(commands.insertImageCommand, { bytes, options: { widthEmu: 2743200, heightEmu: 2057400 } });
    input.value = '';
  }

  function hexCss(hex: string): string {
    return HEX6.test(hex) ? `#${hex}` : 'transparent';
  }
</script>

<svelte:head><title>word-kit editor</title></svelte:head>
<svelte:window onkeydown={onGlobalKey} onpointerdown={onGlobalPointer} />

{#snippet iconButton(icon: IconName, tip: string, onclick: () => void, opts: { on?: boolean; disabled?: boolean } = {})}
  <button class="rb" class:on={opts.on} disabled={opts.disabled} {onclick} title={tip} aria-label={tip} aria-pressed={opts.on}><RibbonIcon name={icon} /></button>
{/snippet}

{#snippet glyphButton(glyph: string, cls: string, tip: string, onclick: () => void, opts: { on?: boolean; disabled?: boolean } = {})}
  <button class="rb glyph {cls}" class:on={opts.on} disabled={opts.disabled} {onclick} title={tip} aria-label={tip} aria-pressed={opts.on}>{@html glyph}</button>
{/snippet}

{#snippet bigButton(icon: IconName, label: string, onclick: () => void, disabled = false)}
  <button class="big" {onclick} {disabled} title={label}><RibbonIcon name={icon} size={32} /><span>{label}</span></button>
{/snippet}

<div class="app" class:marks={showMarks} bind:this={appEl} style="--app-top: {appTop}px">
  <!-- Title bar: Quick Access Toolbar, document name, search (Word for Mac). -->
  <header class="titlebar" onmousedown={keepSelection} role="toolbar" tabindex="-1">
    <div class="qat">
      <button class="qb" onclick={newDoc} title={t('action.new')} aria-label={t('action.new')}><RibbonIcon name="newDoc" size={18} /></button>
      <label class="qb" title={t('action.open')} aria-label={t('action.open')}><RibbonIcon name="open" size={18} /><input type="file" accept=".docx" onchange={openFile} hidden /></label>
      <button class="qb" onclick={download} title={t('action.save')} aria-label={t('action.save')}><RibbonIcon name="save" size={18} /></button>
      <button class="qb" onclick={undo} disabled={tick < 0 || !model?.canUndo()} title={`${t('action.undo')} (⌘Z)`} aria-label={t('action.undo')}><RibbonIcon name="undo" size={18} /></button>
      <button class="qb" onclick={redo} disabled={tick < 0 || !model?.canRedo()} title={`${t('action.redo')} (⌘Y)`} aria-label={t('action.redo')}><RibbonIcon name="redo" size={18} /></button>
    </div>
    <input class="docname" bind:value={fileName} aria-label="File name" spellcheck="false" />
    <label class="search">
      <RibbonIcon name="search" size={16} />
      <input
        placeholder={t('search.placeholder')}
        bind:value={findQuery}
        onfocus={() => (showFind = true)}
        aria-label={t('search.placeholder')}
      />
    </label>
  </header>

  <!-- Ribbon tabs, with Comments / Save on the right like Word's Comments / Share. -->
  <div class="tabs" onmousedown={keepSelection} role="toolbar" tabindex="-1">
    <div class="tablist" role="tablist">
      {#each TAB_IDS as id (id)}
        <button class="tab" class:active={tab === id} role="tab" aria-selected={tab === id} onclick={() => (tab = id)}>{t(`tab.${id}`)}</button>
      {/each}
    </div>
    <div class="tab-actions">
      <button class="pill" onclick={addComment} disabled={!model?.selection}><RibbonIcon name="comment" size={16} />{t('action.comments')}</button>
      <button class="pill" onclick={download}><RibbonIcon name="share" size={16} />{t('action.download')}</button>
    </div>
  </div>

  <!-- Ribbon body. Word for Mac shows no group captions; groups are split by rules. -->
  <div class="ribbon" onmousedown={keepSelection} role="toolbar" tabindex="-1">
    {#if tab === 'home'}
      <div class="group" aria-label={t('group.clipboard')}>
        <button class="big paste" onclick={paste} title={t('tip.paste')}><RibbonIcon name="paste" size={32} /><span>{t('tip.paste')}</span></button>
        <div class="col">
          {@render iconButton('cut', t('tip.cut'), () => clipboard('cut'), { disabled: !model?.selection })}
          {@render iconButton('copy', t('tip.copy'), () => clipboard('copy'), { disabled: !model?.selection })}
        </div>
      </div>

      <div class="group" aria-label={t('group.font')}>
        <div class="rows">
          <div class="row">
            <span class="combo-wrap">
            <input
              bind:this={fontInput}
              class="combo font"
              list="wk-fonts"
              value={fieldText(fontState)}
              placeholder={fieldPlaceholder(fontState)}
              disabled={!enabled(commands.setFontCommand)}
              onchange={onFontChange}
              title={t('tip.fontName')}
              aria-label={t('tip.fontName')}
              spellcheck="false"
            />
            <button class="combo-arrow" onclick={() => fontInput?.showPicker()} disabled={!enabled(commands.setFontCommand)} aria-label={t('tip.fontName')} tabindex="-1"><RibbonIcon name="chevronDown" size={10} /></button>
            </span>
            <datalist id="wk-fonts">{#each FONT_CHOICES as font (font)}<option value={font}></option>{/each}</datalist>
            <span class="combo-wrap">
            <input
              bind:this={sizeInput}
              class="combo size"
              list="wk-sizes"
              inputmode="decimal"
              value={fieldText(sizeState)}
              disabled={!enabled(commands.setFontSizeCommand)}
              onchange={onFontSizeChange}
              title={t('tip.fontSize')}
              aria-label={t('tip.fontSize')}
            />
            <button class="combo-arrow" onclick={() => sizeInput?.showPicker()} disabled={!enabled(commands.setFontSizeCommand)} aria-label={t('tip.fontSize')} tabindex="-1"><RibbonIcon name="chevronDown" size={10} /></button>
            </span>
            <datalist id="wk-sizes">{#each FONT_SIZES as size (size)}<option value={size}></option>{/each}</datalist>
            {@render glyphButton('A<sup>^</sup>', 'grow', t('tip.grow'), () => stepFontSize(1), { disabled: sizeState.kind !== 'value' })}
            {@render glyphButton('A<sup>ˇ</sup>', 'grow', t('tip.shrink'), () => stepFontSize(-1), { disabled: sizeState.kind !== 'value' })}
            <span class="sep"></span>
            {@render glyphButton('A<i class="eraser"></i>', 'clear', t('tip.clearAll'), () => apply(commands.clearFormattingCommand, undefined), { disabled: !enabled(commands.clearFormattingCommand) })}
          </div>
          <div class="row">
            {@render glyphButton('B', 'b', t('tip.bold') + ' (⌘B)', () => apply(commands.toggleBoldCommand, undefined), { on: active(commands.toggleBoldCommand), disabled: !enabled(commands.toggleBoldCommand) })}
            {@render glyphButton('I', 'i', t('tip.italic') + ' (⌘I)', () => apply(commands.toggleItalicCommand, undefined), { on: active(commands.toggleItalicCommand), disabled: !enabled(commands.toggleItalicCommand) })}
            <span class="split" class:open={openMenu === 'underline'}>
              {@render glyphButton('U', 'u', t('tip.underline') + ' (⌘U)', () => apply(commands.toggleUnderlineCommand, undefined), { on: active(commands.toggleUnderlineCommand), disabled: !enabled(commands.toggleUnderlineCommand) })}
              <button class="arrow" onclick={() => toggleMenu('underline')} disabled={!enabled(commands.setUnderlineStyleCommand)} aria-label={t('tip.more')} aria-expanded={openMenu === 'underline'}><RibbonIcon name="chevronDown" size={10} /></button>
              {#if openMenu === 'underline'}
                <div class="menu" role="menu">
                  {#each UNDERLINES as u (u.style)}
                    <button class="mi" role="menuitem" onclick={() => apply(commands.setUnderlineStyleCommand, { style: u.style })}><span class="ul-sample ul-{u.style}">{t(u.key)}</span></button>
                  {/each}
                </div>
              {/if}
            </span>
            {@render glyphButton('ab', 'strike', t('tip.strike'), () => apply(commands.toggleStrikeCommand, undefined), { on: active(commands.toggleStrikeCommand), disabled: !enabled(commands.toggleStrikeCommand) })}
            {@render glyphButton('x<sub>2</sub>', 'script', t('tip.subscript'), () => toggleVertAlign('subscript'), { on: vertAlignState.kind === 'value' && vertAlignState.value === 'subscript', disabled: !enabled(vertAlignCommand) })}
            {@render glyphButton('x<sup>2</sup>', 'script', t('tip.superscript'), () => toggleVertAlign('superscript'), { on: vertAlignState.kind === 'value' && vertAlignState.value === 'superscript', disabled: !enabled(vertAlignCommand) })}
            <span class="sep"></span>
            <span class="split" class:open={openMenu === 'highlight'}>
              <button class="rb" onclick={() => apply(commands.setHighlightCommand, { color: highlightPen })} disabled={!enabled(commands.setHighlightCommand)} title={t('tip.highlight')} aria-label={t('tip.highlight')}>
                <span class="swatch-icon"><svg viewBox="0 0 20 20" width="20" height="16" fill="none" stroke="currentColor" stroke-width="1.2" stroke-linejoin="round" aria-hidden="true"><path d="M5 12.5 12.5 5l2.5 2.5L7.5 15H5zM11 6.5l2.5 2.5" /></svg><i style="background:{highlightCss(highlightPen)}"></i></span>
              </button>
              <button class="arrow" onclick={() => toggleMenu('highlight')} disabled={!enabled(commands.setHighlightCommand)} aria-label={t('tip.more')} aria-expanded={openMenu === 'highlight'}><RibbonIcon name="chevronDown" size={10} /></button>
              {#if openMenu === 'highlight'}
                <div class="menu palette" role="menu">
                  <div class="grid five">
                    {#each HIGHLIGHT_MENU as color (color)}
                      <button class="chip" role="menuitem" style="background:{highlightCss(color)}" title={highlightName(color)} aria-label={highlightName(color)} onclick={() => { highlightPen = color; apply(commands.setHighlightCommand, { color }); }}></button>
                    {/each}
                  </div>
                  <button class="mi" role="menuitem" onclick={() => apply(commands.setHighlightCommand, { color: 'none' })}>{t('highlight.none')}</button>
                </div>
              {/if}
            </span>
            <span class="split" class:open={openMenu === 'fontColor'}>
              <button class="rb" onclick={() => apply(commands.setColorCommand, { color: fontColorPen })} disabled={!enabled(commands.setColorCommand)} title={t('tip.fontColor')} aria-label={t('tip.fontColor')}>
                <span class="swatch-icon"><span class="a">A</span><i style="background:{fontColorPen === 'auto' ? '#000' : hexCss(fontColorPen)}"></i></span>
              </button>
              <button class="arrow" onclick={() => toggleMenu('fontColor')} disabled={!enabled(commands.setColorCommand)} aria-label={t('tip.more')} aria-expanded={openMenu === 'fontColor'}><RibbonIcon name="chevronDown" size={10} /></button>
              {#if openMenu === 'fontColor'}
                <div class="menu palette" role="menu">
                  <button class="mi" role="menuitem" onclick={() => { fontColorPen = 'auto'; apply(commands.setColorCommand, { color: 'auto' }); }}><i class="auto-chip"></i>{t('color.auto')}</button>
                  <div class="menu-head">{t('color.standard')}</div>
                  <div class="grid ten">
                    {#each STANDARD_COLORS as hex (hex)}
                      <button class="chip" role="menuitem" style="background:#{hex}" title={`#${hex}`} aria-label={`#${hex}`} onclick={() => { fontColorPen = hex; apply(commands.setColorCommand, { color: hex }); }}></button>
                    {/each}
                  </div>
                  <label class="mi">
                    {t('color.more')}
                    <input
                      type="color"
                      value={HEX6.test(fontColorPen) ? `#${fontColorPen}` : '#000000'}
                      onchange={(e) => { fontColorPen = (e.currentTarget as HTMLInputElement).value.slice(1).toUpperCase(); apply(commands.setColorCommand, { color: fontColorPen }); }}
                      hidden
                    />
                  </label>
                </div>
              {/if}
            </span>
          </div>
        </div>
      </div>

      <div class="group" aria-label={t('group.paragraph')}>
        <div class="rows">
          <div class="row">
            {@render iconButton('bullets', t('tip.bullets'), () => apply(commands.applyListCommand, { kind: 'bullet' }), { disabled: !enabled(commands.applyListCommand) })}
            {@render iconButton('numbering', t('tip.numbering'), () => apply(commands.applyListCommand, { kind: 'numbered' }), { disabled: !enabled(commands.applyListCommand) })}
            <span class="sep"></span>
            {@render iconButton('indentLess', t('tip.indentLess'), () => apply(commands.indentStepCommand, { direction: 'decrease' }), { disabled: !enabled(commands.indentStepCommand) })}
            {@render iconButton('indentMore', t('tip.indentMore'), () => apply(commands.indentStepCommand, { direction: 'increase' }), { disabled: !enabled(commands.indentStepCommand) })}
            <span class="sep"></span>
            {@render iconButton('pilcrow', t('tip.marks') + ' (⌘8)', () => (showMarks = !showMarks), { on: showMarks })}
          </div>
          <div class="row">
            {@render iconButton('alignLeft', t('tip.alignLeft') + ' (⌘L)', () => apply(commands.alignLeftCommand, undefined), { on: active(commands.alignLeftCommand), disabled: !enabled(commands.alignLeftCommand) })}
            {@render iconButton('alignCenter', t('tip.center') + ' (⌘E)', () => apply(commands.alignCenterCommand, undefined), { on: active(commands.alignCenterCommand), disabled: !enabled(commands.alignCenterCommand) })}
            {@render iconButton('alignRight', t('tip.alignRight') + ' (⌘R)', () => apply(commands.alignRightCommand, undefined), { on: active(commands.alignRightCommand), disabled: !enabled(commands.alignRightCommand) })}
            {@render iconButton('alignJustify', t('tip.justify') + ' (⌘J)', () => apply(commands.alignJustifyCommand, undefined), { on: active(commands.alignJustifyCommand), disabled: !enabled(commands.alignJustifyCommand) })}
            <span class="sep"></span>
            <span class="split" class:open={openMenu === 'lineSpacing'}>
              <button class="rb" onclick={() => toggleMenu('lineSpacing')} disabled={!enabled(commands.setLineSpacingCommand)} title={t('tip.lineSpacing')} aria-label={t('tip.lineSpacing')} aria-expanded={openMenu === 'lineSpacing'}><RibbonIcon name="lineSpacing" /></button>
              <button class="arrow" onclick={() => toggleMenu('lineSpacing')} disabled={!enabled(commands.setLineSpacingCommand)} aria-label={t('tip.lineSpacing')}><RibbonIcon name="chevronDown" size={10} /></button>
              {#if openMenu === 'lineSpacing'}
                <div class="menu" role="menu">
                  {#each LINE_SPACINGS as multiple (multiple)}
                    <button class="mi check" class:checked={lineSpacingNow() === multiple} role="menuitemradio" aria-checked={lineSpacingNow() === multiple} onclick={() => apply(commands.setLineSpacingCommand, { multiple })}>{multiple.toFixed(multiple === 1.15 ? 2 : 1)}</button>
                  {/each}
                </div>
              {/if}
            </span>
          </div>
        </div>
      </div>

      <div class="group gallery-group" aria-label={t('group.styles')}>
        <div class="gallery" role="listbox" aria-label={t('group.styles')}>
          {#each STYLE_GALLERY as s (s.key)}
            {@const selected = s.id === undefined ? styleState.kind === 'inherited' : styleState.kind === 'value' && styleState.value === s.id}
            <button
              class="tile"
              class:selected
              role="option"
              aria-selected={selected}
              disabled={!enabled(commands.setParagraphStyleCommand)}
              onclick={() => apply(commands.setParagraphStyleCommand, { styleId: s.id })}
              title={t(s.key)}
            >
              <span class="sample sample-{s.id ?? 'Normal'}">AaBbCcDdEe</span>
              <span class="tile-name">{t(s.key)}</span>
            </button>
          {/each}
        </div>
      </div>
    {:else if tab === 'insert'}
      <div class="group" aria-label={t('group.pages')}>
        {@render bigButton('pageBreak', t('ins.pageBreak'), () => apply(commands.insertPageBreakCommand, undefined), !model?.selection)}
      </div>
      <div class="group" aria-label={t('group.tables')}>
        {@render bigButton('table', t('ins.table'), insertTable)}
      </div>
      <div class="group" aria-label={t('group.illustrations')}>
        <label class="big" title={t('ins.picture')}><RibbonIcon name="picture" size={32} /><span>{t('ins.picture')}</span><input type="file" accept="image/*" onchange={insertImage} hidden /></label>
      </div>
      <div class="group" aria-label={t('group.links')}>
        {@render bigButton('link', t('ins.link'), insertLink)}
        {@render bigButton('bookmark', t('ref.bookmark'), () => { const n = prompt('Bookmark name'); if (n) apply(commands.addBookmarkCommand, { name: n }); })}
      </div>
      <div class="group" aria-label={t('group.comments')}>
        {@render bigButton('comment', t('ins.comment'), addComment, !model?.selection)}
      </div>
      <div class="group" aria-label={t('group.headerFooter')}>
        {@render bigButton('pageNumber', t('layout.pageNumbers'), () => apply(commands.addPageNumberFooterCommand, {}))}
      </div>
    {:else if tab === 'layout'}
      <div class="group" aria-label={t('group.pageSetup')}>
        {@render bigButton('orientation', t('layout.portrait'), () => apply(commands.setOrientationCommand, { orientation: 'portrait' }))}
        {@render bigButton('orientation', t('layout.landscape'), () => apply(commands.setOrientationCommand, { orientation: 'landscape' }))}
        {@render bigButton('size', 'A4', () => apply(commands.setPageSizeCommand, { size: PAGE_SIZE_A4 }))}
        {@render bigButton('size', 'Letter', () => apply(commands.setPageSizeCommand, { size: PAGE_SIZE_LETTER }))}
        {@render bigButton('sectionBreak', t('layout.sectionBreak'), () => apply(commands.insertSectionBreakCommand, { type: 'nextPage' }), !model?.selection)}
      </div>
      <div class="group" aria-label={t('group.paragraph')}>
        <div class="rows">
          <div class="row">
            {@render iconButton('indentLess', t('tip.indentLess'), () => apply(commands.indentStepCommand, { direction: 'decrease' }), { disabled: !enabled(commands.indentStepCommand) })}
            {@render iconButton('indentMore', t('tip.indentMore'), () => apply(commands.indentStepCommand, { direction: 'increase' }), { disabled: !enabled(commands.indentStepCommand) })}
          </div>
        </div>
      </div>
    {:else if tab === 'references'}
      <div class="group" aria-label={t('group.toc')}>
        {@render bigButton('toc', t('ref.toc'), () => apply(commands.insertTocCommand, {}))}
      </div>
      <div class="group" aria-label={t('group.footnotes')}>
        {@render bigButton('footnote', t('ref.footnote'), () => { const txt = prompt('Footnote text'); if (txt) apply(commands.addFootnoteCommand, { text: txt }); }, !model?.selection)}
        {@render bigButton('endnote', t('ref.endnote'), () => { const txt = prompt('Endnote text'); if (txt) apply(commands.addEndnoteCommand, { text: txt }); }, !model?.selection)}
      </div>
    {:else if tab === 'mailings'}
      <div class="group" aria-label={t('group.fields')}>
        {@render bigButton('newDoc', t('mail.mergeField'), insertMergeField, !model?.selection)}
      </div>
    {:else if tab === 'review'}
      <div class="group" aria-label={t('group.comments')}>
        {@render bigButton('comment', t('review.newComment'), addComment, !model?.selection)}
      </div>
      <div class="group" aria-label={t('group.changes')}>
        {@render bigButton('accept', t('review.acceptAll'), () => apply(commands.acceptAllRevisionsCommand, undefined))}
        {@render bigButton('reject', t('review.rejectAll'), () => apply(commands.rejectAllRevisionsCommand, undefined))}
      </div>
    {:else if tab === 'view'}
      <div class="group" aria-label={t('group.show')}>
        <label class="check-item"><input type="checkbox" bind:checked={showFind} />{t('view.navigation')}</label>
        <label class="check-item"><input type="checkbox" bind:checked={showXml} />{t('view.xml')}</label>
      </div>
      <div class="group" aria-label={t('group.zoom')}>
        {@render bigButton('zoomIn', t('view.zoomIn'), () => setZoom(zoom + 0.1))}
        {@render bigButton('zoomOut', t('view.zoomOut'), () => setZoom(zoom - 0.1))}
        {@render bigButton('onePage', t('view.zoom100'), () => setZoom(1))}
      </div>
      <div class="group" aria-label={t('group.language')}>
        <label class="big lang" title={t('language')}>
          <RibbonIcon name="language" size={32} />
          <select value={locale()} onchange={(e) => setLocale((e.currentTarget as HTMLSelectElement).value as LocaleId)} aria-label={t('language')}>
            {#each Object.entries(LOCALES) as [id, name] (id)}
              <option value={id}>{name}</option>
            {/each}
          </select>
        </label>
      </div>
    {/if}
  </div>

  <!-- Editing surface -->
  <div class="workspace">
    {#if showFind}
      <aside class="nav-pane" aria-label={t('find.title')}>
        <div class="pane-head">
          <span>{t('find.title')}</span>
          <button class="pane-close" onclick={() => (showFind = false)} aria-label={t('find.close')}><RibbonIcon name="close" size={14} /></button>
        </div>
        <label class="pane-search">
          <RibbonIcon name="search" size={14} />
          <input placeholder={t('find.find')} bind:value={findQuery} aria-label={t('find.find')} />
        </label>
        <input class="pane-input" placeholder={t('find.replaceWith')} bind:value={replaceValue} aria-label={t('find.replaceWith')} />
        <button class="pane-button" onclick={doReplaceAll} disabled={!findQuery}>{t('find.replaceAll')}</button>
        <p class="find-status" aria-live="polite">{findStatus}</p>
      </aside>
    {/if}
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

  <footer class="statusbar">
    <span title={`${charCount} ${t('status.chars')}`}>{wordCount} {t('status.words')}</span>
    {#if status}<span class="status-msg" role="status">{status}</span>{/if}
    <div class="spacer"></div>
    <div class="zoom">
      <button onclick={() => setZoom(zoom - 0.1)} aria-label={t('view.zoomOut')}>−</button>
      <input
        type="range"
        min={MIN_ZOOM * 100}
        max={MAX_ZOOM * 100}
        step="10"
        value={Math.round(zoom * 100)}
        oninput={(e) => setZoom(Number((e.currentTarget as HTMLInputElement).value) / 100)}
        aria-label="Zoom"
      />
      <button onclick={() => setZoom(zoom + 0.1)} aria-label={t('view.zoomIn')}>+</button>
      <span class="zoom-val">{Math.round(zoom * 100)}%</span>
    </div>
  </footer>
</div>

<style>
  /* Word for Mac (light appearance) chrome. */
  .app {
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
  button { font: inherit; color: inherit; }

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
  .pill:disabled { opacity: 0.45; cursor: default; }

  .ribbon {
    display: flex;
    align-items: stretch;
    min-height: 74px;
    padding: 6px 8px;
    border-bottom: 1px solid var(--chrome-line);
    overflow-x: auto;
    outline: none;
  }
  .group {
    display: flex;
    align-items: center;
    gap: 2px;
    padding: 0 10px;
    border-right: 1px solid var(--chrome-line);
  }
  .group:last-child { border-right: none; }
  .rows { display: flex; flex-direction: column; gap: 6px; }
  .row { display: flex; align-items: center; gap: 2px; }
  .col { display: flex; flex-direction: column; gap: 2px; }
  .sep { width: 1px; height: 20px; margin: 0 5px; background: var(--chrome-line); }

  .rb, .arrow, .big, .qb { color: var(--icon); }
  .rb {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    min-width: 28px;
    height: 28px;
    padding: 0 4px;
    border: none;
    border-radius: 4px;
    background: none;
    cursor: pointer;
  }
  .rb:hover:not(:disabled), .arrow:hover:not(:disabled), .big:hover, .tile:hover:not(:disabled), .mi:hover { background: var(--hover); }
  .rb.on { background: var(--pressed); }
  .rb:disabled, .arrow:disabled, .tile:disabled, .big:disabled { opacity: 0.35; cursor: default; }
  .glyph { font-size: 15px; line-height: 1; font-family: Georgia, 'Times New Roman', serif; }
  .glyph.b { font-weight: 700; font-family: -apple-system, BlinkMacSystemFont, sans-serif; }
  .glyph.i { font-style: italic; }
  .glyph.u { text-decoration: underline; text-underline-offset: 2px; }
  .glyph.strike { text-decoration: line-through; font-family: -apple-system, BlinkMacSystemFont, sans-serif; font-size: 13px; }
  .glyph.script :global(sub), .glyph.script :global(sup) { font-size: 9px; }
  .glyph.grow { font-family: -apple-system, BlinkMacSystemFont, sans-serif; font-size: 15px; }
  .glyph.grow :global(sup) { font-size: 9px; }
  .glyph.clear { font-family: -apple-system, BlinkMacSystemFont, sans-serif; position: relative; }
  .glyph.clear :global(.eraser) {
    position: absolute;
    right: 3px;
    bottom: 5px;
    width: 7px;
    height: 5px;
    border-radius: 1px;
    background: #d13438;
    transform: rotate(-35deg);
  }

  .split { position: relative; display: inline-flex; align-items: center; }
  .arrow {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 14px;
    height: 28px;
    border: none;
    border-radius: 4px;
    background: none;
    cursor: pointer;
    padding: 0;
  }
  .split.open .arrow { background: var(--pressed); }
  .swatch-icon { display: inline-flex; flex-direction: column; align-items: center; }
  .swatch-icon .a { font-size: 15px; line-height: 15px; font-family: -apple-system, BlinkMacSystemFont, sans-serif; }
  .swatch-icon i { display: block; width: 16px; height: 4px; margin-top: 1px; border: 0.5px solid rgba(0, 0, 0, 0.15); }

  .menu {
    position: absolute;
    top: calc(100% + 4px);
    left: 0;
    z-index: 20;
    min-width: 170px;
    padding: 6px;
    border: 1px solid var(--chrome-line);
    border-radius: 8px;
    background: #fff;
    box-shadow: 0 8px 24px rgba(0, 0, 0, 0.16);
  }
  .menu-head { padding: 6px 6px 4px; font-size: 11px; color: var(--muted); }
  .mi {
    display: flex;
    align-items: center;
    gap: 8px;
    width: 100%;
    padding: 5px 8px;
    border: none;
    border-radius: 4px;
    background: none;
    text-align: left;
    cursor: pointer;
    white-space: nowrap;
  }
  .mi.check::before { content: ''; width: 12px; }
  .mi.check.checked::before { content: '✓'; }
  .grid { display: grid; gap: 4px; padding: 4px; }
  .grid.five { grid-template-columns: repeat(5, 22px); }
  .grid.ten { grid-template-columns: repeat(10, 18px); gap: 3px; }
  .chip { width: 100%; aspect-ratio: 1; border: 1px solid rgba(0, 0, 0, 0.2); border-radius: 2px; cursor: pointer; padding: 0; }
  .chip:hover { outline: 2px solid var(--accent); outline-offset: 1px; }
  .auto-chip { width: 14px; height: 14px; background: #000; border: 1px solid rgba(0, 0, 0, 0.2); }
  .ul-sample { text-decoration-line: underline; text-underline-offset: 3px; }
  .ul-double { text-decoration-style: double; }
  .ul-thick { text-decoration-thickness: 2px; }
  .ul-dotted { text-decoration-style: dotted; }
  .ul-wave { text-decoration-style: wavy; }

  .combo {
    height: 24px;
    padding: 0 6px;
    border: 1px solid var(--control-line);
    border-radius: 4px;
    background: #fff;
    font: inherit;
    color: var(--text);
  }
  .combo:focus { border-color: var(--accent); outline: none; }
  .combo-wrap { position: relative; display: inline-flex; }
  .combo-wrap .combo { padding-right: 20px; }
  /* Chromium draws its own datalist indicator; Word shows one chevron. */
  .combo::-webkit-calendar-picker-indicator { display: none !important; }
  .combo-arrow {
    position: absolute;
    right: 1px;
    top: 1px;
    bottom: 1px;
    width: 18px;
    border: none;
    border-radius: 0 3px 3px 0;
    background: none;
    color: var(--icon);
    cursor: pointer;
    padding: 0;
  }
  .combo-arrow:hover:not(:disabled) { background: var(--hover); }
  .combo:disabled { background: #f7f7f7; color: var(--muted); }
  .combo.font { width: 140px; }
  .combo.size { width: 52px; }

  .big {
    display: inline-flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    gap: 4px;
    min-width: 56px;
    height: 64px;
    padding: 2px 6px;
    border: none;
    border-radius: 4px;
    background: none;
    cursor: pointer;
    font-size: 12px;
    white-space: nowrap;
  }
  .big.paste { color: var(--text); }
  .big.lang select { border: 1px solid var(--chrome-line); border-radius: 4px; background: #fff; font: inherit; }
  .check-item { display: flex; align-items: center; gap: 6px; padding: 2px 4px; white-space: nowrap; }
  .group:has(.check-item) { flex-direction: column; align-items: flex-start; justify-content: center; }

  .gallery-group { padding-right: 4px; }
  .gallery {
    display: flex;
    gap: 8px;
    padding: 4px 8px;
    border: 1px solid var(--control-line);
    border-radius: 4px;
    background: #fff;
  }
  .tile {
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: space-between;
    width: 72px;
    height: 52px;
    padding: 6px 4px 3px;
    border: 1px solid #dedede;
    border-radius: 3px;
    background: #f8f8f8;
    cursor: pointer;
    overflow: hidden;
  }
  .tile.selected { border-color: #a9bcd9; background: #ebf0f6; box-shadow: 0 0 0 1px #a9bcd9; }
  .sample { white-space: nowrap; overflow: hidden; max-width: 100%; color: #000; line-height: 1.1; }
  .sample-Normal { font-family: Calibri, Aptos, sans-serif; font-size: 12px; }
  /* Scaled previews of the styles ensureHeadingStyles writes (20/16/14 pt). */
  .sample-Heading1 { font-family: Calibri, Aptos, sans-serif; font-size: 16px; font-weight: 700; color: #1f497d; }
  .sample-Heading2 { font-family: Calibri, Aptos, sans-serif; font-size: 14px; font-weight: 700; color: #1f497d; }
  .sample-Heading3 { font-family: Calibri, Aptos, sans-serif; font-size: 13px; color: #4f81bd; }
  .tile-name { font-size: 11px; color: var(--text); }

  .workspace { flex: 1; display: flex; min-height: 0; background: var(--canvas-bg); }
  .surface { flex: 1; overflow: auto; padding: 24px 24px 48px; }
  .nav-pane {
    display: flex;
    flex-direction: column;
    gap: 8px;
    width: 260px;
    padding: 10px 12px;
    border-right: 1px solid var(--chrome-line);
    background: var(--chrome);
  }
  .pane-head { display: flex; align-items: center; justify-content: space-between; font-weight: 600; }
  .pane-close { border: none; background: none; cursor: pointer; padding: 4px; border-radius: 4px; }
  .pane-close:hover { background: var(--hover); }
  .pane-search, .pane-input {
    display: flex;
    align-items: center;
    gap: 6px;
    height: 28px;
    padding: 0 8px;
    border: 1px solid var(--chrome-line);
    border-radius: 6px;
    background: #fff;
    font: inherit;
  }
  .pane-search input { flex: 1; min-width: 0; border: none; outline: none; font: inherit; }
  .pane-button {
    height: 28px;
    border: 1px solid var(--chrome-line);
    border-radius: 6px;
    background: #fff;
    cursor: pointer;
  }
  .pane-button:hover:not(:disabled) { background: var(--hover); }
  .pane-button:disabled { opacity: 0.5; cursor: default; }
  .find-status { margin: 0; font-size: 12px; color: var(--muted); }

  .xml-panel {
    width: 340px;
    border-left: 1px solid var(--chrome-line);
    background: #fff;
    display: flex;
    flex-direction: column;
    min-height: 0;
  }
  .xml-head {
    font-size: 11px;
    font-weight: 600;
    color: var(--muted);
    padding: 6px 10px;
    border-bottom: 1px solid var(--chrome-line);
    background: var(--chrome);
  }
  .loading { text-align: center; color: var(--muted); }

  .statusbar {
    display: flex;
    align-items: center;
    gap: 16px;
    height: 26px;
    padding: 0 14px;
    border-top: 1px solid var(--status-line);
    background: var(--status-bg);
    font-size: 13px;
    color: var(--text);
  }
  .status-msg { color: #a4262c; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
  .spacer { flex: 1; }
  .zoom { display: flex; align-items: center; gap: 6px; }
  .zoom button { width: 18px; height: 18px; border: none; background: none; cursor: pointer; font-size: 14px; line-height: 1; border-radius: 3px; }
  .zoom button:hover { background: var(--hover); }
  .zoom input[type='range'] { width: 130px; accent-color: #8a8a8a; }
  .zoom-val { min-width: 40px; text-align: right; }

  /* Show/Hide ¶: a pilcrow at the end of every paragraph, as Word draws it. */
  .app.marks :global(.wk-p)::after { content: '¶'; color: #7a7a7a; font-weight: normal; font-style: normal; }
</style>
