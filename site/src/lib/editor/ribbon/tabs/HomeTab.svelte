<script lang="ts">
  import {
    commands,
    highlightCss,
    runAtPath,
    paragraphAt,
    orderSelection,
    paragraphsInRange,
    runsInRange,
    createStyleResolver,
    type ResolvedRunFormat,
  } from '@office-kit/docx-editor';
  import { getParagraphStyle, type HighlightColor, type RunFormatting } from '@office-kit/docx';
  import RibbonIcon from '../../RibbonIcon.svelte';
  import Button from '../Button.svelte';
  import SplitButton from '../SplitButton.svelte';
  import Group from '../Group.svelte';
  import { getSession } from '../../session.svelte';
  import { t, highlightName, type MessageKey } from '../../i18n/index.svelte';

  const session = getSession();

  /**
   * What a ribbon control shows for the current selection. Nothing here is
   * ever written back to the document; only an explicit user change runs a
   * command.
   */
  type FieldState<T> = { kind: 'none' } | { kind: 'mixed' } | { kind: 'inherited' } | { kind: 'value'; value: T };
  const NONE = { kind: 'none' } as const;

  // What the Font Color / Highlight buttons apply on a plain click: the last
  // color picked from their menus (Word's split-button behavior).
  let fontColorPen = $state('C00000');
  let highlightPen = $state<HighlightColor>('yellow');
  let fontInput = $state<HTMLInputElement | null>(null);
  let sizeInput = $state<HTMLInputElement | null>(null);

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

  /** Collapse per-run (or per-paragraph) values into one control state. */
  function fieldOf<T>(values: ReadonlyArray<T | undefined>): FieldState<T> {
    if (values.length === 0) return NONE;
    const first = values[0];
    if (values.some((v) => v !== first)) return { kind: 'mixed' };
    return first === undefined ? { kind: 'inherited' } : { kind: 'value', value: first };
  }

  /**
   * The selection's formatting as Word reports it: the effective value (styles
   * and document defaults included), so a caret in a Heading 1 reads its
   * style's font and size. A caret reads the run it is in (or the paragraph
   * mark in an empty paragraph); a range reads every run it covers, so a
   * differently formatted run in the middle makes a field "mixed".
   */
  const selectionFormat = $derived.by(() => {
    const model = session.tick >= 0 ? session.model : null;
    const sel = model?.selection;
    if (!model || !sel) return { formats: [] as ResolvedRunFormat[], styles: [] as Array<string | undefined> };
    const doc = model.doc;
    const styles = createStyleResolver(doc);
    const ordered = orderSelection(sel);
    // paragraphsInRange takes every cell of a table in range, so a caret uses
    // only its own paragraph.
    const caretPara = ordered.collapsed ? paragraphAt(doc, sel.focus) : undefined;
    const paras = ordered.collapsed ? (caretPara ? [caretPara] : []) : paragraphsInRange(doc, ordered);
    let formats: ResolvedRunFormat[];
    if (caretPara) {
      formats = [styles.run(caretPara, runAtPath(doc, sel.focus))];
    } else {
      const owner = new Map(paras.flatMap((p) => p.children.filter((c) => c.kind === 'run').map((r) => [r, p] as const)));
      formats = runsInRange(doc, ordered).flatMap((run) => {
        const para = owner.get(run);
        return para ? [styles.run(para, run)] : [];
      });
    }
    return { formats, styles: paras.map((p) => getParagraphStyle(p)) };
  });

  const fontState = $derived(
    fieldOf(selectionFormat.formats.map((f) => (f.font === undefined ? undefined : `${f.font}${f.fontRole ? ` ${t(`font.${f.fontRole}`)}` : ''}`))),
  );
  const sizeState = $derived(fieldOf(selectionFormat.formats.map((f) => (f.sizeHalfPoints === undefined ? undefined : f.sizeHalfPoints / 2))));
  const vertAlignState = $derived(fieldOf(selectionFormat.formats.map((f) => f.vertAlign ?? 'baseline')));
  const styleState = $derived(fieldOf(selectionFormat.styles));
  const lineSpacing = $derived(session.tick >= 0 && session.model ? commands.lineSpacingOf(session.model) : undefined);

  /** Text shown in the Font / Size boxes: Word leaves them blank when mixed. */
  function fieldText<T>(state: FieldState<T>): string {
    return state.kind === 'value' ? String(state.value) : '';
  }

  function fieldPlaceholder(state: FieldState<unknown>): string {
    return state.kind === 'inherited' ? t('state.inherited') : '';
  }

  function onFontChange(e: Event): void {
    const input = e.currentTarget as HTMLInputElement;
    // "Calibri (Body)" names the theme font; typing it back means that font.
    const font = input.value.replace(/\s*\([^)]*\)$/, '').trim();
    if (font) session.apply(commands.setFontCommand, { font });
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
    session.apply(commands.setFontSizeCommand, { points });
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
    session.apply(commands.setFontSizeCommand, { points: listed ?? fallback });
  }

  function toggleVertAlign(val: 'superscript' | 'subscript'): void {
    const on = vertAlignState.kind === 'value' && vertAlignState.value === val;
    session.apply(commands.setVertAlignCommand, { val: on ? undefined : val });
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
      if (text) session.apply(commands.insertTextCommand, { text });
    } catch (err) {
      session.status = `Paste failed: ${(err as Error).message}`;
    }
  }

  function hexCss(hex: string): string {
    return HEX6.test(hex) ? `#${hex}` : 'transparent';
  }

  const hasSelection = $derived(session.tick >= 0 && !!session.model?.selection);
</script>

<Group label={t('group.clipboard')}>
  <button class="big paste" onclick={paste} title={t('tip.paste')}><RibbonIcon name="paste" size={32} /><span>{t('tip.paste')}</span></button>
  <div class="col">
    <Button icon="cut" tip={t('tip.cut')} onclick={() => clipboard('cut')} disabled={!hasSelection} />
    <Button icon="copy" tip={t('tip.copy')} onclick={() => clipboard('copy')} disabled={!hasSelection} />
  </div>
</Group>

<Group label={t('group.font')}>
  <div class="rows">
    <div class="row">
      <span class="combo-wrap">
        <input
          bind:this={fontInput}
          class="combo font"
          list="wk-fonts"
          value={fieldText(fontState)}
          placeholder={fieldPlaceholder(fontState)}
          disabled={!session.enabled(commands.setFontCommand)}
          onchange={onFontChange}
          title={t('tip.fontName')}
          aria-label={t('tip.fontName')}
          spellcheck="false"
        />
        <button class="combo-arrow" onclick={() => fontInput?.showPicker()} disabled={!session.enabled(commands.setFontCommand)} aria-label={t('tip.fontName')} tabindex="-1"><RibbonIcon name="chevronDown" size={10} /></button>
      </span>
      <datalist id="wk-fonts">{#each FONT_CHOICES as font (font)}<option value={font}></option>{/each}</datalist>
      <span class="combo-wrap">
        <input
          bind:this={sizeInput}
          class="combo size"
          list="wk-sizes"
          inputmode="decimal"
          value={fieldText(sizeState)}
          disabled={!session.enabled(commands.setFontSizeCommand)}
          onchange={onFontSizeChange}
          title={t('tip.fontSize')}
          aria-label={t('tip.fontSize')}
        />
        <button class="combo-arrow" onclick={() => sizeInput?.showPicker()} disabled={!session.enabled(commands.setFontSizeCommand)} aria-label={t('tip.fontSize')} tabindex="-1"><RibbonIcon name="chevronDown" size={10} /></button>
      </span>
      <datalist id="wk-sizes">{#each FONT_SIZES as size (size)}<option value={size}></option>{/each}</datalist>
      <Button glyph="A<sup>^</sup>" glyphClass="grow" tip={t('tip.grow')} onclick={() => stepFontSize(1)} disabled={sizeState.kind !== 'value'} />
      <Button glyph="A<sup>ˇ</sup>" glyphClass="grow" tip={t('tip.shrink')} onclick={() => stepFontSize(-1)} disabled={sizeState.kind !== 'value'} />
      <span class="sep"></span>
      <Button glyph={'A<i class="eraser"></i>'} glyphClass="clear" tip={t('tip.clearAll')} onclick={() => session.apply(commands.clearFormattingCommand, undefined)} disabled={!session.enabled(commands.clearFormattingCommand)} />
    </div>
    <div class="row">
      <Button glyph="B" glyphClass="b" tip={`${t('tip.bold')} (⌘B)`} onclick={() => session.apply(commands.toggleBoldCommand, undefined)} on={session.active(commands.toggleBoldCommand)} disabled={!session.enabled(commands.toggleBoldCommand)} />
      <Button glyph="I" glyphClass="i" tip={`${t('tip.italic')} (⌘I)`} onclick={() => session.apply(commands.toggleItalicCommand, undefined)} on={session.active(commands.toggleItalicCommand)} disabled={!session.enabled(commands.toggleItalicCommand)} />
      <SplitButton id="home.underline" tip={`${t('tip.underline')} (⌘U)`} onclick={() => session.apply(commands.toggleUnderlineCommand, undefined)} on={session.active(commands.toggleUnderlineCommand)} disabled={!session.enabled(commands.toggleUnderlineCommand)}>
        {#snippet face()}<span class="glyph u">U</span>{/snippet}
        {#snippet menu()}
          {#each UNDERLINES as u (u.style)}
            <button class="mi" role="menuitem" onclick={() => session.apply(commands.setUnderlineStyleCommand, { style: u.style })}><span class="ul-sample ul-{u.style}">{t(u.key)}</span></button>
          {/each}
        {/snippet}
      </SplitButton>
      <Button glyph="ab" glyphClass="strike" tip={t('tip.strike')} onclick={() => session.apply(commands.toggleStrikeCommand, undefined)} on={session.active(commands.toggleStrikeCommand)} disabled={!session.enabled(commands.toggleStrikeCommand)} />
      <Button glyph="x<sub>2</sub>" glyphClass="script" tip={t('tip.subscript')} onclick={() => toggleVertAlign('subscript')} on={vertAlignState.kind === 'value' && vertAlignState.value === 'subscript'} disabled={!session.enabled(commands.setVertAlignCommand)} />
      <Button glyph="x<sup>2</sup>" glyphClass="script" tip={t('tip.superscript')} onclick={() => toggleVertAlign('superscript')} on={vertAlignState.kind === 'value' && vertAlignState.value === 'superscript'} disabled={!session.enabled(commands.setVertAlignCommand)} />
      <span class="sep"></span>
      <SplitButton id="home.highlight" tip={t('tip.highlight')} onclick={() => session.apply(commands.setHighlightCommand, { color: highlightPen })} disabled={!session.enabled(commands.setHighlightCommand)}>
        {#snippet face()}<span class="swatch-icon"><svg viewBox="0 0 20 20" width="20" height="16" fill="none" stroke="currentColor" stroke-width="1.2" stroke-linejoin="round" aria-hidden="true"><path d="M5 12.5 12.5 5l2.5 2.5L7.5 15H5zM11 6.5l2.5 2.5" /></svg><i style="background:{highlightCss(highlightPen)}"></i></span>{/snippet}
        {#snippet menu()}
          <div class="grid five">
            {#each HIGHLIGHT_MENU as color (color)}
              <button class="chip" role="menuitem" style="background:{highlightCss(color)}" title={highlightName(color)} aria-label={highlightName(color)} onclick={() => { highlightPen = color; session.apply(commands.setHighlightCommand, { color }); }}></button>
            {/each}
          </div>
          <button class="mi" role="menuitem" onclick={() => session.apply(commands.setHighlightCommand, { color: 'none' })}>{t('highlight.none')}</button>
        {/snippet}
      </SplitButton>
      <SplitButton id="home.fontColor" tip={t('tip.fontColor')} onclick={() => session.apply(commands.setColorCommand, { color: fontColorPen })} disabled={!session.enabled(commands.setColorCommand)}>
        {#snippet face()}<span class="swatch-icon"><span class="a">A</span><i style="background:{fontColorPen === 'auto' ? '#000' : hexCss(fontColorPen)}"></i></span>{/snippet}
        {#snippet menu()}
          <button class="mi" role="menuitem" onclick={() => { fontColorPen = 'auto'; session.apply(commands.setColorCommand, { color: 'auto' }); }}><i class="auto-chip"></i>{t('color.auto')}</button>
          <div class="menu-head">{t('color.standard')}</div>
          <div class="grid ten">
            {#each STANDARD_COLORS as hex (hex)}
              <button class="chip" role="menuitem" style="background:#{hex}" title={`#${hex}`} aria-label={`#${hex}`} onclick={() => { fontColorPen = hex; session.apply(commands.setColorCommand, { color: hex }); }}></button>
            {/each}
          </div>
          <label class="mi">
            {t('color.more')}
            <input
              type="color"
              value={HEX6.test(fontColorPen) ? `#${fontColorPen}` : '#000000'}
              onchange={(e) => { fontColorPen = (e.currentTarget as HTMLInputElement).value.slice(1).toUpperCase(); session.apply(commands.setColorCommand, { color: fontColorPen }); }}
              hidden
            />
          </label>
        {/snippet}
      </SplitButton>
    </div>
  </div>
</Group>

<Group label={t('group.paragraph')}>
  <div class="rows">
    <div class="row">
      <Button icon="bullets" tip={t('tip.bullets')} onclick={() => session.apply(commands.applyListCommand, { kind: 'bullet' })} disabled={!session.enabled(commands.applyListCommand)} />
      <Button icon="numbering" tip={t('tip.numbering')} onclick={() => session.apply(commands.applyListCommand, { kind: 'numbered' })} disabled={!session.enabled(commands.applyListCommand)} />
      <span class="sep"></span>
      <Button icon="indentLess" tip={t('tip.indentLess')} onclick={() => session.apply(commands.indentStepCommand, { direction: 'decrease' })} disabled={!session.enabled(commands.indentStepCommand)} />
      <Button icon="indentMore" tip={t('tip.indentMore')} onclick={() => session.apply(commands.indentStepCommand, { direction: 'increase' })} disabled={!session.enabled(commands.indentStepCommand)} />
      <span class="sep"></span>
      <Button icon="pilcrow" tip={`${t('tip.marks')} (⌘8)`} onclick={() => (session.showMarks = !session.showMarks)} on={session.showMarks} />
    </div>
    <div class="row">
      <Button icon="alignLeft" tip={`${t('tip.alignLeft')} (⌘L)`} onclick={() => session.apply(commands.alignLeftCommand, undefined)} on={session.active(commands.alignLeftCommand)} disabled={!session.enabled(commands.alignLeftCommand)} />
      <Button icon="alignCenter" tip={`${t('tip.center')} (⌘E)`} onclick={() => session.apply(commands.alignCenterCommand, undefined)} on={session.active(commands.alignCenterCommand)} disabled={!session.enabled(commands.alignCenterCommand)} />
      <Button icon="alignRight" tip={`${t('tip.alignRight')} (⌘R)`} onclick={() => session.apply(commands.alignRightCommand, undefined)} on={session.active(commands.alignRightCommand)} disabled={!session.enabled(commands.alignRightCommand)} />
      <Button icon="alignJustify" tip={`${t('tip.justify')} (⌘J)`} onclick={() => session.apply(commands.alignJustifyCommand, undefined)} on={session.active(commands.alignJustifyCommand)} disabled={!session.enabled(commands.alignJustifyCommand)} />
      <span class="sep"></span>
      <SplitButton id="home.lineSpacing" icon="lineSpacing" tip={t('tip.lineSpacing')} disabled={!session.enabled(commands.setLineSpacingCommand)}>
        {#snippet menu()}
          {#each LINE_SPACINGS as multiple (multiple)}
            <button class="mi check" class:checked={lineSpacing === multiple} role="menuitemradio" aria-checked={lineSpacing === multiple} onclick={() => session.apply(commands.setLineSpacingCommand, { multiple })}>{multiple.toFixed(multiple === 1.15 ? 2 : 1)}</button>
          {/each}
        {/snippet}
      </SplitButton>
    </div>
  </div>
</Group>

<Group label={t('group.styles')} class="gallery-group">
  <div class="gallery" role="listbox" aria-label={t('group.styles')}>
    {#each STYLE_GALLERY as s (s.key)}
      {@const selected = s.id === undefined ? styleState.kind === 'inherited' : styleState.kind === 'value' && styleState.value === s.id}
      <button
        class="tile"
        class:selected
        role="option"
        aria-selected={selected}
        disabled={!session.enabled(commands.setParagraphStyleCommand)}
        onclick={() => session.apply(commands.setParagraphStyleCommand, { styleId: s.id })}
        title={t(s.key)}
      >
        <span class="sample sample-{s.id ?? 'Normal'}">AaBbCcDdEe</span>
        <span class="tile-name">{t(s.key)}</span>
      </button>
    {/each}
  </div>
</Group>
