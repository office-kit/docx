<script lang="ts">
  /**
   * Word for Mac's Home tab: Clipboard, Font, Paragraph and Styles. The
   * dialogs these open live in `home/HomeDialogs.svelte` (mounted with the
   * page, so ⌘D and ⌘⌥M work from any tab).
   */
  import {
    commands,
    highlightCss,
    readThemeFonts,
    BULLET_PRESETS,
    NUMBERING_PRESETS,
    MULTILEVEL_PRESETS,
    type CaseMode,
    type ListPreset,
    type RunToggle,
  } from '@office-kit/docx-editor';
  import { UNDERLINE_STYLES, type ColorValue, type HighlightColor, type UnderlineStyle } from '@office-kit/docx';
  import RibbonIcon from '../../RibbonIcon.svelte';
  import Button from '../Button.svelte';
  import SplitButton from '../SplitButton.svelte';
  import Group from '../Group.svelte';
  import ColorMenu from './home/ColorMenu.svelte';
  import { getSession } from '../../session.svelte';
  import { t, highlightName, type MessageKey } from '../../i18n/index.svelte';
  import { common, homeState, selectionFormats } from './home/state.svelte';
  import { multilevelPreview, previewLabels } from './home/list-preview';
  import { galleryStyles, stylePreviews } from './home/style-preview';

  const session = getSession();
  const home = homeState(session);

  // Fonts Word lists under All Fonts on a Mac (the browser cannot enumerate
  // installed fonts without a permission prompt).
  const ALL_FONTS = [
    'Aptos', 'Aptos Display', 'Arial', 'Arial Black', 'Avenir Next', 'Calibri', 'Calibri Light', 'Cambria', 'Candara',
    'Century Gothic', 'Comic Sans MS', 'Consolas', 'Constantia', 'Corbel', 'Courier New', 'Franklin Gothic Book',
    'Garamond', 'Georgia', 'Gill Sans', 'Helvetica', 'Helvetica Neue', 'Hiragino Kaku Gothic ProN', 'Hiragino Mincho ProN',
    'Impact', 'Lucida Grande', 'Meiryo', 'MS Gothic', 'MS Mincho', 'Palatino', 'Segoe UI', 'Symbol', 'Tahoma',
    'Times New Roman', 'Trebuchet MS', 'Verdana', 'Wingdings', 'Yu Gothic', 'Yu Mincho',
  ];
  const HIGHLIGHT_MENU: HighlightColor[] = [
    'yellow', 'green', 'cyan', 'magenta', 'blue', 'red', 'darkBlue', 'darkCyan',
    'darkGreen', 'darkMagenta', 'darkRed', 'darkYellow', 'darkGray', 'lightGray', 'black',
  ];
  const LINE_SPACINGS = [1, 1.15, 1.5, 2, 2.5, 3];
  const CASES: ReadonlyArray<{ mode: CaseMode; key: MessageKey }> = [
    { mode: 'sentence', key: 'home.case.sentence' },
    { mode: 'lower', key: 'home.case.lower' },
    { mode: 'upper', key: 'home.case.upper' },
    { mode: 'title', key: 'home.case.title' },
    { mode: 'toggle', key: 'home.case.toggle' },
    { mode: 'halfWidth', key: 'home.case.halfWidth' },
    { mode: 'fullWidth', key: 'home.case.fullWidth' },
    { mode: 'katakana', key: 'home.case.katakana' },
    { mode: 'hiragana', key: 'home.case.hiragana' },
  ];
  const EFFECTS = [
    { effect: 'outline', key: 'home.effect.outline' },
    { effect: 'shadow', key: 'home.effect.shadow' },
    { effect: 'emboss', key: 'home.effect.emboss' },
    { effect: 'imprint', key: 'home.effect.imprint' },
    { effect: 'smallCaps', key: 'home.effect.smallCaps' },
    { effect: 'caps', key: 'home.effect.caps' },
    { effect: 'dstrike', key: 'home.effect.dstrike' },
  ] as const satisfies ReadonlyArray<{ effect: RunToggle; key: MessageKey }>;
  const BORDER_PRESETS = [
    { preset: 'bottom', key: 'home.border.bottom' },
    { preset: 'top', key: 'home.border.top' },
    { preset: 'left', key: 'home.border.left' },
    { preset: 'right', key: 'home.border.right' },
    { preset: 'none', key: 'home.border.none' },
    { preset: 'all', key: 'home.border.all' },
    { preset: 'outside', key: 'home.border.outside' },
    { preset: 'inside', key: 'home.border.inside' },
    { preset: 'insideHorizontal', key: 'home.border.insideHorizontal' },
  ] as const satisfies ReadonlyArray<{ preset: commands.BorderPreset; key: MessageKey }>;
  const CHARACTER_SCALES = [200, 150, 100, 90, 80, 66, 50, 33];
  const LIST_LEVELS = 9;
  // Tiles the Styles gallery shows in the ribbon before the expand arrow.
  const GALLERY_TILES = 5;

  const model = $derived(session.tick >= 0 ? session.model : null);
  const formats = $derived(model ? selectionFormats(model) : { runs: [], paragraphs: [], paragraphStyles: [] });
  const fontName = $derived(common(formats.runs.map((f) => f.font)));
  const fontRole = $derived(common(formats.runs.map((f) => f.fontRole)));
  const size = $derived(common(formats.runs.map((f) => (f.sizeHalfPoints === undefined ? undefined : f.sizeHalfPoints / 2))));
  const vertAlign = $derived(common(formats.runs.map((f) => f.vertAlign ?? 'baseline')));
  const underline = $derived(common(formats.runs.map((f) => f.underline ?? 'none')));
  const lineSpacing = $derived(model ? commands.lineSpacingOf(model) : undefined);
  const listKind = $derived(model ? commands.selectionListKind(model) : undefined);
  const themeFonts = $derived(session.version >= 0 && session.model ? readThemeFonts(session.model.doc) : {});
  const gallery = $derived(session.version >= 0 && session.model ? galleryStyles(session.model) : []);
  const previews = $derived(session.version >= 0 && session.model ? stylePreviews(session.model) : undefined);
  const currentStyle = $derived.by(() => {
    const ids = formats.paragraphStyles;
    if (ids.length === 0) return undefined;
    const id = common(ids);
    // No pStyle is the default paragraph style.
    return id ?? (ids.every((x) => x === undefined) && model ? commands.defaultParagraphStyleId(model) : undefined);
  });
  const hasSelection = $derived(!!model?.selection);
  const caretInList = $derived(!!model && commands.restartNumberingCommand.isEnabled?.(model));

  function effectOn(effect: RunToggle): boolean {
    return formats.runs.length > 0 && formats.runs.every((f) => f.toggles.has(effect));
  }

  function fontLabel(): string {
    if (fontName === undefined) return '';
    return fontRole ? `${fontName} ${t(`font.${fontRole}`)}` : fontName;
  }

  function setFont(font: { font: string } | { theme: 'major' | 'minor' }): void {
    if ('font' in font) home.useFont(font.font);
    session.apply(commands.setFontCommand, font);
  }

  function onFontChange(e: Event): void {
    const input = e.currentTarget as HTMLInputElement;
    // "Aptos (Body)" names the theme font; typing it back means that font.
    const font = input.value.replace(/\s*\([^)]*\)$/, '').trim();
    if (font) setFont({ font });
    else input.value = fontLabel();
  }

  function onFontSizeChange(e: Event): void {
    const input = e.currentTarget as HTMLInputElement;
    const points = Number(input.value);
    // Word accepts 1–1638 pt in half-point steps; anything else is refused
    // without touching the document.
    if (input.value === '' || !Number.isFinite(points) || points < 1 || points > 1638) {
      input.value = size === undefined ? '' : String(size);
      return;
    }
    session.apply(commands.setFontSizeCommand, { points });
  }

  function toggleVertAlign(val: 'superscript' | 'subscript'): void {
    session.apply(commands.setVertAlignCommand, { val: vertAlign === val ? undefined : val });
  }

  function toggleEffect(effect: RunToggle): void {
    session.apply(commands.fontFormatCommand, { effects: { [effect]: !effectOn(effect) } });
  }

  function setUnderline(style: UnderlineStyle): void {
    home.underlinePen = style;
    session.apply(commands.setUnderlineStyleCommand, { style });
  }

  function setFontColor(color: ColorValue | undefined): void {
    const value = color ?? { rgb: 'auto' };
    home.fontColorPen = value;
    session.apply(commands.setColorCommand, {
      color: value.rgb,
      ...(value.themeColor ? { themeColor: value.themeColor } : {}),
      ...(value.themeTint === undefined ? {} : { themeTint: value.themeTint }),
      ...(value.themeShade === undefined ? {} : { themeShade: value.themeShade }),
    });
  }

  function setShading(fill: ColorValue | undefined): void {
    home.shadingPen = fill;
    session.apply(commands.paragraphShadingColorCommand, { fill });
  }

  function applyList(preset: ListPreset): void {
    session.apply(commands.applyListPresetCommand, { levels: preset.levels });
  }

  function chip(color: ColorValue | undefined): string {
    if (!color) return 'transparent';
    return color.rgb === 'auto' ? '#000' : `#${color.rgb}`;
  }

  /** Cut / Copy go through the browser so the canvas' own clipboard handlers run. */
  function clipboard(action: 'cut' | 'copy'): void {
    document.execCommand(action);
  }

  async function paste(): Promise<void> {
    // Reading the clipboard needs the browser's permission; a refusal is shown
    // in the status bar and ⌘V keeps working.
    try {
      const text = await navigator.clipboard.readText();
      if (text) session.apply(commands.insertTextCommand, { text });
    } catch (err) {
      session.status = `${t('tip.paste')}: ${(err as Error).message}`;
    }
  }

  // Format Painter: a click arms it for one paste, a double-click keeps it on
  // until Esc or another click (Word's behaviour).
  let painterClick: ReturnType<typeof setTimeout> | undefined;
  const DOUBLE_CLICK_MS = 300;

  function formatPainter(): void {
    if (painterClick) {
      clearTimeout(painterClick);
      painterClick = undefined;
      if (home.painter) home.painter = { ...home.painter, sticky: true };
      return;
    }
    if (home.painter) {
      home.painter = null;
      return;
    }
    const format = model && commands.copyFormat(model);
    if (format) home.painter = { format, sticky: false };
    painterClick = setTimeout(() => (painterClick = undefined), DOUBLE_CLICK_MS);
  }

  $effect(() => {
    const armed = home.painter;
    if (!armed) return;
    const app = document.querySelector('.wk-app');
    app?.classList.add('painting');
    // Paint when the selection gesture in the page ends.
    const onUp = (e: PointerEvent): void => {
      if (!(e.target as Element).closest('.surface')) return;
      queueMicrotask(() => {
        session.apply(commands.pasteFormatCommand, armed.format);
        if (!armed.sticky) home.painter = null;
      });
    };
    const onKey = (e: KeyboardEvent): void => {
      if (e.key === 'Escape') home.painter = null;
    };
    window.addEventListener('pointerup', onUp);
    window.addEventListener('keydown', onKey);
    return () => {
      app?.classList.remove('painting');
      window.removeEventListener('pointerup', onUp);
      window.removeEventListener('keydown', onKey);
    };
  });
</script>

<Group label={t('group.clipboard')}>
  <SplitButton id="home.paste" size="large" icon="paste" label={t('tip.paste')} tip={t('tip.paste')} onclick={paste}>
    {#snippet menu()}
      <button class="mi" role="menuitem" onclick={paste}>{t('home.paste.textOnly')}</button>
    {/snippet}
  </SplitButton>
  <div class="col">
    <Button icon="cut" tip={`${t('tip.cut')} (⌘X)`} onclick={() => clipboard('cut')} disabled={!hasSelection} />
    <Button icon="copy" tip={`${t('tip.copy')} (⌘C)`} onclick={() => clipboard('copy')} disabled={!hasSelection} />
    <Button icon="formatPainter" tip={`${t('home.formatPainter')} (⌘⇧C)`} onclick={formatPainter} on={!!home.painter} disabled={!hasSelection} />
  </div>
</Group>

<Group label={t('group.font')}>
  <div class="rows">
    <div class="row">
      <span class="split combo-wrap">
        <input
          class="combo font"
          value={fontLabel()}
          disabled={!session.enabled(commands.setFontCommand)}
          onchange={onFontChange}
          title={t('tip.fontName')}
          aria-label={t('tip.fontName')}
          spellcheck="false"
        />
        <button class="combo-arrow" onclick={() => session.toggleMenu('home.fontList')} disabled={!session.enabled(commands.setFontCommand)} aria-label={t('tip.fontName')} aria-expanded={session.openMenu === 'home.fontList'} tabindex="-1"><RibbonIcon name="chevronDown" size={10} /></button>
        {#if session.openMenu === 'home.fontList'}
          <div class="menu font-list" role="menu">
            <div class="menu-head">{t('home.font.theme')}</div>
            <button class="mi" role="menuitem" onclick={() => setFont({ theme: 'major' })}><span style="font-family:'{themeFonts.major ?? 'Aptos Display'}'">{themeFonts.major ?? 'Aptos Display'}</span><span class="mi-note">{t('font.headings')}</span></button>
            <button class="mi" role="menuitem" onclick={() => setFont({ theme: 'minor' })}><span style="font-family:'{themeFonts.minor ?? 'Aptos'}'">{themeFonts.minor ?? 'Aptos'}</span><span class="mi-note">{t('font.body')}</span></button>
            {#if home.recentFonts.length > 0}
              <div class="menu-head">{t('home.font.recent')}</div>
              {#each home.recentFonts as font (font)}
                <button class="mi" role="menuitem" style="font-family:'{font}'" onclick={() => setFont({ font })}>{font}</button>
              {/each}
            {/if}
            <div class="menu-head">{t('home.font.all')}</div>
            {#each ALL_FONTS as font (font)}
              <button class="mi" role="menuitem" style="font-family:'{font}'" onclick={() => setFont({ font })}>{font}</button>
            {/each}
          </div>
        {/if}
      </span>
      <span class="split combo-wrap">
        <input
          class="combo size"
          inputmode="decimal"
          value={size === undefined ? '' : String(size)}
          disabled={!session.enabled(commands.setFontSizeCommand)}
          onchange={onFontSizeChange}
          title={t('tip.fontSize')}
          aria-label={t('tip.fontSize')}
        />
        <button class="combo-arrow" onclick={() => session.toggleMenu('home.sizeList')} disabled={!session.enabled(commands.setFontSizeCommand)} aria-label={t('tip.fontSize')} aria-expanded={session.openMenu === 'home.sizeList'} tabindex="-1"><RibbonIcon name="chevronDown" size={10} /></button>
        {#if session.openMenu === 'home.sizeList'}
          <div class="menu size-list" role="menu">
            {#each commands.FONT_SIZES as points (points)}
              <button class="mi check" class:checked={size === points} role="menuitemradio" aria-checked={size === points} onclick={() => session.apply(commands.setFontSizeCommand, { points })}>{points}</button>
            {/each}
          </div>
        {/if}
      </span>
      <Button glyph="A<sup>^</sup>" glyphClass="grow" tip={`${t('tip.grow')} (⌘⇧>)`} onclick={() => session.apply(commands.growFontCommand, { direction: 1 })} disabled={!session.enabled(commands.growFontCommand)} />
      <Button glyph="A<sup>ˇ</sup>" glyphClass="grow" tip={`${t('tip.shrink')} (⌘⇧<)`} onclick={() => session.apply(commands.growFontCommand, { direction: -1 })} disabled={!session.enabled(commands.growFontCommand)} />
      <span class="sep"></span>
      <SplitButton id="home.changeCase" icon="changeCase" tip={t('home.changeCase')} disabled={!session.enabled(commands.changeCaseCommand)}>
        {#snippet menu()}
          {#each CASES as c (c.mode)}
            <button class="mi" role="menuitem" onclick={() => session.apply(commands.changeCaseCommand, { mode: c.mode })}>{t(c.key)}</button>
          {/each}
        {/snippet}
      </SplitButton>
      <Button icon="clearFormatting" tip={t('tip.clearAll')} onclick={() => session.apply(commands.clearFormattingCommand, undefined)} disabled={!session.enabled(commands.clearFormattingCommand)} />
      <Button icon="phoneticGuide" tip={t('home.phoneticGuide')} onclick={() => session.openDialog('home.ruby')} disabled={!session.enabled(commands.phoneticGuideCommand)} />
      <Button icon="characterBorder" tip={t('home.characterBorder')} onclick={() => session.apply(commands.toggleCharacterBorderCommand, undefined)} on={session.active(commands.toggleCharacterBorderCommand)} disabled={!session.enabled(commands.toggleCharacterBorderCommand)} />
    </div>
    <div class="row">
      <Button glyph="B" glyphClass="b" tip={`${t('tip.bold')} (⌘B)`} onclick={() => session.apply(commands.toggleBoldCommand, undefined)} on={session.active(commands.toggleBoldCommand)} disabled={!session.enabled(commands.toggleBoldCommand)} />
      <Button glyph="I" glyphClass="i" tip={`${t('tip.italic')} (⌘I)`} onclick={() => session.apply(commands.toggleItalicCommand, undefined)} on={session.active(commands.toggleItalicCommand)} disabled={!session.enabled(commands.toggleItalicCommand)} />
      <SplitButton id="home.underline" tip={`${t('tip.underline')} (⌘U)`} onclick={() => (underline !== undefined && underline !== 'none' ? session.apply(commands.toggleUnderlineCommand, undefined) : setUnderline(home.underlinePen))} on={session.active(commands.toggleUnderlineCommand)} disabled={!session.enabled(commands.toggleUnderlineCommand)}>
        {#snippet face()}<span class="glyph u">U</span>{/snippet}
        {#snippet menu()}
          <div class="underline-list">
            {#each UNDERLINE_STYLES.filter((s) => s !== 'none' && s !== 'words') as style (style)}
              <button class="mi check" class:checked={underline === style} role="menuitemradio" aria-checked={underline === style} title={style} onclick={() => setUnderline(style)}><span class="ul-line ul-{style}"></span></button>
            {/each}
          </div>
          <button class="mi check" class:checked={underline === 'words'} role="menuitemradio" aria-checked={underline === 'words'} onclick={() => setUnderline('words')}>{t('home.underline.words')}</button>
          <hr />
          <button class="mi" role="menuitem" onclick={() => session.openDialog('home.font')}>{t('home.underline.more')}</button>
          <div class="menu-head">{t('home.underline.color')}</div>
          <ColorMenu none="auto" onpick={(color) => session.apply(commands.setUnderlineColorCommand, { color })} />
        {/snippet}
      </SplitButton>
      <Button glyph="ab" glyphClass="strike" tip={t('tip.strike')} onclick={() => session.apply(commands.toggleStrikeCommand, undefined)} on={session.active(commands.toggleStrikeCommand)} disabled={!session.enabled(commands.toggleStrikeCommand)} />
      <Button glyph="x<sub>2</sub>" glyphClass="script" tip={`${t('tip.subscript')} (⌘=)`} onclick={() => toggleVertAlign('subscript')} on={vertAlign === 'subscript'} disabled={!session.enabled(commands.setVertAlignCommand)} />
      <Button glyph="x<sup>2</sup>" glyphClass="script" tip={`${t('tip.superscript')} (⌘⇧=)`} onclick={() => toggleVertAlign('superscript')} on={vertAlign === 'superscript'} disabled={!session.enabled(commands.setVertAlignCommand)} />
      <span class="sep"></span>
      <SplitButton id="home.textEffects" icon="textEffects" tip={t('home.textEffects')} disabled={!session.enabled(commands.fontFormatCommand)}>
        {#snippet menu()}
          {#each EFFECTS as e (e.effect)}
            <button class="mi check" class:checked={effectOn(e.effect)} role="menuitemcheckbox" aria-checked={effectOn(e.effect)} onclick={() => toggleEffect(e.effect)}>{t(e.key)}</button>
          {/each}
        {/snippet}
      </SplitButton>
      <SplitButton id="home.highlight" tip={t('tip.highlight')} onclick={() => session.apply(commands.setHighlightCommand, { color: home.highlightPen })} disabled={!session.enabled(commands.setHighlightCommand)}>
        {#snippet face()}<span class="swatch-icon"><svg viewBox="0 0 20 20" width="20" height="16" fill="none" stroke="currentColor" stroke-width="1.2" stroke-linejoin="round" aria-hidden="true"><path d="M5 12.5 12.5 5l2.5 2.5L7.5 15H5zM11 6.5l2.5 2.5" /></svg><i style="background:{highlightCss(home.highlightPen)}"></i></span>{/snippet}
        {#snippet menu()}
          <div class="grid five">
            {#each HIGHLIGHT_MENU as color (color)}
              <button class="chip" role="menuitem" style="background:{highlightCss(color)}" title={highlightName(color)} aria-label={highlightName(color)} onclick={() => { home.highlightPen = color; session.apply(commands.setHighlightCommand, { color }); }}></button>
            {/each}
          </div>
          <button class="mi" role="menuitem" onclick={() => session.apply(commands.setHighlightCommand, { color: 'none' })}>{t('highlight.none')}</button>
        {/snippet}
      </SplitButton>
      <SplitButton id="home.fontColor" tip={t('tip.fontColor')} onclick={() => setFontColor(home.fontColorPen)} disabled={!session.enabled(commands.setColorCommand)}>
        {#snippet face()}<span class="swatch-icon"><span class="a">A</span><i style="background:{chip(home.fontColorPen)}"></i></span>{/snippet}
        {#snippet menu()}<ColorMenu none="auto" onpick={setFontColor} />{/snippet}
      </SplitButton>
      <Button icon="characterShading" tip={t('home.characterShading')} onclick={() => session.apply(commands.toggleCharacterShadingCommand, undefined)} on={session.active(commands.toggleCharacterShadingCommand)} disabled={!session.enabled(commands.toggleCharacterShadingCommand)} />
      <Button icon="encloseCharacters" tip={t('home.encloseCharacters')} onclick={() => session.openDialog('home.enclose')} disabled={!session.enabled(commands.encloseCharactersCommand)} />
    </div>
  </div>
</Group>

<Group label={t('group.paragraph')}>
  <div class="rows">
    <div class="row">
      <SplitButton id="home.bullets" icon="bullets" tip={t('tip.bullets')} onclick={() => (listKind === 'bullet' ? session.apply(commands.removeListCommand, undefined) : applyList(BULLET_PRESETS[0] as ListPreset))} on={listKind === 'bullet'} disabled={!session.enabled(commands.applyListPresetCommand)}>
        {#snippet menu()}
          <div class="menu-head">{t('home.list.bulletLibrary')}</div>
          <div class="list-gallery">
            <button class="list-tile none" role="menuitem" onclick={() => session.apply(commands.removeListCommand, undefined)}>{t('home.list.none')}</button>
            {#each BULLET_PRESETS as preset (preset.id)}
              <button class="list-tile" role="menuitem" title={preset.id} onclick={() => applyList(preset)}>
                {#each previewLabels(preset.levels[0]) as label, i (i)}<span><b>{label}</b><i></i></span>{/each}
              </button>
            {/each}
          </div>
          {@render levelMenu()}
          <hr />
          <button class="mi" role="menuitem" onclick={() => session.openDialog('home.defineBullet')}>{t('home.list.defineBullet')}</button>
        {/snippet}
      </SplitButton>
      <SplitButton id="home.numbering" icon="numbering" tip={t('tip.numbering')} onclick={() => (listKind === 'numbered' ? session.apply(commands.removeListCommand, undefined) : applyList(NUMBERING_PRESETS[0] as ListPreset))} on={listKind === 'numbered'} disabled={!session.enabled(commands.applyListPresetCommand)}>
        {#snippet menu()}
          <div class="menu-head">{t('home.list.numberingLibrary')}</div>
          <div class="list-gallery">
            <button class="list-tile none" role="menuitem" onclick={() => session.apply(commands.removeListCommand, undefined)}>{t('home.list.none')}</button>
            {#each NUMBERING_PRESETS as preset (preset.id)}
              <button class="list-tile" role="menuitem" title={preset.id} onclick={() => applyList(preset)}>
                {#each previewLabels(preset.levels[0]) as label, i (i)}<span><b>{label}</b><i></i></span>{/each}
              </button>
            {/each}
          </div>
          {@render levelMenu()}
          <hr />
          <button class="mi" role="menuitem" disabled={!caretInList} onclick={() => session.apply(commands.restartNumberingCommand, {})}>{t('home.list.restart')}</button>
          <button class="mi" role="menuitem" disabled={!caretInList} onclick={() => session.apply(commands.continueNumberingCommand, undefined)}>{t('home.list.continue')}</button>
          <button class="mi" role="menuitem" disabled={!caretInList} onclick={() => session.openDialog('home.numberingValue')}>{t('home.list.setValue')}</button>
        {/snippet}
      </SplitButton>
      <SplitButton id="home.multilevel" icon="multilevel" tip={t('home.list.multilevel')} disabled={!session.enabled(commands.applyListPresetCommand)}>
        {#snippet menu()}
          <div class="menu-head">{t('home.list.listLibrary')}</div>
          <div class="list-gallery">
            <button class="list-tile none" role="menuitem" onclick={() => session.apply(commands.removeListCommand, undefined)}>{t('home.list.none')}</button>
            {#each MULTILEVEL_PRESETS as preset (preset.id)}
              <button class="list-tile multi" role="menuitem" title={preset.id} onclick={() => applyList(preset)}>
                {#each multilevelPreview(preset.levels) as label, i (i)}<span style="padding-left:{i * 8}px"><b>{label}</b><i></i></span>{/each}
              </button>
            {/each}
          </div>
          {@render levelMenu()}
        {/snippet}
      </SplitButton>
      <span class="sep"></span>
      <Button icon="indentLess" tip={t('tip.indentLess')} onclick={() => session.apply(commands.indentStepCommand, { direction: 'decrease' })} disabled={!session.enabled(commands.indentStepCommand)} />
      <Button icon="indentMore" tip={t('tip.indentMore')} onclick={() => session.apply(commands.indentStepCommand, { direction: 'increase' })} disabled={!session.enabled(commands.indentStepCommand)} />
      <span class="sep"></span>
      <SplitButton id="home.asianLayout" icon="asianLayout" tip={t('home.asian.title')} disabled={!session.enabled(commands.eastAsianLayoutCommand)}>
        {#snippet menu()}
          <button class="mi" role="menuitem" onclick={() => session.apply(commands.eastAsianLayoutCommand, { layout: { kind: 'horizontalInVertical', fitLine: true } })}>{t('home.asian.horizontalInVertical')}</button>
          <button class="mi" role="menuitem" onclick={() => session.openDialog('home.twoLines')}>{t('home.asian.twoLines')}</button>
          <button class="mi" role="menuitem" onclick={() => session.openDialog('home.fitText')}>{t('home.asian.fitText')}</button>
          <div class="menu-head">{t('home.asian.scale')}</div>
          <div class="scale-list">
            {#each CHARACTER_SCALES as scale (scale)}
              <button class="mi" role="menuitem" onclick={() => session.apply(commands.fontFormatCommand, { scale })}>{scale}%</button>
            {/each}
          </div>
          <button class="mi" role="menuitem" onclick={() => session.openDialog('home.font')}>{t('home.asian.moreScale')}</button>
        {/snippet}
      </SplitButton>
      <span class="sep"></span>
      <Button icon="sortText" tip={t('home.sort')} onclick={() => session.openDialog('home.sort')} disabled={!session.enabled(commands.sortParagraphsCommand)} />
      <span class="sep"></span>
      <Button icon="pilcrow" tip={`${t('tip.marks')} (⌘8)`} onclick={() => (session.showMarks = !session.showMarks)} on={session.showMarks} />
    </div>
    <div class="row">
      <Button icon="alignLeft" tip={`${t('tip.alignLeft')} (⌘L)`} onclick={() => session.apply(commands.alignLeftCommand, undefined)} on={session.active(commands.alignLeftCommand)} disabled={!session.enabled(commands.alignLeftCommand)} />
      <Button icon="alignCenter" tip={`${t('tip.center')} (⌘E)`} onclick={() => session.apply(commands.alignCenterCommand, undefined)} on={session.active(commands.alignCenterCommand)} disabled={!session.enabled(commands.alignCenterCommand)} />
      <Button icon="alignRight" tip={`${t('tip.alignRight')} (⌘R)`} onclick={() => session.apply(commands.alignRightCommand, undefined)} on={session.active(commands.alignRightCommand)} disabled={!session.enabled(commands.alignRightCommand)} />
      <Button icon="alignJustify" tip={`${t('tip.justify')} (⌘J)`} onclick={() => session.apply(commands.alignJustifyCommand, undefined)} on={session.active(commands.alignJustifyCommand)} disabled={!session.enabled(commands.alignJustifyCommand)} />
      <Button icon="alignDistributed" tip={t('home.distributed')} onclick={() => session.apply(commands.alignDistributedCommand, undefined)} on={session.active(commands.alignDistributedCommand)} disabled={!session.enabled(commands.alignDistributedCommand)} />
      <span class="sep"></span>
      <SplitButton id="home.lineSpacing" icon="lineSpacing" tip={t('tip.lineSpacing')} disabled={!session.enabled(commands.setLineSpacingCommand)}>
        {#snippet menu()}
          {#each LINE_SPACINGS as multiple (multiple)}
            <button class="mi check" class:checked={lineSpacing === multiple} role="menuitemradio" aria-checked={lineSpacing === multiple} onclick={() => session.apply(commands.setLineSpacingCommand, { multiple })}>{multiple.toFixed(multiple === 1.15 ? 2 : 1)}</button>
          {/each}
          <hr />
          <button class="mi" role="menuitem" onclick={() => session.openDialog('home.paragraph')}>{t('home.spacing.options')}</button>
          <hr />
          {#each ['before', 'after'] as const as side (side)}
            {@const has = !!model && commands.hasParagraphSpace(model, side)}
            <button class="mi" role="menuitem" onclick={() => session.apply(commands.toggleParagraphSpaceCommand, { side })}>{t(`home.spacing.${has ? 'remove' : 'add'}.${side}`)}</button>
          {/each}
        {/snippet}
      </SplitButton>
      <SplitButton id="home.shading" tip={t('home.shading')} onclick={() => setShading(home.shadingPen)} disabled={!session.enabled(commands.paragraphShadingColorCommand)}>
        {#snippet face()}<span class="swatch-icon"><RibbonIcon name="paragraphShading" size={18} /><i style="background:{chip(home.shadingPen)}"></i></span>{/snippet}
        {#snippet menu()}<ColorMenu none="noColor" onpick={setShading} />{/snippet}
      </SplitButton>
      <SplitButton id="home.borders" icon="paragraphBorders" tip={t('home.borders')} onclick={() => session.apply(commands.bordersPresetCommand, { preset: 'bottom' })} disabled={!session.enabled(commands.bordersPresetCommand)}>
        {#snippet menu()}
          {#each BORDER_PRESETS as b (b.preset)}
            {@const on = b.preset !== 'none' && !!model && commands.bordersPresetActive(model, b.preset)}
            <button class="mi check" class:checked={on} role="menuitemcheckbox" aria-checked={on} onclick={() => session.apply(commands.bordersPresetCommand, { preset: b.preset })}>{t(b.key)}</button>
          {/each}
          <hr />
          <button class="mi" role="menuitem" onclick={() => session.apply(commands.insertHorizontalLineCommand, undefined)}>{t('home.border.horizontalLine')}</button>
          <button class="mi check" class:checked={session.showGridlines} role="menuitemcheckbox" aria-checked={session.showGridlines} onclick={() => { session.showGridlines = !session.showGridlines; session.openMenu = null; }}>{t('home.border.gridlines')}</button>
          <hr />
          <button class="mi" role="menuitem" onclick={() => session.openDialog('home.borders')}>{t('home.border.dialog')}</button>
        {/snippet}
      </SplitButton>
    </div>
  </div>
</Group>

{#snippet levelMenu()}
  {@const disabled = listKind === undefined}
  <div class="menu-head">{t('home.list.changeLevel')}</div>
  <div class="level-list">
    {#each Array.from({ length: LIST_LEVELS }, (_, i) => i) as ilvl (ilvl)}
      <button class="mi" role="menuitem" {disabled} onclick={() => session.apply(commands.setListLevelCommand, { ilvl })}>{ilvl + 1}</button>
    {/each}
  </div>
{/snippet}

{#snippet tile(entry: commands.StyleEntry)}
  {@const selected = currentStyle === entry.styleId}
  <button class="tile" class:selected role="option" aria-selected={selected} disabled={!session.enabled(commands.applyStyleCommand)} onclick={() => session.apply(commands.applyStyleCommand, { styleId: entry.styleId })} title={entry.name}>
    <span class="sample" style={previews?.css(entry)}>{entry.type === 'character' ? 'AaBbCcDd' : 'AaBbCcDdEe'}</span>
    <span class="tile-name">{entry.name}</span>
  </button>
{/snippet}

<Group label={t('group.styles')} class="gallery-group">
  <div class="split gallery-wrap">
    <div class="gallery" role="listbox" aria-label={t('group.styles')}>
      {#each gallery.slice(0, GALLERY_TILES) as entry (entry.styleId)}{@render tile(entry)}{/each}
    </div>
    <button class="gallery-more" onclick={() => session.toggleMenu('home.styleGallery')} aria-label={t('home.styles.more')} title={t('home.styles.more')} aria-expanded={session.openMenu === 'home.styleGallery'}><RibbonIcon name="chevronRight" size={12} /></button>
    {#if session.openMenu === 'home.styleGallery'}
      <div class="menu style-gallery" role="menu">
        <div class="gallery-grid" role="listbox" aria-label={t('group.styles')}>
          {#each gallery as entry (entry.styleId)}{@render tile(entry)}{/each}
        </div>
        <hr />
        <button class="mi" role="menuitem" onclick={() => { home.styleTarget = null; session.openDialog('home.style'); }}>{t('home.styles.create')}</button>
        <button class="mi" role="menuitem" onclick={() => session.apply(commands.clearFormattingCommand, undefined)}>{t('tip.clearAll')}</button>
        <button class="mi" role="menuitem" onclick={() => { session.openMenu = null; session.pane.right = 'styles'; }}>{t('home.styles.apply')}</button>
      </div>
    {/if}
  </div>
  <Button size="large" icon="stylesPane" label={t('home.styles.pane')} tip={t('home.styles.pane')} onclick={() => session.togglePane('right', 'styles')} on={session.pane.right === 'styles'} />
</Group>
