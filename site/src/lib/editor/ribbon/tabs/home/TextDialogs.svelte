<script lang="ts">
  /**
   * The Home tab's smaller dialogs: Phonetic Guide, Enclose Characters, Sort
   * Text, Set Numbering Value, Define New Bullet, Two Lines in One, Fit Text,
   * and More Colors.
   */
  import { ENCLOSURES, RUBY_ALIGNMENTS, type Enclosure, type ListLevel, type RubyAlignment } from '@office-kit/docx';
  import { BULLET_PRESETS, commands } from '@office-kit/docx-editor';
  import Dialog from '../../../Dialog.svelte';
  import { getSession } from '../../../session.svelte';
  import { locale, t, type MessageKey } from '../../../i18n/index.svelte';
  import { homeState, selectionFormats } from './state.svelte';
  import { toTwips, unitFor } from './units';

  const session = getSession();
  const home = homeState(session);
  const HALF_POINTS = 2;
  const HEX6 = /^#?[0-9a-fA-F]{6}$/;
  // Word's Phonetic Guide defaults for 10.5 pt Japanese text.
  const RUBY_DEFAULTS = { alignment: 'center' as RubyAlignment, offset: 0, size: 5 };
  const RUBY_LANGUAGE = 'ja-JP';
  // Word's default body size (10.5 pt in Japanese documents) when the selection has none.
  const DEFAULT_BASE_HALF_POINTS = 21;
  const ALIGN_KEYS: Readonly<Record<RubyAlignment, MessageKey>> = {
    center: 'home.ruby.center',
    distributeLetter: 'home.ruby.distributeLetter',
    distributeSpace: 'home.ruby.distributeSpace',
    left: 'home.ruby.left',
    right: 'home.ruby.right',
    rightVertical: 'home.ruby.rightVertical',
  };
  const ENCLOSURE_KEYS: Readonly<Record<Enclosure, MessageKey>> = {
    circle: 'home.enclose.circle',
    square: 'home.enclose.square',
    triangle: 'home.enclose.triangle',
    diamond: 'home.enclose.diamond',
  };

  /** The selected text as the page shows it (the canvas keeps the DOM selection). */
  function selectedText(): string {
    return document.getSelection()?.toString() ?? '';
  }

  // --- Phonetic Guide ---
  let rubyBase = $state('');
  let rubyText = $state('');
  let rubyAlign = $state<RubyAlignment>(RUBY_DEFAULTS.alignment);
  let rubyOffset = $state(RUBY_DEFAULTS.offset);
  let rubySize = $state(RUBY_DEFAULTS.size);
  let rubyFont = $state('');
  let hasRuby = $state(false);

  $effect(() => {
    if (session.dialog !== 'home.ruby' || !session.model) return;
    const existing = commands.selectedRuby(session.model);
    hasRuby = !!existing;
    rubyBase = existing?.base ?? selectedText();
    rubyText = existing?.ruby ?? '';
    rubyAlign = existing?.alignment ?? RUBY_DEFAULTS.alignment;
    rubySize = existing ? existing.rubySizeHalfPoints / HALF_POINTS : RUBY_DEFAULTS.size;
    rubyOffset = existing ? Math.max(0, (existing.raiseHalfPoints - existing.baseSizeHalfPoints) / HALF_POINTS) : RUBY_DEFAULTS.offset;
    rubyFont = existing?.rubyFont ?? '';
  });

  function applyRuby(ruby: string): void {
    const model = session.model;
    if (!model) return;
    const base = commands.selectedRuby(model)?.baseSizeHalfPoints ?? selectionFormats(model).runs[0]?.sizeHalfPoints ?? DEFAULT_BASE_HALF_POINTS;
    session.apply(commands.phoneticGuideCommand, {
      ruby,
      options: {
        alignment: rubyAlign,
        rubySizeHalfPoints: Math.round(rubySize * HALF_POINTS),
        // Word measures the offset from the top of the base text.
        raiseHalfPoints: base + Math.round(rubyOffset * HALF_POINTS),
        baseSizeHalfPoints: base,
        language: RUBY_LANGUAGE,
        ...(rubyFont.trim() ? { rubyFont: rubyFont.trim() } : {}),
      },
    });
  }

  // --- Enclose Characters ---
  let encloseStyle = $state<'shrinkText' | 'enlargeSymbol'>('shrinkText');
  let enclosure = $state<Enclosure>('circle');
  let encloseText = $state('');
  $effect(() => {
    if (session.dialog === 'home.enclose') encloseText = selectedText();
  });

  // --- Sort Text ---
  let sortBy = $state<commands.SortOptions['by']>('text');
  let sortOrder = $state<commands.SortOptions['order']>('ascending');
  let sortHeader = $state(false);
  let sortCase = $state(false);

  // --- Set Numbering Value ---
  let numberingMode = $state<'new' | 'continue'>('new');
  let startAt = $state(1);

  // --- Define New Bullet ---
  let bulletChar = $state('•');
  let bulletFont = $state('');

  function defineBullet(): boolean {
    const [first, ...rest] = BULLET_PRESETS[0]?.levels ?? [];
    const char = [...bulletChar][0];
    if (!first || !char) return false;
    const font = bulletFont.trim();
    const level: ListLevel = {
      format: 'bullet',
      text: char,
      indentLeft: first.indentLeft,
      hanging: first.hanging,
      ...(font ? { font } : {}),
    };
    session.apply(commands.applyListPresetCommand, { levels: [level, ...rest] });
    return true;
  }

  // --- Two Lines in One / Fit Text ---
  let brackets = $state<'none' | 'round' | 'square' | 'angle' | 'curly'>('none');
  let fitWidth = $state(1);
  const unit = $derived(unitFor(locale()));

  // --- More Colors ---
  let customColor = $state('#000000');
  $effect(() => {
    if (session.dialog === 'home.moreColors' && home.moreColors) customColor = `#${home.moreColors.initial}`;
  });
</script>

<Dialog id="home.ruby" title={t('home.phoneticGuide')} onok={() => applyRuby(rubyText)} okDisabled={!rubyText && !hasRuby}>
  <div class="form-grid">
    <label class="field">{t('home.ruby.base')}<input value={rubyBase} readonly /></label>
    <label class="field">{t('home.ruby.text')}<input bind:value={rubyText} /></label>
    <label class="field">{t('home.para.alignment')}
      <select bind:value={rubyAlign}>{#each RUBY_ALIGNMENTS as a (a)}<option value={a}>{t(ALIGN_KEYS[a])}</option>{/each}</select>
    </label>
    <label class="field">{t('home.ruby.offset')}<input type="number" min="0" step="0.5" bind:value={rubyOffset} />pt</label>
    <label class="field">{t('tip.fontName')}<input bind:value={rubyFont} spellcheck="false" /></label>
    <label class="field">{t('tip.fontSize')}<input type="number" min="1" step="0.5" bind:value={rubySize} />pt</label>
  </div>
  <fieldset>
    <legend>{t('home.preview')}</legend>
    <ruby class="ruby-preview" style="ruby-align:{rubyAlign === 'center' ? 'center' : 'space-around'}">{rubyBase}<rt style="font-size:{rubySize}pt">{rubyText}</rt></ruby>
  </fieldset>
  {#if hasRuby}<div><button type="button" class="push" onclick={() => { applyRuby(''); session.dialog = null; }}>{t('home.ruby.remove')}</button></div>{/if}
</Dialog>

<Dialog id="home.enclose" title={t('home.encloseCharacters')} onok={() => { session.apply(commands.encloseCharactersCommand, { enclosure, style: encloseStyle }); }} okDisabled={!encloseText}>
  <div class="form-grid" role="radiogroup" aria-label={t('home.enclose.style')}>
    <label class="field"><input type="radio" name="wk-enclose-style" value="shrinkText" bind:group={encloseStyle} />{t('home.enclose.shrinkText')}</label>
    <label class="field"><input type="radio" name="wk-enclose-style" value="enlargeSymbol" bind:group={encloseStyle} />{t('home.enclose.enlargeSymbol')}</label>
  </div>
  <div class="form-grid">
    <label class="field">{t('home.enclose.text')}<input value={encloseText} readonly /></label>
    <label class="field">{t('home.enclose.enclosure')}
      <select bind:value={enclosure}>{#each Object.keys(ENCLOSURES) as Enclosure[] as e (e)}<option value={e}>{ENCLOSURES[e]} {t(ENCLOSURE_KEYS[e])}</option>{/each}</select>
    </label>
  </div>
</Dialog>

<Dialog id="home.sort" title={t('home.sort')} onok={() => { session.apply(commands.sortParagraphsCommand, { by: sortBy, order: sortOrder, header: sortHeader, matchCase: sortCase }); }}>
  <div class="form-grid">
    <span>{t('home.sort.by')}</span>
    <label class="field">{t('home.sort.type')}
      <select bind:value={sortBy}>
        <option value="text">{t('home.sort.text')}</option>
        <option value="number">{t('home.sort.number')}</option>
        <option value="date">{t('home.sort.date')}</option>
      </select>
    </label>
    <label class="field"><input type="radio" name="wk-sort-order" value="ascending" bind:group={sortOrder} />{t('home.sort.ascending')}</label>
    <label class="field"><input type="radio" name="wk-sort-order" value="descending" bind:group={sortOrder} />{t('home.sort.descending')}</label>
  </div>
  <label class="field"><input type="checkbox" bind:checked={sortHeader} />{t('home.sort.header')}</label>
  <label class="field"><input type="checkbox" bind:checked={sortCase} />{t('home.find.matchCase')}</label>
</Dialog>

<Dialog id="home.numberingValue" title={t('home.list.setValue')} onok={() => { if (numberingMode === 'new') session.apply(commands.restartNumberingCommand, { start: startAt }); else session.apply(commands.continueNumberingCommand, undefined); }}>
  <label class="field"><input type="radio" name="wk-numbering-mode" value="new" bind:group={numberingMode} />{t('home.list.startNew')}</label>
  <label class="field"><input type="radio" name="wk-numbering-mode" value="continue" bind:group={numberingMode} />{t('home.list.continuePrevious')}</label>
  <label class="field">{t('home.list.valueTo')}<input type="number" min="0" step="1" bind:value={startAt} disabled={numberingMode !== 'new'} /></label>
</Dialog>

<Dialog id="home.defineBullet" title={t('home.list.defineBullet')} onok={defineBullet} okDisabled={!bulletChar}>
  <label class="field">{t('home.list.symbol')}<input bind:value={bulletChar} maxlength="2" /></label>
  <label class="field">{t('tip.fontName')}<input bind:value={bulletFont} spellcheck="false" placeholder="Symbol" /></label>
  <fieldset>
    <legend>{t('home.preview')}</legend>
    <div class="bullet-preview" style={bulletFont ? `font-family:'${bulletFont}'` : ''}>{bulletChar}</div>
  </fieldset>
</Dialog>

<Dialog id="home.twoLines" title={t('home.asian.twoLines')} onok={() => { session.apply(commands.eastAsianLayoutCommand, { layout: { kind: 'twoLinesInOne', brackets } }); }}>
  <label class="field">{t('home.asian.brackets')}
    <select bind:value={brackets}>
      <option value="none">{t('home.para.none')}</option>
      <option value="round">( )</option>
      <option value="square">[ ]</option>
      <option value="angle">&lt; &gt;</option>
      <option value="curly">{'{ }'}</option>
    </select>
  </label>
  <div><button type="button" class="push" onclick={() => { session.apply(commands.eastAsianLayoutCommand, { layout: undefined }); session.dialog = null; }}>{t('home.asian.remove')}</button></div>
</Dialog>

<Dialog id="home.fitText" title={t('home.asian.fitText')} onok={() => { session.apply(commands.fitTextCommand, { twips: toTwips(fitWidth, unit) }); }} okDisabled={!(fitWidth > 0)}>
  <label class="field">{t('home.asian.fitTo')}<input type="number" min="0.1" step="0.1" bind:value={fitWidth} />{unit}</label>
  <div><button type="button" class="push" onclick={() => { session.apply(commands.fitTextCommand, { twips: undefined }); session.dialog = null; }}>{t('home.asian.remove')}</button></div>
</Dialog>

<Dialog id="home.moreColors" title={t('home.color.moreTitle')} onok={() => { if (!HEX6.test(customColor)) return false; home.moreColors?.pick({ rgb: customColor.replace('#', '').toUpperCase() }); }}>
  <label class="field">{t('home.color.custom')}<input type="color" bind:value={customColor} /></label>
  <label class="field">{t('home.color.hex')}<input bind:value={customColor} spellcheck="false" maxlength="7" /></label>
</Dialog>
