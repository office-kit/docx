<script lang="ts">
  /**
   * Word's Page Setup dialog: Margins (incl. gutter, gutter position and the
   * multiple-pages modes), Paper, Layout (section start, headers and footers,
   * vertical alignment, Line Numbers…, Borders…) and Document Grid. OK applies
   * every tab at once to the chosen "Apply to" scope, as Word does.
   */
  import { commands, createStyleResolver, type SectionTarget } from '@office-kit/docx-editor';
  import {
    getDocumentSetting,
    type DocumentGrid,
    type PageMargins,
    type SectionStart,
    type SectionTextDirection,
    type VerticalAlignment,
  } from '@office-kit/docx';
  import Dialog from '../Dialog.svelte';
  import LengthField from './LengthField.svelte';
  import { getSession } from '../session.svelte';
  import { locale, t, type MessageKey } from '../i18n/index.svelte';
  import { lengthUnitFor, TWIPS_PER_POINT } from './units';
  import { matchesPaper, PAPER_PRESETS } from './presets';
  import { layoutState, type PageSetupTab } from './layout-state.svelte';

  const ID = 'layout.pageSetup';
  const session = getSession();
  const unit = $derived(lengthUnitFor(locale()));
  const TABS: ReadonlyArray<{ id: PageSetupTab; key: MessageKey }> = [
    { id: 'margins', key: 'lay.ps.tabMargins' },
    { id: 'paper', key: 'lay.ps.tabPaper' },
    { id: 'layout', key: 'lay.ps.tabLayout' },
    { id: 'grid', key: 'lay.ps.tabGrid' },
  ];
  const STARTS: ReadonlyArray<{ value: SectionStart; key: MessageKey }> = [
    { value: 'continuous', key: 'lay.ps.startContinuous' },
    { value: 'nextColumn', key: 'lay.ps.startColumn' },
    { value: 'nextPage', key: 'lay.ps.startPage' },
    { value: 'evenPage', key: 'lay.ps.startEven' },
    { value: 'oddPage', key: 'lay.ps.startOdd' },
  ];
  const VALIGNS: ReadonlyArray<{ value: VerticalAlignment; key: MessageKey }> = [
    { value: 'top', key: 'lay.ps.vTop' },
    { value: 'center', key: 'lay.ps.vCenter' },
    { value: 'both', key: 'lay.ps.vJustified' },
    { value: 'bottom', key: 'lay.ps.vBottom' },
  ];
  const SHEETS = [0, 4, 8, 12, 16, 20, 24, 28, 32, 36, 40];
  // Word's limits (twips): 22 in page edge, and its grid ranges.
  const MAX_PAGE = 31680;
  const MIN_PAGE = 144;
  const LINE_UNIT_DEFAULT_HALF_POINTS = 21;
  const CHAR_SPACE_UNIT = 4096;

  type Pages = 'normal' | 'mirror' | 'twoOnOne' | 'bookFold';
  let open = false;
  let margins = $state<PageMargins>({ top: 1440, bottom: 1440, left: 1440, right: 1440, header: 720, footer: 720, gutter: 0 });
  let gutterAtTop = $state(false);
  let orientation = $state<'portrait' | 'landscape'>('portrait');
  let pages = $state<Pages>('normal');
  let sheets = $state(0);
  let width = $state(12240);
  let height = $state(15840);
  let start = $state<SectionStart>('nextPage');
  let evenOdd = $state(false);
  let titlePage = $state(false);
  let vAlign = $state<VerticalAlignment>('top');
  let vertical = $state(false);
  let columns = $state(1);
  let gridType = $state<DocumentGrid['type']>('default');
  let linePitch = $state(360);
  let charSpace = $state(0);
  let applyTo = $state<SectionTarget>('section');
  let fontHalfPoints = LINE_UNIT_DEFAULT_HALF_POINTS;

  $effect(() => {
    const showing = session.dialog === ID;
    if (showing && !open) load();
    open = showing;
  });

  function load(): void {
    const model = session.model;
    if (!model) return;
    const p = commands.currentSectionProperties(model);
    const setting = (name: string): boolean => getDocumentSetting(model.doc, name).present;
    margins = { ...p.margins };
    gutterAtTop = setting('gutterAtTop');
    orientation = p.pageSize.orientation;
    pages = setting('mirrorMargins') ? 'mirror' : setting('printTwoOnOne') ? 'twoOnOne' : setting('bookFoldPrinting') ? 'bookFold' : 'normal';
    sheets = Number(getDocumentSetting(model.doc, 'bookFoldPrintingSheets').val ?? 0);
    width = p.pageSize.widthTwips;
    height = p.pageSize.heightTwips;
    start = p.start;
    evenOdd = setting('evenAndOddHeaders');
    titlePage = p.titlePage;
    vAlign = p.verticalAlignment;
    vertical = p.textDirection === 'tbRl';
    columns = p.columns.count;
    gridType = p.documentGrid?.type ?? 'default';
    linePitch = p.documentGrid?.linePitch ?? 360;
    charSpace = p.documentGrid?.charSpace ?? 0;
    fontHalfPoints = createStyleResolver(model.doc).run({ kind: 'paragraph', children: [], extras: [] }).sizeHalfPoints ?? LINE_UNIT_DEFAULT_HALF_POINTS;
    applyTo = 'section';
  }

  function setOrientation(next: 'portrait' | 'landscape'): void {
    if (next === orientation) return;
    orientation = next;
    [width, height] = [height, width];
  }

  function pickPaper(index: string): void {
    const preset = PAPER_PRESETS[Number(index)];
    if (!preset) return;
    const [w, h] = orientation === 'landscape' ? [preset.heightTwips, preset.widthTwips] : [preset.widthTwips, preset.heightTwips];
    width = w;
    height = h;
  }

  const paperIndex = $derived(PAPER_PRESETS.findIndex((p) => matchesPaper(p, width, height)));
  const textWidth = $derived(width - margins.left - margins.right - margins.gutter);
  const textHeight = $derived(height - margins.top - margins.bottom);
  const charPitch = $derived(fontHalfPoints * 10 + Math.round((charSpace / CHAR_SPACE_UNIT) * TWIPS_PER_POINT));
  const charsPerLine = $derived(Math.max(1, Math.floor(textWidth / Math.max(1, charPitch))));
  const linesPerPage = $derived(Math.max(1, Math.floor(textHeight / Math.max(1, linePitch))));

  function setCharsPerLine(n: number): void {
    if (!(n >= 1)) return;
    const pitchPt = textWidth / n / TWIPS_PER_POINT;
    charSpace = Math.round((pitchPt - fontHalfPoints / 2) * CHAR_SPACE_UNIT);
  }

  function setLinesPerPage(n: number): void {
    if (n >= 1) linePitch = Math.floor(textHeight / n);
  }

  function ok(): boolean | void {
    if (width < MIN_PAGE || height < MIN_PAGE || textWidth <= 0 || textHeight <= 0) return false;
    const textDirection: SectionTextDirection = vertical ? 'tbRl' : 'lrTb';
    const preset = PAPER_PRESETS[paperIndex];
    session.apply(commands.pageSetupCommand, {
      target: applyTo,
      pageSize: { widthTwips: width, heightTwips: height, orientation, ...(preset ? { paperCode: preset.code } : {}) },
      margins,
      section: {
        start,
        verticalAlignment: vAlign,
        titlePage,
        textDirection,
        documentGrid:
          gridType === 'default'
            ? null
            : { type: gridType, linePitch, ...(gridType === 'linesAndChars' || gridType === 'snapToChars' ? { charSpace } : {}) },
        ...(columns !== (session.model ? commands.currentSectionProperties(session.model).columns.count : 1)
          ? { columns: { count: columns, spaceTwips: 720, separator: false } }
          : {}),
      },
      settings: {
        mirrorMargins: pages === 'mirror',
        printTwoOnOne: pages === 'twoOnOne',
        bookFoldPrinting: pages === 'bookFold',
        bookFoldPrintingSheets: pages === 'bookFold' ? sheets : 0,
        gutterAtTop,
        evenAndOddHeaders: evenOdd,
      },
    });
    layoutState.lastCustomMargins = { ...margins };
  }

  const inside = $derived(pages === 'mirror' || pages === 'bookFold');
</script>

<Dialog id={ID} title={t('lay.ps.title')} onok={ok}>
  <div class="dialog-tabs" role="tablist">
    {#each TABS as tab (tab.id)}
      <button type="button" role="tab" class="dialog-tab" class:on={layoutState.pageSetupTab === tab.id} aria-selected={layoutState.pageSetupTab === tab.id} onclick={() => (layoutState.pageSetupTab = tab.id)}>{t(tab.key)}</button>
    {/each}
  </div>

  {#if layoutState.pageSetupTab === 'margins'}
    <fieldset>
      <legend>{t('lay.ps.margins')}</legend>
      <div class="form-grid four">
        <label for="ps-top">{t('lay.top')}</label>
        <LengthField id="ps-top" value={margins.top} {unit} label={t('lay.top')} min={-31680} onchange={(v) => (margins = { ...margins, top: v })} />
        <label for="ps-bottom">{t('lay.bottom')}</label>
        <LengthField id="ps-bottom" value={margins.bottom} {unit} label={t('lay.bottom')} min={-31680} onchange={(v) => (margins = { ...margins, bottom: v })} />
        <label for="ps-left">{t(inside ? 'lay.inside' : 'lay.left')}</label>
        <LengthField id="ps-left" value={margins.left} {unit} label={t(inside ? 'lay.inside' : 'lay.left')} onchange={(v) => (margins = { ...margins, left: v })} />
        <label for="ps-right">{t(inside ? 'lay.outside' : 'lay.right')}</label>
        <LengthField id="ps-right" value={margins.right} {unit} label={t(inside ? 'lay.outside' : 'lay.right')} onchange={(v) => (margins = { ...margins, right: v })} />
        <label for="ps-gutter">{t('lay.ps.gutter')}</label>
        <LengthField id="ps-gutter" value={margins.gutter} {unit} label={t('lay.ps.gutter')} onchange={(v) => (margins = { ...margins, gutter: v })} />
        <label for="ps-gutterpos">{t('lay.ps.gutterPosition')}</label>
        <select id="ps-gutterpos" disabled={pages !== 'normal'} value={gutterAtTop ? 'top' : 'left'} onchange={(e) => (gutterAtTop = (e.currentTarget as HTMLSelectElement).value === 'top')}>
          <option value="left">{t('lay.left')}</option>
          <option value="top">{t('lay.top')}</option>
        </select>
      </div>
    </fieldset>
    <fieldset>
      <legend>{t('lay.orientation')}</legend>
      <label class="radio"><input type="radio" name="ps-orient" checked={orientation === 'portrait'} onchange={() => setOrientation('portrait')} />{t('lay.portrait')}</label>
      <label class="radio"><input type="radio" name="ps-orient" checked={orientation === 'landscape'} onchange={() => setOrientation('landscape')} />{t('lay.landscape')}</label>
    </fieldset>
    <fieldset>
      <legend>{t('lay.ps.pages')}</legend>
      <div class="form-grid">
        <label for="ps-pages">{t('lay.ps.multiplePages')}</label>
        <select id="ps-pages" bind:value={pages}>
          <option value="normal">{t('lay.ps.normal')}</option>
          <option value="mirror">{t('lay.ps.mirror')}</option>
          <option value="twoOnOne">{t('lay.ps.twoOnOne')}</option>
          <option value="bookFold">{t('lay.ps.bookFold')}</option>
        </select>
        {#if pages === 'bookFold'}
          <label for="ps-sheets">{t('lay.ps.sheets')}</label>
          <select id="ps-sheets" bind:value={sheets}>{#each SHEETS as s (s)}<option value={s}>{s === 0 ? t('lay.ps.all') : s}</option>{/each}</select>
        {/if}
      </div>
    </fieldset>
  {:else if layoutState.pageSetupTab === 'paper'}
    <fieldset>
      <legend>{t('lay.ps.paperSize')}</legend>
      <select aria-label={t('lay.ps.paperSize')} value={String(paperIndex)} onchange={(e) => pickPaper((e.currentTarget as HTMLSelectElement).value)}>
        {#each PAPER_PRESETS as p, i (p.name)}<option value={String(i)}>{p.name}</option>{/each}
        <option value="-1" disabled>{t('lay.ps.customSize')}</option>
      </select>
      <div class="form-grid">
        <label for="ps-width">{t('lay.ps.width')}</label>
        <LengthField id="ps-width" value={width} {unit} label={t('lay.ps.width')} min={MIN_PAGE} max={MAX_PAGE} onchange={(v) => (width = v)} />
        <label for="ps-height">{t('lay.ps.height')}</label>
        <LengthField id="ps-height" value={height} {unit} label={t('lay.ps.height')} min={MIN_PAGE} max={MAX_PAGE} onchange={(v) => (height = v)} />
      </div>
    </fieldset>
  {:else if layoutState.pageSetupTab === 'layout'}
    <fieldset>
      <legend>{t('lay.ps.section')}</legend>
      <div class="form-grid">
        <label for="ps-start">{t('lay.ps.sectionStart')}</label>
        <select id="ps-start" bind:value={start}>{#each STARTS as s (s.value)}<option value={s.value}>{t(s.key)}</option>{/each}</select>
      </div>
    </fieldset>
    <fieldset>
      <legend>{t('lay.ps.headersFooters')}</legend>
      <label class="check"><input type="checkbox" bind:checked={evenOdd} />{t('lay.ps.differentOddEven')}</label>
      <label class="check"><input type="checkbox" bind:checked={titlePage} />{t('lay.ps.differentFirst')}</label>
      <div class="form-grid">
        <span class="muted">{t('lay.ps.fromEdge')}</span><span></span>
        <label for="ps-header">{t('lay.ps.header')}</label>
        <LengthField id="ps-header" value={margins.header} {unit} label={t('lay.ps.header')} onchange={(v) => (margins = { ...margins, header: v })} />
        <label for="ps-footer">{t('lay.ps.footer')}</label>
        <LengthField id="ps-footer" value={margins.footer} {unit} label={t('lay.ps.footer')} onchange={(v) => (margins = { ...margins, footer: v })} />
      </div>
    </fieldset>
    <fieldset>
      <legend>{t('lay.ps.page')}</legend>
      <div class="form-grid">
        <label for="ps-valign">{t('lay.ps.verticalAlignment')}</label>
        <select id="ps-valign" bind:value={vAlign}>{#each VALIGNS as v (v.value)}<option value={v.value}>{t(v.key)}</option>{/each}</select>
      </div>
    </fieldset>
    <div class="row">
      <button type="button" class="push" onclick={() => (session.dialog = 'layout.lineNumbers')}>{t('lay.ps.lineNumbers')}</button>
      <button type="button" class="push" onclick={() => (session.dialog = 'design.pageBorders')}>{t('lay.ps.borders')}</button>
    </div>
  {:else}
    <fieldset>
      <legend>{t('lay.ps.textFlow')}</legend>
      <label class="radio"><input type="radio" name="ps-dir" checked={!vertical} onchange={() => (vertical = false)} />{t('lay.ps.horizontal')}</label>
      <label class="radio"><input type="radio" name="ps-dir" checked={vertical} onchange={() => (vertical = true)} />{t('lay.ps.vertical')}</label>
      <div class="form-grid">
        <label for="ps-cols">{t('lay.ps.numberOfColumns')}</label>
        <input id="ps-cols" type="number" min="1" max="45" bind:value={columns} />
      </div>
    </fieldset>
    <fieldset>
      <legend>{t('lay.ps.grid')}</legend>
      <label class="radio"><input type="radio" bind:group={gridType} value="default" />{t('lay.ps.noGrid')}</label>
      <label class="radio"><input type="radio" bind:group={gridType} value="linesAndChars" />{t('lay.ps.lineAndCharGrid')}</label>
      <label class="radio"><input type="radio" bind:group={gridType} value="lines" />{t('lay.ps.lineGrid')}</label>
      <label class="radio"><input type="radio" bind:group={gridType} value="snapToChars" />{t('lay.ps.snapToChars')}</label>
    </fieldset>
    <div class="row two-col">
      <fieldset>
        <legend>{t('lay.ps.characters')}</legend>
        <label class="field">{t('lay.ps.perLine')}<input type="number" min="1" disabled={gridType !== 'linesAndChars' && gridType !== 'snapToChars'} value={charsPerLine} onchange={(e) => setCharsPerLine(Number((e.currentTarget as HTMLInputElement).value))} /></label>
        <span class="muted">{t('lay.ps.pitch')}: {(charPitch / TWIPS_PER_POINT).toFixed(1)} pt</span>
      </fieldset>
      <fieldset>
        <legend>{t('lay.ps.lines')}</legend>
        <label class="field">{t('lay.ps.perPage')}<input type="number" min="1" disabled={gridType === 'default'} value={linesPerPage} onchange={(e) => setLinesPerPage(Number((e.currentTarget as HTMLInputElement).value))} /></label>
        <span class="muted">{t('lay.ps.pitch')}: {(linePitch / TWIPS_PER_POINT).toFixed(1)} pt</span>
      </fieldset>
    </div>
  {/if}

  <div class="field apply-to">
    <label for="ps-apply">{t('dsn.applyTo')}</label>
    <select id="ps-apply" bind:value={applyTo}>
      <option value="document">{t('dsn.applyTo.document')}</option>
      <option value="section">{t('dsn.applyTo.section')}</option>
      <option value="forward">{t('lay.forward')}</option>
    </select>
  </div>
</Dialog>
