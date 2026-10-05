<script lang="ts">
  /**
   * Design ▸ Page Borders: the Page Border tab of Word's Borders and Shading
   * dialog — setting (None / Box / Shadow / 3-D / Custom), line style,
   * colour, width, art, per-side toggles in the preview, Apply to, and the
   * Options sub-dialog (margins, measure from, front / surround flags).
   */
  import { commands } from '@office-kit/docx-editor';
  import {
    getDocumentSetting,
    PAGE_BORDER_ART,
    PAGE_BORDER_LINE_STYLES,
    type PageBorder,
    type PageBorders,
  } from '@office-kit/docx';
  import Dialog from '../Dialog.svelte';
  import { getSession } from '../session.svelte';
  import { t, type MessageKey } from '../i18n/index.svelte';

  const ID = 'design.pageBorders';
  const OPTIONS_ID = 'design.pageBorderOptions';
  const session = getSession();
  type Setting = 'none' | 'box' | 'shadow' | 'threeD' | 'custom';
  type Side = 'top' | 'left' | 'bottom' | 'right';
  type ApplyTo = 'document' | 'section' | 'firstPage' | 'notFirstPage';
  const SIDES: readonly Side[] = ['top', 'left', 'bottom', 'right'];
  const SETTINGS: ReadonlyArray<{ value: Setting; key: MessageKey }> = [
    { value: 'none', key: 'dsn.pb.none' },
    { value: 'box', key: 'dsn.pb.box' },
    { value: 'shadow', key: 'dsn.pb.shadow' },
    { value: 'threeD', key: 'dsn.pb.threeD' },
    { value: 'custom', key: 'dsn.pb.custom' },
  ];
  // Word's Width list for line borders: eighths of a point.
  const WIDTHS: ReadonlyArray<{ eighths: number; label: string }> = [
    { eighths: 2, label: '¼ pt' },
    { eighths: 4, label: '½ pt' },
    { eighths: 6, label: '¾ pt' },
    { eighths: 8, label: '1 pt' },
    { eighths: 12, label: '1½ pt' },
    { eighths: 18, label: '2¼ pt' },
    { eighths: 24, label: '3 pt' },
    { eighths: 36, label: '4½ pt' },
    { eighths: 48, label: '6 pt' },
  ];
  // Word's defaults: borders sit 24 pt from the text; art is 20 pt wide by default
  // and at most 31 pt (art `w:sz` is in points, not eighths).
  const DEFAULT_SPACE = 24;
  const DEFAULT_ART_SIZE = 20;
  const MAX_ART_SIZE = 31;

  let open = false;
  let setting = $state<Setting>('none');
  let style = $state('single');
  let color = $state('auto');
  let width = $state(4);
  let art = $state('');
  let artSize = $state(DEFAULT_ART_SIZE);
  let sides = $state<Record<Side, boolean>>({ top: true, left: true, bottom: true, right: true });
  let applyTo = $state<ApplyTo>('document');
  let space = $state<Record<Side, number>>({ top: DEFAULT_SPACE, left: DEFAULT_SPACE, bottom: DEFAULT_SPACE, right: DEFAULT_SPACE });
  let offsetFrom = $state<'page' | 'text'>('text');
  let front = $state(true);
  let alignBorders = $state(false);
  let surroundHeader = $state(true);
  let surroundFooter = $state(true);

  $effect(() => {
    const showing = session.dialog === ID;
    if (showing && !open && session.model) load();
    open = showing || session.dialog === OPTIONS_ID;
  });

  function load(): void {
    const model = session.model;
    if (!model) return;
    const pb = commands.currentSectionProperties(model).pageBorders;
    const first = pb && SIDES.map((s) => pb[s]).find((b): b is PageBorder => !!b);
    if (!first || !pb) {
      setting = 'none';
      sides = { top: true, left: true, bottom: true, right: true };
    } else {
      sides = { top: !!pb.top, left: !!pb.left, bottom: !!pb.bottom, right: !!pb.right };
      const all = SIDES.every((s) => pb[s]);
      setting = !all ? 'custom' : first.shadow ? 'shadow' : first.frame ? 'threeD' : 'box';
      const isArt = PAGE_BORDER_ART.includes(first.style);
      art = isArt ? first.style : '';
      style = isArt ? 'single' : first.style;
      if (isArt) artSize = first.size;
      else width = first.size;
      color = first.color;
      space = { top: pb.top?.spacePt ?? DEFAULT_SPACE, left: pb.left?.spacePt ?? DEFAULT_SPACE, bottom: pb.bottom?.spacePt ?? DEFAULT_SPACE, right: pb.right?.spacePt ?? DEFAULT_SPACE };
      offsetFrom = pb.offsetFrom;
      front = pb.zOrder === 'front';
      applyTo = pb.display === 'firstPage' ? 'firstPage' : pb.display === 'notFirstPage' ? 'notFirstPage' : 'document';
    }
    alignBorders = getDocumentSetting(model.doc, 'alignBordersAndEdges').present;
    surroundHeader = !getDocumentSetting(model.doc, 'bordersDoNotSurroundHeader').present;
    surroundFooter = !getDocumentSetting(model.doc, 'bordersDoNotSurroundFooter').present;
  }

  function pickSetting(next: Setting): void {
    setting = next;
    if (next !== 'custom' && next !== 'none') sides = { top: true, left: true, bottom: true, right: true };
  }

  /** Clicking a side in the preview toggles it and makes the setting Custom, as in Word. */
  function toggleSide(side: Side): void {
    sides = { ...sides, [side]: !sides[side] };
    if (setting === 'none') setting = 'custom';
    else if (setting !== 'custom' && !SIDES.every((s) => sides[s])) setting = 'custom';
  }

  function border(side: Side): PageBorder {
    return {
      style: art || style,
      size: art ? Math.min(MAX_ART_SIZE, Math.max(1, artSize)) : width,
      spacePt: space[side],
      color: art ? 'auto' : color,
      ...(setting === 'shadow' ? { shadow: true } : {}),
      ...(setting === 'threeD' ? { frame: true } : {}),
    };
  }

  function ok(): void {
    let borders: PageBorders | null = null;
    if (setting !== 'none' && SIDES.some((s) => sides[s])) {
      borders = {
        offsetFrom,
        display: applyTo === 'firstPage' ? 'firstPage' : applyTo === 'notFirstPage' ? 'notFirstPage' : 'allPages',
        zOrder: front ? 'front' : 'back',
        ...Object.fromEntries(SIDES.filter((s) => sides[s]).map((s) => [s, border(s)])),
      };
    }
    session.apply(commands.pageBordersCommand, {
      borders,
      target: applyTo === 'document' ? 'document' : 'section',
      options: { alignBordersAndEdges: alignBorders, surroundHeader, surroundFooter },
    });
  }

  function lineCss(): string {
    if (art) return '3px dotted #c55';
    const px = Math.max(1, Math.round(width / 8));
    const css = style.includes('double') ? 'double' : style.startsWith('dot') ? 'dotted' : style.startsWith('dash') ? 'dashed' : 'solid';
    return `${css === 'double' ? Math.max(3, px) : px}px ${css} ${color === 'auto' ? '#000' : `#${color}`}`;
  }
</script>

<Dialog id={ID} title={t('dsn.pb.title')} onok={ok}>
  <div class="dialog-tabs" role="tablist"><span class="dialog-tab on" role="tab" aria-selected="true">{t('dsn.pb.pageBorder')}</span></div>
  <div class="pb-layout">
    <fieldset class="pb-settings">
      <legend>{t('dsn.pb.setting')}</legend>
      {#each SETTINGS as s (s.value)}
        <label class="radio pb-setting"><input type="radio" name="pb-setting" checked={setting === s.value} onchange={() => pickSetting(s.value)} /><span class="pb-swatch pb-{s.value}"></span>{t(s.key)}</label>
      {/each}
    </fieldset>
    <div class="pb-style">
      <label for="pb-style">{t('dsn.pb.style')}</label>
      <select id="pb-style" size="6" bind:value={style} disabled={!!art}>
        {#each PAGE_BORDER_LINE_STYLES as s (s)}<option value={s}>{s}</option>{/each}
      </select>
      <label for="pb-color">{t('dsn.pb.color')}</label>
      <span class="row">
        <input id="pb-color" type="color" value={color === 'auto' ? '#000000' : `#${color}`} disabled={!!art} oninput={(e) => (color = (e.currentTarget as HTMLInputElement).value.slice(1).toUpperCase())} />
        <button type="button" class="push" disabled={!!art} onclick={() => (color = 'auto')}>{t('dsn.automatic')}</button>
      </span>
      <label for="pb-width">{t('dsn.pb.width')}</label>
      {#if art}
        <input id="pb-width" type="number" min="1" max={MAX_ART_SIZE} bind:value={artSize} />
      {:else}
        <select id="pb-width" bind:value={width}>{#each WIDTHS as w (w.eighths)}<option value={w.eighths}>{w.label}</option>{/each}</select>
      {/if}
      <label for="pb-art">{t('dsn.pb.art')}</label>
      <select id="pb-art" bind:value={art} onchange={() => { if (setting === 'none') setting = 'box'; }}>
        <option value="">{t('dsn.pb.noArt')}</option>
        {#each PAGE_BORDER_ART as a (a)}<option value={a}>{a}</option>{/each}
      </select>
    </div>
    <fieldset class="pb-preview">
      <legend>{t('dsn.preview')}</legend>
      <p class="muted">{t('dsn.pb.previewHint')}</p>
      <div class="pb-page">
        {#each SIDES as side (side)}
          <button type="button" class="pb-side pb-side-{side}" aria-pressed={setting !== 'none' && sides[side]} aria-label={t(`dsn.pb.${side}`)} onclick={() => toggleSide(side)} style={setting !== 'none' && sides[side] ? `border-${side === 'left' || side === 'right' ? 'left' : 'top'}:${lineCss()}` : ''}></button>
        {/each}
      </div>
      <label for="pb-apply">{t('dsn.applyTo')}</label>
      <select id="pb-apply" bind:value={applyTo}>
        <option value="document">{t('dsn.applyTo.document')}</option>
        <option value="section">{t('dsn.applyTo.section')}</option>
        <option value="firstPage">{t('dsn.applyTo.firstPage')}</option>
        <option value="notFirstPage">{t('dsn.applyTo.notFirstPage')}</option>
      </select>
      <button type="button" class="push" onclick={() => (session.dialog = OPTIONS_ID)}>{t('dsn.pb.options')}</button>
    </fieldset>
  </div>
</Dialog>

<Dialog id={OPTIONS_ID} title={t('dsn.pb.optionsTitle')} onok={() => { session.dialog = ID; return false; }} okLabel={t('dialog.ok')}>
  <fieldset>
    <legend>{t('dsn.pb.margin')}</legend>
    <div class="form-grid four">
      {#each SIDES as side (side)}
        <label for="pbo-{side}">{t(`dsn.pb.${side}`)}</label>
        <input id="pbo-{side}" type="number" min="0" max="31" bind:value={space[side]} />
      {/each}
    </div>
  </fieldset>
  <div class="field">
    <label for="pbo-from">{t('dsn.pb.measureFrom')}</label>
    <select id="pbo-from" bind:value={offsetFrom}>
      <option value="page">{t('dsn.pb.edgeOfPage')}</option>
      <option value="text">{t('dsn.pb.text')}</option>
    </select>
  </div>
  <fieldset>
    <legend>{t('dsn.pb.optionsLegend')}</legend>
    <label class="check"><input type="checkbox" bind:checked={alignBorders} />{t('dsn.pb.align')}</label>
    <label class="check"><input type="checkbox" bind:checked={surroundHeader} />{t('dsn.pb.surroundHeader')}</label>
    <label class="check"><input type="checkbox" bind:checked={surroundFooter} />{t('dsn.pb.surroundFooter')}</label>
    <label class="check"><input type="checkbox" bind:checked={front} />{t('dsn.pb.front')}</label>
  </fieldset>
</Dialog>
