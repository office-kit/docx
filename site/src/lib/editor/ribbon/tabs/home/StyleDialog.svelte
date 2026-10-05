<script lang="ts">
  /**
   * New Style / Modify Style: name, type, based on, style for following
   * paragraph, and the common formatting the dialog's Formatting row sets.
   * Opened with `homeState().styleTarget` (`null` = New Style).
   */
  import { getElementAttr, stylesPart, type XmlElement } from '@office-kit/docx';
  import { commands, createStyleResolver, type FontPatch, type ParagraphPatch } from '@office-kit/docx-editor';
  import Dialog from '../../../Dialog.svelte';
  import { getSession } from '../../../session.svelte';
  import { t } from '../../../i18n/index.svelte';
  import { homeState } from './state.svelte';

  const session = getSession();
  const home = homeState(session);
  const ID = 'home.style';
  const HALF_POINTS = 2;
  const LINE_UNIT = 240;

  interface Fields {
    name: string;
    type: 'paragraph' | 'character';
    basedOn: string;
    next: string;
    font: string;
    size: string;
    bold: boolean;
    italic: boolean;
    underline: boolean;
    color: string;
    alignment: string;
    line: string;
    quickStyle: boolean;
    autoUpdate: boolean;
  }

  let fields = $state<Fields | null>(null);
  let initial: Fields | null = null;
  let error = $state('');
  const editing = $derived(home.styleTarget);
  const styles = $derived(session.version >= 0 && session.model ? commands.listStyles(session.model) : []);

  function valOf(style: XmlElement | undefined, local: string): string | undefined {
    const el = style?.children.find((c): c is XmlElement => c.kind === 'element' && c.name.local === local);
    return el && (getElementAttr(el, 'val') ?? '1');
  }

  function read(): Fields | null {
    const model = session.model;
    if (!model) return null;
    const styleId = home.styleTarget;
    const normal = commands.defaultParagraphStyleId(model) ?? '';
    if (styleId === null) {
      return { name: '', type: 'paragraph', basedOn: normal, next: '', font: '', size: '', bold: false, italic: false, underline: false, color: '#000000', alignment: '', line: '', quickStyle: true, autoUpdate: false };
    }
    const el = stylesPart(model.doc)?.styles.find((s) => getElementAttr(s, 'styleId') === styleId);
    const own = createStyleResolver(model.doc).style(styleId);
    const entry = styles.find((s) => s.styleId === styleId);
    const line = own.paragraph.lineRule === undefined || own.paragraph.lineRule === 'auto' ? String(own.paragraph.line ?? '') : '';
    return {
      name: entry?.name ?? styleId,
      type: entry?.type ?? 'paragraph',
      basedOn: valOf(el, 'basedOn') ?? '',
      next: valOf(el, 'next') ?? '',
      font: own.run.font ?? '',
      size: own.run.sizeHalfPoints === undefined ? '' : String(own.run.sizeHalfPoints / HALF_POINTS),
      bold: own.run.bold,
      italic: own.run.italic,
      underline: !!own.run.underline && own.run.underline !== 'none',
      color: own.run.color && own.run.color !== 'auto' ? `#${own.run.color}` : '#000000',
      alignment: own.paragraph.alignment ?? '',
      line,
      quickStyle: entry?.quick ?? false,
      autoUpdate: valOf(el, 'autoRedefine') === '1',
    };
  }

  $effect(() => {
    if (session.dialog !== ID) return;
    const loaded = read();
    initial = loaded && structuredClone(loaded);
    fields = loaded;
    error = '';
  });

  function definition(f: Fields, from: Fields): { font?: FontPatch; paragraph?: ParagraphPatch } {
    const font: { -readonly [K in keyof FontPatch]: FontPatch[K] } = {};
    if (f.font.trim() && f.font !== from.font) font.font = { name: f.font.trim() };
    const size = Number(f.size);
    if (f.size !== from.size && size >= 1) font.sizeHalfPoints = Math.round(size * HALF_POINTS);
    if (f.bold !== from.bold || f.italic !== from.italic) font.effects = { bold: f.bold, italic: f.italic };
    if (f.underline !== from.underline) font.underline = f.underline ? { style: 'single' } : null;
    if (f.color !== from.color) font.color = { rgb: f.color.slice(1).toUpperCase() };
    const paragraph: { -readonly [K in keyof ParagraphPatch]: ParagraphPatch[K] } = {};
    if (f.alignment && f.alignment !== from.alignment) paragraph.alignment = f.alignment as NonNullable<ParagraphPatch['alignment']>;
    if (f.line && f.line !== from.line) paragraph.spacing = { line: Number(f.line), lineRule: 'auto' };
    return {
      ...(Object.keys(font).length > 0 ? { font } : {}),
      ...(Object.keys(paragraph).length > 0 && f.type === 'paragraph' ? { paragraph } : {}),
    };
  }

  function ok(): boolean {
    if (!fields || !initial) return false;
    const common = {
      name: fields.name.trim(),
      basedOn: fields.basedOn || undefined,
      next: fields.type === 'paragraph' ? fields.next : undefined,
      quickStyle: fields.quickStyle,
      autoUpdate: fields.autoUpdate,
      ...definition(fields, initial),
    };
    const styleId = home.styleTarget;
    if (styleId === null) session.apply(commands.newStyleCommand, { ...common, type: fields.type });
    else session.apply(commands.modifyStyleCommand, { ...common, styleId });
    // A refused name (empty or taken) is reported in the dialog, which stays
    // open; `apply` leaves the reason in the status line.
    if (session.status) {
      error = session.status;
      return false;
    }
    return true;
  }
</script>

<Dialog id={ID} title={editing === null ? t('home.styles.new') : t('home.styles.modify')} onok={ok}>
  {#if fields}
    <fieldset>
      <legend>{t('home.styles.properties')}</legend>
      <div class="form-grid">
        <label class="field">{t('home.styles.name')}<input bind:value={fields.name} required /></label>
        <label class="field">{t('home.styles.type')}
          <select bind:value={fields.type} disabled={editing !== null}>
            <option value="paragraph">{t('home.styles.paragraph')}</option>
            <option value="character">{t('home.styles.character')}</option>
          </select>
        </label>
        <label class="field">{t('home.styles.basedOn')}
          <select bind:value={fields.basedOn}>
            <option value="">{t('home.styles.noStyle')}</option>
            {#each styles.filter((s) => s.type === fields?.type && s.inDocument && s.styleId !== editing) as s (s.styleId)}<option value={s.styleId}>{s.name}</option>{/each}
          </select>
        </label>
        {#if fields.type === 'paragraph'}
          <label class="field">{t('home.styles.next')}
            <select bind:value={fields.next}>
              <option value="">{t('home.styles.same')}</option>
              {#each styles.filter((s) => s.type === 'paragraph' && s.inDocument) as s (s.styleId)}<option value={s.styleId}>{s.name}</option>{/each}
            </select>
          </label>
        {/if}
      </div>
    </fieldset>
    <fieldset>
      <legend>{t('home.styles.formatting')}</legend>
      <div class="form-grid">
        <label class="field">{t('tip.fontName')}<input bind:value={fields.font} spellcheck="false" /></label>
        <label class="field">{t('tip.fontSize')}<input bind:value={fields.size} inputmode="decimal" /></label>
        <label class="field"><input type="checkbox" bind:checked={fields.bold} />{t('tip.bold')}</label>
        <label class="field"><input type="checkbox" bind:checked={fields.italic} />{t('tip.italic')}</label>
        <label class="field"><input type="checkbox" bind:checked={fields.underline} />{t('tip.underline')}</label>
        <label class="field">{t('tip.fontColor')}<input type="color" bind:value={fields.color} /></label>
        {#if fields.type === 'paragraph'}
          <label class="field">{t('home.para.alignment')}
            <select bind:value={fields.alignment}>
              <option value="" disabled></option>
              <option value="left">{t('tip.alignLeft')}</option>
              <option value="center">{t('tip.center')}</option>
              <option value="right">{t('tip.alignRight')}</option>
              <option value="both">{t('tip.justify')}</option>
            </select>
          </label>
          <label class="field">{t('tip.lineSpacing')}
            <select bind:value={fields.line}>
              <option value="" disabled></option>
              <option value={String(LINE_UNIT)}>{t('home.para.single')}</option>
              <option value={String(LINE_UNIT * 1.5)}>{t('home.para.onePointFive')}</option>
              <option value={String(LINE_UNIT * 2)}>{t('home.para.double')}</option>
            </select>
          </label>
        {/if}
      </div>
    </fieldset>
    <label class="field"><input type="checkbox" bind:checked={fields.quickStyle} />{t('home.styles.addToGallery')}</label>
    <label class="field"><input type="checkbox" bind:checked={fields.autoUpdate} />{t('home.styles.autoUpdate')}</label>
    {#if error}<p class="dialog-error" role="alert">{error}</p>{/if}
  {/if}
</Dialog>
