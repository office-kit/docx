<script lang="ts">
  /**
   * Format ▸ Font (⌘D): the Font and Advanced tabs. Only what the user
   * changes is applied, so a selection with mixed fonts keeps them unless the
   * font box is edited (Word's behaviour).
   */
  import { UNDERLINE_STYLES, type UnderlineStyle } from '@office-kit/docx';
  import { commands, type FontPatch, type RunToggle } from '@office-kit/docx-editor';
  import Dialog from '../../../Dialog.svelte';
  import { getSession } from '../../../session.svelte';
  import { t, type MessageKey } from '../../../i18n/index.svelte';
  import { common, selectionFormats } from './state.svelte';
  import { TWIPS_PER_POINT } from './units';

  const session = getSession();
  const ID = 'home.font';
  const HALF_POINTS = 2;
  const DEFAULT_SCALE = 100;
  // What Word proposes for "Kerning for fonts" when it is off.
  const DEFAULT_KERN_HALF_POINTS = 28;

  type Effect = 'strike' | RunToggle;
  const EFFECTS: ReadonlyArray<{ effect: Effect; key: MessageKey }> = [
    { effect: 'strike', key: 'tip.strike' },
    { effect: 'dstrike', key: 'home.effect.dstrike' },
    { effect: 'smallCaps', key: 'home.effect.smallCaps' },
    { effect: 'caps', key: 'home.effect.caps' },
    { effect: 'vanish', key: 'home.effect.hidden' },
    { effect: 'outline', key: 'home.effect.outline' },
    { effect: 'shadow', key: 'home.effect.shadow' },
    { effect: 'emboss', key: 'home.effect.emboss' },
    { effect: 'imprint', key: 'home.effect.imprint' },
  ];

  interface Fields {
    font: string;
    eastAsiaFont: string;
    style: 'regular' | 'italic' | 'bold' | 'boldItalic' | '';
    size: string;
    color: string;
    autoColor: boolean;
    underline: UnderlineStyle | '';
    underlineColor: string;
    effects: Record<Effect, boolean | null>;
    vertAlign: 'superscript' | 'subscript' | 'baseline' | '';
    scale: number;
    spacing: number;
    position: number;
    kern: boolean;
    kernFrom: number;
  }

  let tab = $state<'font' | 'advanced'>('font');
  let fields = $state<Fields | null>(null);
  let initial: Fields | null = null;

  function read(): Fields | null {
    const model = session.model;
    if (!model) return null;
    const runs = selectionFormats(model).runs;
    const one = <T,>(f: (r: (typeof runs)[number]) => T): T | undefined => common(runs.map(f));
    const bold = one((r) => r.bold);
    const italic = one((r) => r.italic);
    const color = one((r) => r.color);
    const size = one((r) => r.sizeHalfPoints);
    const effect = (e: Effect): boolean | null => (e === 'strike' ? (one((r) => r.strike) ?? null) : (one((r) => r.toggles.has(e)) ?? null));
    return {
      font: one((r) => r.font) ?? '',
      eastAsiaFont: one((r) => r.eastAsiaFont) ?? '',
      style: bold === undefined || italic === undefined ? '' : bold ? (italic ? 'boldItalic' : 'bold') : italic ? 'italic' : 'regular',
      size: size === undefined ? '' : String(size / HALF_POINTS),
      color: color && color !== 'auto' ? `#${color}` : '#000000',
      autoColor: !color || color === 'auto',
      underline: (one((r) => r.underline ?? 'none') ?? '') as UnderlineStyle | '',
      underlineColor: `#${one((r) => r.underlineColor) ?? '000000'}`,
      effects: Object.fromEntries(EFFECTS.map(({ effect: e }) => [e, effect(e)])) as Record<Effect, boolean | null>,
      vertAlign: (one((r) => r.vertAlign ?? 'baseline') ?? '') as Fields['vertAlign'],
      scale: one((r) => r.scale) ?? DEFAULT_SCALE,
      spacing: (one((r) => r.spacing) ?? 0) / TWIPS_PER_POINT,
      position: (one((r) => r.position) ?? 0) / HALF_POINTS,
      kern: one((r) => r.kern !== undefined) ?? false,
      kernFrom: (one((r) => r.kern) ?? DEFAULT_KERN_HALF_POINTS) / HALF_POINTS,
    };
  }

  $effect(() => {
    if (session.dialog !== ID) return;
    const loaded = read();
    initial = loaded && structuredClone(loaded);
    fields = loaded;
    tab = 'font';
  });

  /** The patch of what changed since the dialog opened. */
  function patch(f: Fields, from: Fields): FontPatch {
    const out: { -readonly [K in keyof FontPatch]: FontPatch[K] } = {};
    if (f.font.trim() && f.font !== from.font) out.font = { name: f.font.trim() };
    if (f.eastAsiaFont.trim() && f.eastAsiaFont !== from.eastAsiaFont) out.eastAsiaFont = { name: f.eastAsiaFont.trim() };
    const effects: Partial<Record<Effect | "bold" | "italic", boolean>> = {};
    if (f.style && f.style !== from.style) {
      effects.bold = f.style === 'bold' || f.style === 'boldItalic';
      effects.italic = f.style === 'italic' || f.style === 'boldItalic';
    }
    for (const { effect } of EFFECTS) {
      const v = f.effects[effect];
      if (v !== null && v !== from.effects[effect]) effects[effect] = v;
    }
    if (Object.keys(effects).length > 0) out.effects = effects;
    const size = Number(f.size);
    if (f.size !== from.size && Number.isFinite(size) && size >= 1) out.sizeHalfPoints = Math.round(size * HALF_POINTS);
    if (f.autoColor !== from.autoColor || (!f.autoColor && f.color !== from.color)) {
      out.color = { rgb: f.autoColor ? 'auto' : f.color.slice(1).toUpperCase() };
    }
    if (f.underline && (f.underline !== from.underline || f.underlineColor !== from.underlineColor)) {
      out.underline = f.underline === 'none' ? null : { style: f.underline, color: { rgb: f.underlineColor.slice(1).toUpperCase() } };
    }
    if (f.vertAlign && f.vertAlign !== from.vertAlign) out.vertAlign = f.vertAlign;
    if (f.scale !== from.scale) out.scale = Math.round(f.scale);
    if (f.spacing !== from.spacing) out.spacing = Math.round(f.spacing * TWIPS_PER_POINT);
    if (f.position !== from.position) out.position = Math.round(f.position * HALF_POINTS);
    if (f.kern !== from.kern || f.kernFrom !== from.kernFrom) out.kern = f.kern ? Math.round(f.kernFrom * HALF_POINTS) : null;
    return out;
  }

  function ok(): void {
    if (!fields || !initial) return;
    const p = patch(fields, initial);
    if (Object.keys(p).length > 0) session.apply(commands.fontFormatCommand, p);
  }

  function setVert(v: 'superscript' | 'subscript', on: boolean): void {
    if (fields) fields.vertAlign = on ? v : 'baseline';
  }

  const previewCss = $derived.by(() => {
    if (!fields) return '';
    const f = fields;
    const css = [
      f.font ? `font-family:"${f.font}"` : '',
      f.style === 'bold' || f.style === 'boldItalic' ? 'font-weight:700' : '',
      f.style === 'italic' || f.style === 'boldItalic' ? 'font-style:italic' : '',
      f.autoColor ? '' : `color:${f.color}`,
      f.underline && f.underline !== 'none' ? `text-decoration:underline;text-decoration-color:${f.underlineColor}` : '',
      f.effects.strike || f.effects.dstrike ? 'text-decoration-line:line-through' : '',
      f.effects.caps ? 'text-transform:uppercase' : '',
      f.effects.smallCaps ? 'font-variant-caps:small-caps' : '',
      `letter-spacing:${f.spacing}pt`,
    ];
    return css.filter(Boolean).join(';');
  });
</script>

<Dialog id={ID} title={t('home.font.dialog')} onok={ok}>
  {#if fields}
    <div class="dialog-tabs" role="tablist">
      <button type="button" role="tab" aria-selected={tab === 'font'} class:active={tab === 'font'} onclick={() => (tab = 'font')}>{t('home.font.tabFont')}</button>
      <button type="button" role="tab" aria-selected={tab === 'advanced'} class:active={tab === 'advanced'} onclick={() => (tab = 'advanced')}>{t('home.font.tabAdvanced')}</button>
    </div>
    {#if tab === 'font'}
      <div class="form-grid">
        <label class="field">{t('tip.fontName')}<input bind:value={fields.font} spellcheck="false" /></label>
        <label class="field">{t('home.font.asian')}<input bind:value={fields.eastAsiaFont} spellcheck="false" /></label>
        <label class="field">{t('home.font.style')}
          <select bind:value={fields.style}>
            <option value="" disabled></option>
            <option value="regular">{t('home.font.regular')}</option>
            <option value="italic">{t('tip.italic')}</option>
            <option value="bold">{t('tip.bold')}</option>
            <option value="boldItalic">{t('home.font.boldItalic')}</option>
          </select>
        </label>
        <label class="field">{t('tip.fontSize')}<input bind:value={fields.size} inputmode="decimal" /></label>
        <label class="field">{t('tip.fontColor')}
          <input type="color" bind:value={fields.color} disabled={fields.autoColor} />
          <span class="field"><input type="checkbox" bind:checked={fields.autoColor} />{t('color.auto')}</span>
        </label>
        <label class="field">{t('home.font.underlineStyle')}
          <select bind:value={fields.underline}>
            <option value="" disabled></option>
            {#each UNDERLINE_STYLES as style (style)}<option value={style}>{style === 'none' ? t('home.font.noUnderline') : style}</option>{/each}
          </select>
        </label>
        <label class="field">{t('home.underline.color')}<input type="color" bind:value={fields.underlineColor} disabled={!fields.underline || fields.underline === 'none'} /></label>
      </div>
      <fieldset>
        <legend>{t('home.font.effects')}</legend>
        <div class="check-grid">
          {#each EFFECTS as e (e.effect)}
            <label class="field"><input type="checkbox" checked={fields.effects[e.effect] === true} indeterminate={fields.effects[e.effect] === null} onchange={(ev) => { if (fields) fields.effects[e.effect] = (ev.currentTarget as HTMLInputElement).checked; }} />{t(e.key)}</label>
          {/each}
          <label class="field"><input type="checkbox" checked={fields.vertAlign === 'superscript'} onchange={(ev) => setVert('superscript', (ev.currentTarget as HTMLInputElement).checked)} />{t('tip.superscript')}</label>
          <label class="field"><input type="checkbox" checked={fields.vertAlign === 'subscript'} onchange={(ev) => setVert('subscript', (ev.currentTarget as HTMLInputElement).checked)} />{t('tip.subscript')}</label>
        </div>
      </fieldset>
    {:else}
      <fieldset>
        <legend>{t('home.font.characterSpacing')}</legend>
        <div class="form-grid">
          <label class="field">{t('home.font.scale')}<input type="number" min="1" max="600" bind:value={fields.scale} />%</label>
          <label class="field">{t('home.font.spacing')}<input type="number" step="0.1" bind:value={fields.spacing} />pt</label>
          <label class="field">{t('home.font.position')}<input type="number" step="0.5" bind:value={fields.position} />pt</label>
          <label class="field"><input type="checkbox" bind:checked={fields.kern} />{t('home.font.kerning')}<input type="number" min="1" step="0.5" bind:value={fields.kernFrom} disabled={!fields.kern} />{t('home.font.pointsAbove')}</label>
        </div>
      </fieldset>
    {/if}
    <fieldset>
      <legend>{t('home.preview')}</legend>
      <div class="font-preview" style={previewCss}>{fields.font || 'Aptos'}</div>
    </fieldset>
  {/if}
</Dialog>
