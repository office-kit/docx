<script lang="ts">
  /**
   * Borders and Shading: the Borders tab (setting, line style, colour, width,
   * sides) and the Shading tab (fill, pattern, pattern colour), applied to the
   * paragraphs or to the selected text.
   */
  import { BORDER_LINE_STYLES, SHADING_PATTERNS, type BorderLine, type BorderLineStyle, type ShadingPattern } from '@office-kit/docx';
  import { commands } from '@office-kit/docx-editor';
  import Dialog from '../../../Dialog.svelte';
  import { getSession } from '../../../session.svelte';
  import { t, type MessageKey } from '../../../i18n/index.svelte';

  const session = getSession();
  const ID = 'home.borders';
  // Word's Width list, in eighths of a point (§17.18.2 ST_EighthPointMeasure).
  const WIDTHS = [2, 4, 6, 8, 12, 18, 24, 36, 48];
  const EIGHTHS = 8;
  type Side = 'top' | 'bottom' | 'left' | 'right' | 'between';
  const SIDES: ReadonlyArray<{ side: Side; key: MessageKey }> = [
    { side: 'top', key: 'home.border.top' },
    { side: 'bottom', key: 'home.border.bottom' },
    { side: 'left', key: 'home.border.left' },
    { side: 'right', key: 'home.border.right' },
    { side: 'between', key: 'home.border.between' },
  ];

  let tab = $state<'borders' | 'shading'>('borders');
  let applyTo = $state<'paragraph' | 'text'>('paragraph');
  let setting = $state<'none' | 'box' | 'shadow' | 'custom'>('box');
  let style = $state<BorderLineStyle>('single');
  let color = $state('#000000');
  let autoColor = $state(true);
  let width = $state(4);
  let sides = $state<Record<Side, boolean>>({ top: true, bottom: true, left: true, right: true, between: false });
  let fill = $state('#FFFFFF');
  let noFill = $state(true);
  let pattern = $state<ShadingPattern>('clear');
  let patternColor = $state('#000000');
  let autoPatternColor = $state(true);
  let touched = $state({ borders: false, shading: false });

  $effect(() => {
    if (session.dialog !== ID) return;
    tab = 'borders';
    touched = { borders: false, shading: false };
  });

  function chooseSetting(next: typeof setting): void {
    setting = next;
    touched.borders = true;
    if (next !== 'custom') sides = { top: next !== 'none', bottom: next !== 'none', left: next !== 'none', right: next !== 'none', between: false };
  }

  function line(): BorderLine {
    return {
      style,
      sizeEighths: width,
      spacePt: applyTo === 'text' ? 0 : 1,
      ...(autoColor ? {} : { color: { rgb: color.slice(1).toUpperCase() } }),
      ...(setting === 'shadow' ? { shadow: true } : {}),
    };
  }

  function ok(): void {
    if (touched.borders) {
      const anySide = setting !== 'none' && Object.values(sides).some(Boolean);
      if (applyTo === 'text') {
        session.apply(commands.setRunBorderCommand, { border: anySide ? line() : undefined });
      } else {
        const l = line();
        session.apply(commands.paragraphBordersCommand, {
          sides: Object.fromEntries(SIDES.map(({ side }) => [side, setting !== 'none' && sides[side] ? l : undefined])),
        });
      }
    }
    if (touched.shading) {
      const shading = noFill && pattern === 'clear'
        ? undefined
        : { pattern, fill: noFill ? 'auto' : fill.slice(1).toUpperCase(), color: autoPatternColor ? 'auto' : patternColor.slice(1).toUpperCase() };
      if (applyTo === 'text') session.apply(commands.setRunShadingCommand, { shading });
      else if (shading) session.apply(commands.setParagraphShadingCommand, shading);
      else session.apply(commands.paragraphShadingColorCommand, { fill: undefined });
    }
  }
</script>

<Dialog id={ID} title={t('home.border.dialog')} onok={ok}>
  <div class="dialog-tabs" role="tablist">
    <button type="button" role="tab" aria-selected={tab === 'borders'} class:active={tab === 'borders'} onclick={() => (tab = 'borders')}>{t('home.borders')}</button>
    <button type="button" role="tab" aria-selected={tab === 'shading'} class:active={tab === 'shading'} onclick={() => (tab = 'shading')}>{t('home.shading')}</button>
  </div>
  {#if tab === 'borders'}
    <div class="form-grid" role="radiogroup" aria-label={t('home.border.setting')}>
      {#each [['none', 'home.border.settingNone'], ['box', 'home.border.settingBox'], ['shadow', 'home.border.settingShadow'], ['custom', 'home.border.settingCustom']] as const as [value, key] (value)}
        <label class="field"><input type="radio" name="wk-border-setting" checked={setting === value} onchange={() => chooseSetting(value)} />{t(key)}</label>
      {/each}
    </div>
    <div class="form-grid" oninput={() => (touched.borders = true)}>
      <label class="field">{t('home.border.style')}
        <select bind:value={style}>{#each BORDER_LINE_STYLES.filter((s) => s !== 'nil' && s !== 'none') as s (s)}<option value={s}>{s}</option>{/each}</select>
      </label>
      <label class="field">{t('home.border.color')}<input type="color" bind:value={color} disabled={autoColor} /><span class="field"><input type="checkbox" bind:checked={autoColor} />{t('color.auto')}</span></label>
      <label class="field">{t('home.border.width')}
        <select bind:value={width}>{#each WIDTHS as w (w)}<option value={w}>{w / EIGHTHS} pt</option>{/each}</select>
      </label>
    </div>
    <fieldset>
      <legend>{t('home.preview')}</legend>
      <div class="check-grid">
        {#each SIDES as s (s.side)}
          <label class="field"><input type="checkbox" bind:checked={sides[s.side]} disabled={applyTo === 'text'} onchange={() => { setting = 'custom'; touched.borders = true; }} />{t(s.key)}</label>
        {/each}
      </div>
    </fieldset>
  {:else}
    <div class="form-grid" oninput={() => (touched.shading = true)} onchange={() => (touched.shading = true)}>
      <label class="field">{t('home.border.fill')}<input type="color" bind:value={fill} disabled={noFill} /><span class="field"><input type="checkbox" bind:checked={noFill} />{t('highlight.none')}</span></label>
      <label class="field">{t('home.border.pattern')}
        <select bind:value={pattern}>{#each SHADING_PATTERNS as p (p)}<option value={p}>{p === 'clear' ? t('home.border.clear') : p}</option>{/each}</select>
      </label>
      <label class="field">{t('home.border.patternColor')}<input type="color" bind:value={patternColor} disabled={autoPatternColor} /><span class="field"><input type="checkbox" bind:checked={autoPatternColor} />{t('color.auto')}</span></label>
    </div>
  {/if}
  <label class="field">{t('home.border.applyTo')}
    <select bind:value={applyTo}>
      <option value="paragraph">{t('home.border.toParagraph')}</option>
      <option value="text">{t('home.border.toText')}</option>
    </select>
  </label>
</Dialog>
