<script lang="ts">
  /**
   * Advanced Find and Replace: match case, whole words, wildcards, format
   * filters for the found text and for the replacement, and Word's special
   * characters (^t, ^&, ^^ …).
   */
  import type { RunFormatting } from '@office-kit/docx';
  import { commands, type FindFormat } from '@office-kit/docx-editor';
  import Dialog from '../../../Dialog.svelte';
  import { getSession } from '../../../session.svelte';
  import { t, type MessageKey } from '../../../i18n/index.svelte';
  import { homeState } from './state.svelte';
  import { findOptions, next, search } from './find-ui';

  const session = getSession();
  const home = homeState(session);
  const ID = 'home.find';
  const HALF_POINTS = 2;

  // Special characters Word's Special menu inserts, by the box that takes
  // them. ^p is Find-only: a replacement cannot split a paragraph here.
  const SPECIALS: ReadonlyArray<{ code: string; key: MessageKey; box: 'both' | 'find' | 'replace' }> = [
    { code: '^p', key: 'home.find.special.paragraph', box: 'find' },
    { code: '^t', key: 'home.find.special.tab', box: 'both' },
    { code: '^^', key: 'home.find.special.caret', box: 'both' },
    { code: '^&', key: 'home.find.special.found', box: 'replace' },
    { code: '^?', key: 'home.find.special.anyChar', box: 'find' },
    { code: '^#', key: 'home.find.special.anyDigit', box: 'find' },
    { code: '^$', key: 'home.find.special.anyLetter', box: 'find' },
    { code: '^w', key: 'home.find.special.whiteSpace', box: 'find' },
  ];

  type Tri = '' | 'on' | 'off';
  let tab = $state<'find' | 'replace'>('find');
  let target = $state<'find' | 'replace'>('find');
  let bold = $state<Tri>('');
  let italic = $state<Tri>('');
  let underline = $state<Tri>('');
  let font = $state('');
  let size = $state('');
  let color = $state('');
  let style = $state('');
  let rBold = $state<Tri>('');
  let rItalic = $state<Tri>('');
  let rColor = $state('');
  let status = $state('');
  const styles = $derived(session.version >= 0 && session.model ? commands.listStyles(session.model).filter((s) => s.inDocument) : []);

  const tri = (v: Tri): boolean | undefined => (v === '' ? undefined : v === 'on');

  // The dialog writes its filters into the shared Find options, so the
  // Navigation pane shows the same results.
  $effect(() => {
    const sizeValue = Number(size);
    const format: { -readonly [K in keyof FindFormat]: FindFormat[K] } = {};
    if (tri(bold) !== undefined) format.bold = tri(bold);
    if (tri(italic) !== undefined) format.italic = tri(italic);
    if (tri(underline) !== undefined) format.underline = tri(underline);
    if (font.trim()) format.font = font.trim();
    if (size && sizeValue > 0) format.sizeHalfPoints = Math.round(sizeValue * HALF_POINTS);
    if (color) format.color = color.slice(1).toUpperCase();
    if (style) format.style = style;
    home.find.format = format;
    const replaceFormat: { -readonly [K in keyof RunFormatting]: RunFormatting[K] } = {};
    if (tri(rBold) !== undefined) replaceFormat.bold = tri(rBold);
    if (tri(rItalic) !== undefined) replaceFormat.italic = tri(rItalic);
    if (rColor) replaceFormat.color = rColor.slice(1).toUpperCase();
    home.find.replaceFormat = replaceFormat;
  });

  function params(): commands.ReplaceParams {
    const format = home.find.replaceFormat;
    return { find: findOptions(session, home), replacement: home.find.replace, ...(Object.keys(format).length > 0 ? { format } : {}) };
  }

  function findNext(): void {
    const result = search(session, findOptions(session, home));
    status = result.error ?? (result.matches.length === 0 ? t('home.find.none') : '');
    next(session, result.matches, 1);
  }

  function insertSpecial(code: string): void {
    if (target === 'find') session.search += code;
    else home.find.replace += code;
  }

  function clearFormat(): void {
    bold = italic = underline = rBold = rItalic = '';
    font = size = color = style = rColor = '';
  }
</script>

{#snippet triSelect(label: MessageKey, value: Tri, set: (v: Tri) => void)}
  <label class="field">{t(label)}
    <select value={value} onchange={(e) => set((e.currentTarget as HTMLSelectElement).value as Tri)}>
      <option value="">{t('home.find.any')}</option>
      <option value="on">{t('home.find.yes')}</option>
      <option value="off">{t('home.find.no')}</option>
    </select>
  </label>
{/snippet}

<Dialog id={ID} title={t('home.find.dialog')}>
  <div class="dialog-tabs" role="tablist">
    <button type="button" role="tab" aria-selected={tab === 'find'} class:active={tab === 'find'} onclick={() => { tab = 'find'; target = 'find'; }}>{t('home.find.tabFind')}</button>
    <button type="button" role="tab" aria-selected={tab === 'replace'} class:active={tab === 'replace'} onclick={() => (tab = 'replace')}>{t('home.find.replace')}</button>
  </div>
  <label class="field">{t('home.find.findWhat')}<input bind:value={session.search} onfocus={() => (target = 'find')} /></label>
  {#if tab === 'replace'}
    <label class="field">{t('find.replaceWith')}<input bind:value={home.find.replace} onfocus={() => (target = 'replace')} /></label>
  {/if}
  <div class="check-grid">
    <label class="field"><input type="checkbox" bind:checked={home.find.matchCase} />{t('home.find.matchCase')}</label>
    <label class="field"><input type="checkbox" bind:checked={home.find.wholeWords} disabled={home.find.wildcards} />{t('home.find.wholeWords')}</label>
    <label class="field"><input type="checkbox" bind:checked={home.find.wildcards} />{t('home.find.wildcards')}</label>
  </div>
  <fieldset>
    <legend>{t('home.find.format')}</legend>
    <div class="form-grid">
      <label class="field">{t('tip.fontName')}<input bind:value={font} spellcheck="false" /></label>
      <label class="field">{t('tip.fontSize')}<input bind:value={size} inputmode="decimal" /></label>
      {@render triSelect('tip.bold', bold, (v) => (bold = v))}
      {@render triSelect('tip.italic', italic, (v) => (italic = v))}
      {@render triSelect('tip.underline', underline, (v) => (underline = v))}
      <label class="field">{t('tip.fontColor')}<input type="color" value={color || '#000000'} onchange={(e) => (color = (e.currentTarget as HTMLInputElement).value)} /></label>
      <label class="field">{t('group.styles')}
        <select bind:value={style}>
          <option value="">{t('home.find.any')}</option>
          {#each styles as s (s.styleId)}<option value={s.styleId}>{s.name}</option>{/each}
        </select>
      </label>
    </div>
  </fieldset>
  {#if tab === 'replace'}
    <fieldset>
      <legend>{t('home.find.replaceFormat')}</legend>
      <div class="form-grid">
        {@render triSelect('tip.bold', rBold, (v) => (rBold = v))}
        {@render triSelect('tip.italic', rItalic, (v) => (rItalic = v))}
        <label class="field">{t('tip.fontColor')}<input type="color" value={rColor || '#000000'} onchange={(e) => (rColor = (e.currentTarget as HTMLInputElement).value)} /></label>
      </div>
    </fieldset>
  {/if}
  <div class="field">
    <span>{t('home.find.special')}</span>
    {#each SPECIALS.filter((s) => s.box === 'both' || s.box === target) as s (s.code)}
      <button type="button" class="push" title={s.code} onclick={() => insertSpecial(s.code)}>{t(s.key)}</button>
    {/each}
  </div>
  <div class="field">
    <button type="button" class="push" onclick={clearFormat}>{t('home.find.noFormatting')}</button>
    <button type="button" class="push" onclick={() => { session.pane.left = 'navigation'; session.dialog = null; }}>{t('home.find.findAll')}</button>
    <button type="button" class="push" onclick={findNext}>{t('home.find.next')}</button>
    {#if tab === 'replace'}
      <button type="button" class="push" onclick={() => session.apply(commands.replaceNextCommand, params())}>{t('home.find.replace')}</button>
      <button type="button" class="push" onclick={() => { const n = session.apply(commands.findReplaceAllCommand, params()); if (n !== undefined) status = t('home.find.replaced').replace('{n}', String(n)); }}>{t('find.replaceAll')}</button>
    {/if}
  </div>
  <p class="pane-note" aria-live="polite">{status}</p>
</Dialog>
