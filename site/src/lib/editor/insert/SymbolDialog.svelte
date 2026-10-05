<script lang="ts">
  /**
   * Insert ▸ Advanced Symbol: a character grid for a symbol font (inserted as
   * `w:sym`, as Word does) or a Unicode subset of normal text (inserted as
   * text), recently used symbols, and the Special Characters tab.
   */
  import { commands, symbolGlyph } from '@office-kit/docx-editor';
  import Dialog from '../Dialog.svelte';
  import { getSession } from '../session.svelte';
  import { t } from '../i18n/index.svelte';
  import { DIALOG } from './state.svelte';

  const session = getSession();

  // Fonts whose characters Word stores as w:sym (font + private-use code).
  const SYMBOL_FONTS = ['Symbol', 'Wingdings', 'Wingdings 2', 'Wingdings 3', 'Webdings'] as const;
  const NORMAL_TEXT = '';
  // Symbol fonts fill 0x20–0xFF, stored with the U+F000 private-use offset.
  const SYMBOL_FIRST = 0xf020;
  const SYMBOL_LAST = 0xf0ff;
  // Unicode blocks Word's Subset list offers (names are the Unicode block names).
  const SUBSETS = [
    { name: 'Basic Latin', from: 0x20, to: 0x7e },
    { name: 'Latin-1 Supplement', from: 0xa1, to: 0xff },
    { name: 'Latin Extended-A', from: 0x100, to: 0x17f },
    { name: 'Greek and Coptic', from: 0x370, to: 0x3ff },
    { name: 'Cyrillic', from: 0x400, to: 0x4ff },
    { name: 'General Punctuation', from: 0x2010, to: 0x205e },
    { name: 'Currency Symbols', from: 0x20a0, to: 0x20c0 },
    { name: 'Letterlike Symbols', from: 0x2100, to: 0x214f },
    { name: 'Number Forms', from: 0x2150, to: 0x218b },
    { name: 'Arrows', from: 0x2190, to: 0x21ff },
    { name: 'Mathematical Operators', from: 0x2200, to: 0x22ff },
    { name: 'Miscellaneous Technical', from: 0x2300, to: 0x23ff },
    { name: 'Enclosed Alphanumerics', from: 0x2460, to: 0x24ff },
    { name: 'Box Drawing', from: 0x2500, to: 0x257f },
    { name: 'Block Elements', from: 0x2580, to: 0x259f },
    { name: 'Geometric Shapes', from: 0x25a0, to: 0x25ff },
    { name: 'Miscellaneous Symbols', from: 0x2600, to: 0x26ff },
    { name: 'Dingbats', from: 0x2700, to: 0x27bf },
    { name: 'CJK Symbols and Punctuation', from: 0x3000, to: 0x303f },
    { name: 'Hiragana', from: 0x3041, to: 0x309f },
    { name: 'Katakana', from: 0x30a0, to: 0x30ff },
    { name: 'Halfwidth and Fullwidth Forms', from: 0xff01, to: 0xffee },
  ] as const;
  const SPECIAL = [
    { key: 'ins.sc.emDash', char: '—' },
    { key: 'ins.sc.enDash', char: '–' },
    { key: 'ins.sc.nbHyphen', char: 'noBreakHyphen' },
    { key: 'ins.sc.optHyphen', char: 'softHyphen' },
    { key: 'ins.sc.emSpace', char: ' ' },
    { key: 'ins.sc.enSpace', char: ' ' },
    { key: 'ins.sc.quarterSpace', char: ' ' },
    { key: 'ins.sc.nbSpace', char: ' ' },
    { key: 'ins.sc.copyright', char: '©' },
    { key: 'ins.sc.registered', char: '®' },
    { key: 'ins.sc.trademark', char: '™' },
    { key: 'ins.sc.section', char: '§' },
    { key: 'ins.sc.paragraph', char: '¶' },
    { key: 'ins.sc.ellipsis', char: '…' },
    { key: 'ins.sc.singleOpen', char: '‘' },
    { key: 'ins.sc.singleClose', char: '’' },
    { key: 'ins.sc.doubleOpen', char: '“' },
    { key: 'ins.sc.doubleClose', char: '”' },
    { key: 'ins.sc.zeroWidthBreak', char: '​' },
    { key: 'ins.sc.zeroWidthNonBreak', char: '⁠' },
  ] as const;
  const RECENT_KEY = 'wk-editor-recent-symbols';
  const RECENT_MAX = 16;

  type Pick = { readonly font: string; readonly code: number };

  let tab = $state<'symbols' | 'special'>('symbols');
  let font = $state<string>(NORMAL_TEXT);
  let subset = $state(0);
  let selected = $state<Pick | null>(null);
  let special = $state(0);
  let recent = $state<Pick[]>(loadRecent());

  function loadRecent(): Pick[] {
    try {
      const parsed: unknown = JSON.parse(localStorage.getItem(RECENT_KEY) ?? '[]');
      if (!Array.isArray(parsed)) return [];
      return parsed.filter(
        (p): p is Pick => typeof p === 'object' && p !== null && typeof p.font === 'string' && typeof p.code === 'number',
      );
    } catch {
      // Unreadable or blocked storage: start with an empty list.
      return [];
    }
  }

  function remember(pick: Pick): void {
    recent = [pick, ...recent.filter((p) => p.font !== pick.font || p.code !== pick.code)].slice(0, RECENT_MAX);
    try {
      localStorage.setItem(RECENT_KEY, JSON.stringify(recent));
    } catch {
      // Not persisted when storage is blocked; the list still works this session.
    }
  }

  const codes = $derived.by((): number[] => {
    const range = font === NORMAL_TEXT ? SUBSETS[subset] : { from: SYMBOL_FIRST, to: SYMBOL_LAST };
    if (!range) return [];
    const out: number[] = [];
    for (let c = range.from; c <= range.to; c++) out.push(c);
    return out;
  });

  const hex = (code: number): string => code.toString(16).toUpperCase().padStart(4, '0');
  const glyph = (p: Pick): string => (p.font === NORMAL_TEXT ? String.fromCodePoint(p.code) : symbolGlyph(p.font, hex(p.code)));
  const fontStyle = (f: string): string => (f === NORMAL_TEXT ? '' : `font-family: '${f}'`);

  function insertPick(pick: Pick): void {
    if (pick.font === NORMAL_TEXT) session.apply(commands.insertTextCommand, { text: String.fromCodePoint(pick.code) });
    else session.apply(commands.insertSymbolCommand, { symbol: { font: pick.font, char: hex(pick.code) } });
    if (session.status === '') remember(pick);
  }

  // Word's Insert leaves the dialog open so several symbols can be added.
  function ok(): boolean {
    if (tab === 'special') {
      const s = SPECIAL[special];
      if (!s) return false;
      if (s.char === 'noBreakHyphen' || s.char === 'softHyphen') session.apply(commands.insertSymbolCommand, { symbol: s.char });
      else session.apply(commands.insertTextCommand, { text: s.char });
      return false;
    }
    if (selected) insertPick(selected);
    return false;
  }

  function sameAs(a: Pick | null, b: Pick): boolean {
    return !!a && a.font === b.font && a.code === b.code;
  }
</script>

<Dialog id={DIALOG.symbol} title={t('ins.sym.title')} onok={ok} okLabel={t('ins.sym.insert')} okDisabled={tab === 'symbols' && !selected}>
  <div class="dialog-tabs" role="tablist">
    <button type="button" role="tab" class:on={tab === 'symbols'} aria-selected={tab === 'symbols'} onclick={() => (tab = 'symbols')}>{t('ins.sym.symbols')}</button>
    <button type="button" role="tab" class:on={tab === 'special'} aria-selected={tab === 'special'} onclick={() => (tab = 'special')}>{t('ins.sym.special')}</button>
  </div>
  {#if tab === 'symbols'}
    <div class="dialog-row">
      <label class="stack">{t('ins.sym.font')}
        <select bind:value={font} onchange={() => (selected = null)}>
          <option value={NORMAL_TEXT}>{t('ins.sym.normalText')}</option>
          {#each SYMBOL_FONTS as f (f)}<option value={f}>{f}</option>{/each}
        </select>
      </label>
      {#if font === NORMAL_TEXT}
        <label class="stack">{t('ins.sym.subset')}
          <select bind:value={subset}>
            {#each SUBSETS as s, i (s.name)}<option value={i}>{s.name}</option>{/each}
          </select>
        </label>
      {/if}
    </div>
    <div class="symbol-grid" role="listbox" aria-label={t('ins.sym.symbols')} style={fontStyle(font)}>
      {#each codes as code (code)}
        {@const pick = { font, code }}
        <button type="button" role="option" aria-selected={sameAs(selected, pick)} class:selected={sameAs(selected, pick)} title="U+{hex(code)}" onclick={() => (selected = pick)} ondblclick={() => insertPick(pick)}>{glyph(pick)}</button>
      {/each}
    </div>
    <span>{t('ins.sym.recent')}</span>
    <div class="symbol-grid recent" role="listbox" aria-label={t('ins.sym.recent')}>
      {#each recent as pick (`${pick.font}:${pick.code}`)}
        <button type="button" role="option" aria-selected={sameAs(selected, pick)} class:selected={sameAs(selected, pick)} style={fontStyle(pick.font)} onclick={() => { font = pick.font; selected = pick; }} ondblclick={() => insertPick(pick)}>{glyph(pick)}</button>
      {/each}
    </div>
    {#if selected}<p class="syntax">{t('ins.sym.code')} {hex(selected.code)}{selected.font ? ` (${selected.font})` : ''}</p>{/if}
  {:else}
    <div class="listbox" role="listbox" aria-label={t('ins.sym.special')} style="height: 300px">
      {#each SPECIAL as s, i (s.key)}
        <button type="button" role="option" aria-selected={i === special} class:selected={i === special} onclick={() => (special = i)}>{t(s.key)}</button>
      {/each}
    </div>
  {/if}
  {#if session.status}<p class="error-note">{session.status}</p>{/if}
</Dialog>
