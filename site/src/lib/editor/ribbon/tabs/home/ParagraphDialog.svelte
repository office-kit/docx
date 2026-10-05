<script lang="ts">
  /**
   * Format ▸ Paragraph (⌘⌥M): Indents and Spacing, Line and Page Breaks, and
   * Asian Typography. As in the Font dialog, only changed fields are applied.
   */
  import { commands, type ParagraphPatch, type ParagraphToggle } from '@office-kit/docx-editor';
  import Dialog from '../../../Dialog.svelte';
  import { getSession } from '../../../session.svelte';
  import { locale, t, type MessageKey } from '../../../i18n/index.svelte';
  import { common, selectionFormats } from './state.svelte';
  import { fromTwips, toTwips, TWIPS_PER_POINT, unitFor } from './units';

  const session = getSession();
  const ID = 'home.paragraph';
  // `w:line` for auto spacing is in 240ths of a line (§17.3.1.33).
  const LINE_UNIT = 240;
  const OUTLINE_LEVELS = 9;

  type LineMode = 'single' | 'onePointFive' | 'double' | 'atLeast' | 'exactly' | 'multiple' | '';
  type Special = 'none' | 'firstLine' | 'hanging' | '';

  const BREAK_TOGGLES: ReadonlyArray<{ toggle: ParagraphToggle; key: MessageKey }> = [
    { toggle: 'widowControl', key: 'home.para.widowControl' },
    { toggle: 'keepNext', key: 'home.para.keepNext' },
    { toggle: 'keepLines', key: 'home.para.keepLines' },
    { toggle: 'pageBreakBefore', key: 'home.para.pageBreakBefore' },
    { toggle: 'suppressLineNumbers', key: 'home.para.suppressLineNumbers' },
    { toggle: 'suppressAutoHyphens', key: 'home.para.suppressAutoHyphens' },
  ];
  // Word labels "Allow Latin text to wrap in the middle of a word", which is
  // `w:wordWrap` off; the checkbox is inverted for that one.
  const ASIAN_TOGGLES: ReadonlyArray<{ toggle: ParagraphToggle; key: MessageKey; inverted?: boolean }> = [
    { toggle: 'kinsoku', key: 'home.para.kinsoku' },
    { toggle: 'wordWrap', key: 'home.para.wordWrap', inverted: true },
    { toggle: 'overflowPunct', key: 'home.para.overflowPunct' },
    { toggle: 'topLinePunct', key: 'home.para.topLinePunct' },
    { toggle: 'autoSpaceDE', key: 'home.para.autoSpaceDE' },
    { toggle: 'autoSpaceDN', key: 'home.para.autoSpaceDN' },
  ];
  const TOGGLES = [...BREAK_TOGGLES, ...ASIAN_TOGGLES, { toggle: 'contextualSpacing' as const }, { toggle: 'mirrorIndents' as const }];

  interface Fields {
    alignment: string;
    outlineLevel: string;
    left: number;
    right: number;
    special: Special;
    by: number;
    before: number;
    after: number;
    lineMode: LineMode;
    at: number;
    textAlignment: string;
    toggles: Record<string, boolean | null>;
  }

  let tab = $state<'indents' | 'breaks' | 'asian'>('indents');
  let fields = $state<Fields | null>(null);
  let initial: Fields | null = null;
  const unit = $derived(unitFor(locale()));

  function lineModeOf(line: number | undefined, rule: string | undefined): LineMode {
    if (rule === 'exact') return 'exactly';
    if (rule === 'atLeast') return 'atLeast';
    if (line === undefined || line === LINE_UNIT) return 'single';
    if (line === LINE_UNIT * 1.5) return 'onePointFive';
    if (line === LINE_UNIT * 2) return 'double';
    return 'multiple';
  }

  function read(): Fields | null {
    const model = session.model;
    if (!model) return null;
    const paras = selectionFormats(model).paragraphs;
    const one = <T,>(f: (p: (typeof paras)[number]) => T): T | undefined => common(paras.map(f));
    const firstLine = one((p) => p.firstLine ?? 0) ?? 0;
    const hanging = one((p) => p.hanging ?? 0) ?? 0;
    const line = one((p) => p.line);
    const rule = one((p) => p.lineRule);
    const mode = lineModeOf(line, rule);
    const u = unitFor(locale());
    return {
      alignment: one((p) => p.alignment ?? 'left') ?? '',
      outlineLevel: String(one((p) => p.outlineLevel ?? OUTLINE_LEVELS) ?? ''),
      left: fromTwips(one((p) => p.left ?? 0) ?? 0, u),
      right: fromTwips(one((p) => p.right ?? 0) ?? 0, u),
      special: hanging > 0 ? 'hanging' : firstLine > 0 ? 'firstLine' : 'none',
      by: fromTwips(hanging || firstLine, u),
      before: (one((p) => p.before ?? 0) ?? 0) / TWIPS_PER_POINT,
      after: (one((p) => p.after ?? 0) ?? 0) / TWIPS_PER_POINT,
      lineMode: mode,
      at: mode === 'atLeast' || mode === 'exactly' ? (line ?? 0) / TWIPS_PER_POINT : (line ?? LINE_UNIT) / LINE_UNIT,
      textAlignment: one((p) => p.textAlignment ?? 'auto') ?? '',
      toggles: Object.fromEntries(TOGGLES.map(({ toggle }) => [toggle, one((p) => p.toggles[toggle]) ?? null])),
    };
  }

  $effect(() => {
    if (session.dialog !== ID) return;
    const loaded = read();
    initial = loaded && structuredClone(loaded);
    fields = loaded;
    tab = 'indents';
  });

  function lineSpacing(f: Fields): Pick<NonNullable<ParagraphPatch['spacing']>, 'line' | 'lineRule'> {
    switch (f.lineMode) {
      case 'single':
        return { line: LINE_UNIT, lineRule: 'auto' };
      case 'onePointFive':
        return { line: LINE_UNIT * 1.5, lineRule: 'auto' };
      case 'double':
        return { line: LINE_UNIT * 2, lineRule: 'auto' };
      case 'multiple':
        return { line: Math.round(f.at * LINE_UNIT), lineRule: 'auto' };
      case 'atLeast':
        return { line: Math.round(f.at * TWIPS_PER_POINT), lineRule: 'atLeast' };
      default:
        return { line: Math.round(f.at * TWIPS_PER_POINT), lineRule: 'exact' };
    }
  }

  function patch(f: Fields, from: Fields): ParagraphPatch {
    const out: { -readonly [K in keyof ParagraphPatch]: ParagraphPatch[K] } = {};
    if (f.alignment && f.alignment !== from.alignment) out.alignment = f.alignment as NonNullable<ParagraphPatch['alignment']>;
    if (f.outlineLevel && f.outlineLevel !== from.outlineLevel) {
      const level = Number(f.outlineLevel);
      out.outlineLevel = level >= OUTLINE_LEVELS ? null : level;
    }
    const indent: { left?: number; right?: number; firstLine?: number; hanging?: number } = {};
    if (f.left !== from.left) indent.left = toTwips(f.left, unit);
    if (f.right !== from.right) indent.right = toTwips(f.right, unit);
    if (f.special && (f.special !== from.special || f.by !== from.by)) {
      const by = toTwips(f.by, unit);
      indent.firstLine = f.special === 'firstLine' ? by : 0;
      indent.hanging = f.special === 'hanging' ? by : 0;
    }
    if (Object.keys(indent).length > 0) out.indent = indent;
    const spacing: { before?: number; after?: number; line?: number; lineRule?: 'auto' | 'exact' | 'atLeast' } = {};
    if (f.before !== from.before) spacing.before = Math.round(f.before * TWIPS_PER_POINT);
    if (f.after !== from.after) spacing.after = Math.round(f.after * TWIPS_PER_POINT);
    if (f.lineMode && (f.lineMode !== from.lineMode || f.at !== from.at)) Object.assign(spacing, lineSpacing(f));
    if (Object.keys(spacing).length > 0) out.spacing = spacing;
    const toggles: Partial<Record<ParagraphToggle, boolean>> = {};
    for (const { toggle } of TOGGLES) {
      const v = f.toggles[toggle];
      if (v !== null && v !== undefined && v !== from.toggles[toggle]) toggles[toggle] = v;
    }
    if (Object.keys(toggles).length > 0) out.toggles = toggles;
    if (f.textAlignment && f.textAlignment !== from.textAlignment) out.textAlignment = f.textAlignment;
    return out;
  }

  function ok(): void {
    if (!fields || !initial) return;
    const p = patch(fields, initial);
    if (Object.keys(p).length > 0) session.apply(commands.paragraphFormatCommand, p);
  }

  /** Tabs… applies what is set so far, then opens the Tabs dialog (as Word does). */
  function openTabs(): void {
    ok();
    session.openDialog('home.tabs');
  }

  function setToggle(toggle: ParagraphToggle, checked: boolean, inverted = false): void {
    if (fields) fields.toggles[toggle] = inverted ? !checked : checked;
  }

  function shown(toggle: ParagraphToggle, inverted = false): boolean | null {
    const v = fields?.toggles[toggle] ?? null;
    return v === null ? null : inverted ? !v : v;
  }
</script>

{#snippet check(toggle: ParagraphToggle, key: MessageKey, inverted = false)}
  {@const v = shown(toggle, inverted)}
  <label class="field"><input type="checkbox" checked={v === true} indeterminate={v === null} onchange={(e) => setToggle(toggle, (e.currentTarget as HTMLInputElement).checked, inverted)} />{t(key)}</label>
{/snippet}

<Dialog id={ID} title={t('home.para.dialog')} onok={ok}>
  {#if fields}
    <div class="dialog-tabs" role="tablist">
      <button type="button" role="tab" aria-selected={tab === 'indents'} class:active={tab === 'indents'} onclick={() => (tab = 'indents')}>{t('home.para.tabIndents')}</button>
      <button type="button" role="tab" aria-selected={tab === 'breaks'} class:active={tab === 'breaks'} onclick={() => (tab = 'breaks')}>{t('home.para.tabBreaks')}</button>
      <button type="button" role="tab" aria-selected={tab === 'asian'} class:active={tab === 'asian'} onclick={() => (tab = 'asian')}>{t('home.para.tabAsian')}</button>
    </div>
    {#if tab === 'indents'}
      <fieldset>
        <legend>{t('home.para.general')}</legend>
        <div class="form-grid">
          <label class="field">{t('home.para.alignment')}
            <select bind:value={fields.alignment}>
              <option value="" disabled></option>
              <option value="left">{t('tip.alignLeft')}</option>
              <option value="center">{t('tip.center')}</option>
              <option value="right">{t('tip.alignRight')}</option>
              <option value="both">{t('tip.justify')}</option>
              <option value="distribute">{t('home.distributed')}</option>
            </select>
          </label>
          <label class="field">{t('home.para.outlineLevel')}
            <select bind:value={fields.outlineLevel}>
              <option value="" disabled></option>
              <option value={String(OUTLINE_LEVELS)}>{t('home.para.bodyText')}</option>
              {#each Array.from({ length: OUTLINE_LEVELS }, (_, i) => i) as level (level)}<option value={String(level)}>{t('home.para.level')} {level + 1}</option>{/each}
            </select>
          </label>
        </div>
      </fieldset>
      <fieldset>
        <legend>{t('home.para.indentation')}</legend>
        <div class="form-grid">
          <label class="field">{t('home.para.left')}<input type="number" step="0.1" bind:value={fields.left} />{unit}</label>
          <label class="field">{t('home.para.right')}<input type="number" step="0.1" bind:value={fields.right} />{unit}</label>
          <label class="field">{t('home.para.special')}
            <select bind:value={fields.special}>
              <option value="" disabled></option>
              <option value="none">{t('home.para.none')}</option>
              <option value="firstLine">{t('home.para.firstLine')}</option>
              <option value="hanging">{t('home.para.hanging')}</option>
            </select>
          </label>
          <label class="field">{t('home.para.by')}<input type="number" min="0" step="0.1" bind:value={fields.by} disabled={fields.special === 'none'} />{unit}</label>
        </div>
        {@render check('mirrorIndents', 'home.para.mirror')}
      </fieldset>
      <fieldset>
        <legend>{t('home.para.spacing')}</legend>
        <div class="form-grid">
          <label class="field">{t('home.para.before')}<input type="number" min="0" step="6" bind:value={fields.before} />pt</label>
          <label class="field">{t('home.para.after')}<input type="number" min="0" step="6" bind:value={fields.after} />pt</label>
          <label class="field">{t('tip.lineSpacing')}
            <select bind:value={fields.lineMode}>
              <option value="" disabled></option>
              <option value="single">{t('home.para.single')}</option>
              <option value="onePointFive">{t('home.para.onePointFive')}</option>
              <option value="double">{t('home.para.double')}</option>
              <option value="atLeast">{t('home.para.atLeast')}</option>
              <option value="exactly">{t('home.para.exactly')}</option>
              <option value="multiple">{t('home.para.multiple')}</option>
            </select>
          </label>
          <label class="field">{t('home.para.at')}<input type="number" min="0" step={fields.lineMode === 'multiple' ? 0.25 : 1} bind:value={fields.at} disabled={!['atLeast', 'exactly', 'multiple'].includes(fields.lineMode)} />{fields.lineMode === 'atLeast' || fields.lineMode === 'exactly' ? 'pt' : ''}</label>
        </div>
        {@render check('contextualSpacing', 'home.para.contextual')}
      </fieldset>
    {:else if tab === 'breaks'}
      <fieldset>
        <legend>{t('home.para.pagination')}</legend>
        {#each BREAK_TOGGLES as b (b.toggle)}{@render check(b.toggle, b.key)}{/each}
      </fieldset>
    {:else}
      <fieldset>
        <legend>{t('home.para.lineBreak')}</legend>
        {#each ASIAN_TOGGLES as a (a.toggle)}{@render check(a.toggle, a.key, a.inverted)}{/each}
      </fieldset>
      <label class="field">{t('home.para.textAlignment')}
        <select bind:value={fields.textAlignment}>
          <option value="" disabled></option>
          <option value="auto">{t('home.para.auto')}</option>
          <option value="top">{t('home.para.top')}</option>
          <option value="center">{t('home.para.centered')}</option>
          <option value="baseline">{t('home.para.baseline')}</option>
          <option value="bottom">{t('home.para.bottom')}</option>
        </select>
      </label>
    {/if}
    <div><button type="button" class="push" onclick={openTabs}>{t('home.para.tabs')}</button></div>
  {/if}
</Dialog>
