<script lang="ts">
  /**
   * Layout ▸ Columns ▸ More Columns…: presets, number of columns, width and
   * spacing per column, equal column width, line between, and Apply to
   * (this section, whole document, or this point forward — a continuous
   * section break at the caret).
   */
  import { commands, type SectionTarget } from '@office-kit/docx-editor';
  import type { ColumnDefinition, SectionColumns } from '@office-kit/docx';
  import Dialog from '../Dialog.svelte';
  import LengthField from './LengthField.svelte';
  import { getSession } from '../session.svelte';
  import { locale, t, type MessageKey } from '../i18n/index.svelte';
  import { lengthUnitFor } from './units';
  import { columnPreset, columnPresetOf, type ColumnPresetId } from './presets';

  const ID = 'layout.columns';
  const session = getSession();
  const unit = $derived(lengthUnitFor(locale()));
  const PRESETS: ReadonlyArray<{ id: ColumnPresetId; key: MessageKey }> = [
    { id: 'one', key: 'lay.cols.one' },
    { id: 'two', key: 'lay.cols.two' },
    { id: 'three', key: 'lay.cols.three' },
    { id: 'left', key: 'lay.cols.left' },
    { id: 'right', key: 'lay.cols.right' },
  ];
  const MAX_COLUMNS = 45;
  // Word keeps each column at least 0.5 in wide.
  const MIN_COLUMN = 720;

  let open = false;
  let textWidth = $state(9360);
  let count = $state(1);
  let equal = $state(true);
  let separator = $state(false);
  let space = $state(720);
  let widths = $state<ColumnDefinition[]>([]);
  let applyTo = $state<SectionTarget>('section');

  $effect(() => {
    const showing = session.dialog === ID;
    if (showing && !open && session.model) {
      const p = commands.currentSectionProperties(session.model);
      textWidth = p.pageSize.widthTwips - p.margins.left - p.margins.right - p.margins.gutter;
      load(p.columns);
      applyTo = 'section';
    }
    open = showing;
  });

  function load(c: SectionColumns): void {
    count = c.count;
    separator = c.separator;
    space = c.spaceTwips;
    equal = !c.columns;
    widths = c.columns ? c.columns.map((col) => ({ ...col })) : equalWidths(c.count, c.spaceTwips);
  }

  function equalWidths(n: number, gap: number): ColumnDefinition[] {
    const w = Math.floor((textWidth - gap * (n - 1)) / n);
    return Array.from({ length: n }, (_, i) => (i < n - 1 ? { widthTwips: w, spaceTwips: gap } : { widthTwips: w }));
  }

  function setCount(n: number): void {
    if (!Number.isInteger(n) || n < 1 || n > MAX_COLUMNS) return;
    count = n;
    if (n === 1) equal = true;
    widths = equalWidths(n, space);
  }

  function setWidth(i: number, w: number): void {
    widths = widths.map((c, j) => (j === i ? { ...c, widthTwips: w } : c));
  }

  function setSpace(i: number, s: number): void {
    widths = widths.map((c, j) => (j === i ? { ...c, spaceTwips: s } : c));
  }

  const preset = $derived(columnPresetOf(current()));

  function current(): SectionColumns {
    return equal ? { count, spaceTwips: space, separator } : { count, spaceTwips: space, separator, columns: widths };
  }

  function ok(): boolean | void {
    const columns = current();
    if (columns.columns?.some((c) => c.widthTwips < MIN_COLUMN)) return false;
    session.apply(commands.columnsCommand, { columns, target: applyTo });
  }
</script>

<Dialog id={ID} title={t('lay.cols.title')} onok={ok}>
  <fieldset>
    <legend>{t('lay.cols.presets')}</legend>
    <div class="row">
      {#each PRESETS as p (p.id)}
        <button type="button" class="preset" class:selected={preset === p.id} aria-pressed={preset === p.id} onclick={() => load({ ...columnPreset(p.id, textWidth), separator })}>
          <span class="col-icon col-icon-{p.id}"></span>{t(p.key)}
        </button>
      {/each}
    </div>
  </fieldset>
  <div class="form-grid">
    <label for="cols-n">{t('lay.cols.number')}</label>
    <input id="cols-n" type="number" min="1" max={MAX_COLUMNS} value={count} onchange={(e) => setCount(Number((e.currentTarget as HTMLInputElement).value))} />
  </div>
  <fieldset>
    <legend>{t('lay.cols.widthSpacing')}</legend>
    <div class="cols-table">
      <span class="muted">{t('lay.cols.col')}</span><span class="muted">{t('lay.cols.width')}</span><span class="muted">{t('lay.cols.spacing')}</span>
      {#each widths as col, i (i)}
        <span>{i + 1}:</span>
        <LengthField value={equal ? Math.floor((textWidth - space * (count - 1)) / count) : col.widthTwips} {unit} label={`${t('lay.cols.width')} ${i + 1}`} min={MIN_COLUMN} disabled={equal && i > 0} onchange={(v) => (equal ? undefined : setWidth(i, v))} />
        {#if i < widths.length - 1}
          <LengthField value={equal ? space : (col.spaceTwips ?? space)} {unit} label={`${t('lay.cols.spacing')} ${i + 1}`} disabled={equal && i > 0} onchange={(v) => (equal ? (space = v) : setSpace(i, v))} />
        {:else}<span></span>{/if}
      {/each}
    </div>
    <label class="check"><input type="checkbox" bind:checked={equal} disabled={count === 1} onchange={() => (widths = equalWidths(count, space))} />{t('lay.cols.equal')}</label>
  </fieldset>
  <label class="check"><input type="checkbox" bind:checked={separator} disabled={count === 1} />{t('lay.cols.lineBetween')}</label>
  <div class="field apply-to">
    <label for="cols-apply">{t('dsn.applyTo')}</label>
    <select id="cols-apply" bind:value={applyTo}>
      <option value="section">{t('dsn.applyTo.section')}</option>
      <option value="document">{t('dsn.applyTo.document')}</option>
      <option value="forward">{t('lay.forward')}</option>
    </select>
  </div>
</Dialog>
