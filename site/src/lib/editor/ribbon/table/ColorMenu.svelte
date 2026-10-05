<script lang="ts">
  /**
   * Word's color drop-down (Shading, Pen Color): the theme palette with its
   * tints and shades, the standard colors, and More Colors… (the system
   * picker). `auto` is the first entry: No Color for fills, Automatic for lines.
   */
  import { t } from '../../i18n/index.svelte';
  import { STANDARD_COLORS, themePalette } from './table-tool.svelte';

  type Props = { autoLabel: 'tbl.color.none' | 'tbl.color.automatic'; onpick: (color: string) => void };
  const { autoLabel, onpick }: Props = $props();
  const palette = themePalette();
  const rows = palette[0]?.map((_, i) => palette.map((column) => column[i] ?? '')) ?? [];

  function custom(e: Event): void {
    const input = e.currentTarget;
    if (input instanceof HTMLInputElement) onpick(input.value.slice(1).toUpperCase());
  }
</script>

<button class="mi" role="menuitem" onclick={() => onpick('auto')}
  ><i class="auto-chip" class:none={autoLabel === 'tbl.color.none'}></i>{t(autoLabel)}</button
>
<div class="menu-head">{t('tbl.color.theme')}</div>
<div class="grid ten">
  {#each rows as row, i (i)}
    {#each row as hex, j (j)}
      <button class="chip" class:gap={i === 1} role="menuitem" style="background:#{hex}" title={`#${hex}`} aria-label={`#${hex}`} onclick={() => onpick(hex)}></button>
    {/each}
  {/each}
</div>
<div class="menu-head">{t('tbl.color.standard')}</div>
<div class="grid ten">
  {#each STANDARD_COLORS as hex (hex)}
    <button class="chip" role="menuitem" style="background:#{hex}" title={`#${hex}`} aria-label={`#${hex}`} onclick={() => onpick(hex)}></button>
  {/each}
</div>
<hr />
<label class="mi">{t('tbl.color.more')}<input type="color" onchange={custom} class="wk-tbl-color-input" /></label>

<style>
  .chip.gap {
    margin-top: 4px;
  }
  .none {
    background: #fff !important;
    background-image: linear-gradient(to top right, transparent 45%, #d00 45%, #d00 55%, transparent 55%) !important;
  }
  .wk-tbl-color-input {
    width: 0;
    height: 0;
    padding: 0;
    border: 0;
    opacity: 0;
  }
</style>
