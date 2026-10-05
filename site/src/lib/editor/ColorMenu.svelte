<script lang="ts">
  /**
   * Word's colour drop-down, shared by every colour picker on the ribbon:
   * an optional Automatic / No Color entry, the Theme Colors grid of the
   * document's theme, Standard Colors, an optional No Fill / No Outline entry,
   * More Colors…, and any extra entries passed as children.
   *
   * Theme picks carry their theme reference (`themeColor` + tint / shade) so a
   * caller writing a WordprocessingML colour can keep it and follow theme
   * changes as Word does; callers that only store RGB read `rgb`.
   */
  import type { Snippet } from 'svelte';
  import type { ColorValue } from '@office-kit/docx';
  import { themeColorGrid, themePalette, type ThemeSwatch } from '@office-kit/docx-editor';
  import { getSession } from './session.svelte';
  import { t, type MessageKey } from './i18n/index.svelte';

  type Entry = { label: string; onpick: () => void };
  type Props = {
    onpick: (color: ColorValue) => void;
    /** Automatic (black chip) or No Color (struck-out chip), above the grid. */
    auto?: Entry & { chip: 'auto' | 'none' };
    /** No Fill / No Outline, below the standard colours. */
    none?: Entry;
    /** Opens a More Colors dialog; without it More Colors… is the browser's colour picker. */
    onmore?: () => void;
    children?: Snippet;
  };
  const { onpick, auto, none, onmore, children }: Props = $props();
  const session = getSession();

  const STANDARD_COLORS = ['C00000', 'FF0000', 'FFC000', 'FFFF00', '92D050', '00B050', '00B0F0', '0070C0', '002060', '7030A0'];
  const GRID_ROWS = 6;

  const grid = $derived(session.version >= 0 && session.model ? themeColorGrid(themePalette(session.model.doc)) : []);
  const rows = $derived(Array.from({ length: GRID_ROWS }, (_, r) => grid.map((col) => col[r]).filter((s) => s !== undefined)));

  function name(swatch: ThemeSwatch): string {
    const base = t(`color.${swatch.themeColor}` as MessageKey);
    if (swatch.percent === 0) return base;
    const amount = t(swatch.percent > 0 ? 'color.lighter' : 'color.darker').replace('{n}', String(Math.abs(swatch.percent)));
    return `${base}, ${amount}`;
  }

  function pickSwatch(s: ThemeSwatch): void {
    onpick({
      rgb: s.rgb,
      themeColor: s.themeColor,
      ...(s.themeTint === undefined ? {} : { themeTint: s.themeTint }),
      ...(s.themeShade === undefined ? {} : { themeShade: s.themeShade }),
    });
  }

  function custom(e: Event & { currentTarget: HTMLInputElement }): void {
    onpick({ rgb: e.currentTarget.value.slice(1).toUpperCase() });
  }
</script>

{#if auto}
  <button class="mi" role="menuitem" onclick={auto.onpick}><i class={auto.chip === 'auto' ? 'auto-chip' : 'none-chip'}></i>{auto.label}</button>
{/if}
<div class="menu-head">{t('color.theme')}</div>
<div class="theme-grid">
  {#each rows as row, r (r)}
    <div class="grid ten" class:base-row={r === 0}>
      {#each row as s (s.themeColor + s.percent)}
        <button class="chip" role="menuitem" style="background:#{s.rgb}" title={name(s)} aria-label={name(s)} onclick={() => pickSwatch(s)}></button>
      {/each}
    </div>
  {/each}
</div>
<div class="menu-head">{t('color.standard')}</div>
<div class="grid ten">
  {#each STANDARD_COLORS as hex (hex)}
    <button class="chip" role="menuitem" style="background:#{hex}" title={`#${hex}`} aria-label={`#${hex}`} onclick={() => onpick({ rgb: hex })}></button>
  {/each}
</div>
{#if none}
  <button class="mi" role="menuitem" onclick={none.onpick}><i class="none-chip"></i>{none.label}</button>
{/if}
<hr />
{#if onmore}
  <button class="mi" role="menuitem" onclick={onmore}>{t('color.more')}</button>
{:else}
  <label class="mi">{t('color.more')}<input type="color" class="picker" onchange={custom} /></label>
{/if}
{#if children}<hr />{@render children()}{/if}

<style>
  /* Not `hidden`: some browsers do not open the picker of an undisplayed input. */
  .picker {
    width: 0;
    height: 0;
    padding: 0;
    border: 0;
    opacity: 0;
  }
</style>
