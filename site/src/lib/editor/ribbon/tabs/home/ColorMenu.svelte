<script lang="ts">
  /**
   * Word's colour drop-down (Font Color, Shading, Underline Color, Borders):
   * Automatic / No Color, the Theme Colors grid of the document's theme,
   * Standard Colors, and More Colors…. Theme picks keep their theme reference
   * (`themeColor` + tint / shade), so they follow a theme change as in Word.
   */
  import type { ColorValue } from '@office-kit/docx';
  import { themeColorGrid, themePalette, type ThemeSwatch } from '@office-kit/docx-editor';
  import { getSession } from '../../../session.svelte';
  import { t, type MessageKey } from '../../../i18n/index.svelte';
  import { homeState } from './state.svelte';

  type Props = {
    /** Label of the "no colour" entry: Automatic (text) or No Color (fills). */
    none: 'auto' | 'noColor';
    /** `undefined` is Automatic / No Color. */
    onpick: (color: ColorValue | undefined) => void;
  };
  const { none, onpick }: Props = $props();
  const session = getSession();
  const home = homeState(session);

  // Word's "Standard Colors" row.
  const STANDARD_COLORS = ['C00000', 'FF0000', 'FFC000', 'FFFF00', '92D050', '00B050', '00B0F0', '0070C0', '002060', '7030A0'];
  const STANDARD_ROWS = 6;

  const grid = $derived(session.version >= 0 && session.model ? themeColorGrid(themePalette(session.model.doc)) : []);
  const rows = $derived(Array.from({ length: STANDARD_ROWS }, (_, r) => grid.map((col) => col[r]).filter((s) => s !== undefined)));

  function name(swatch: ThemeSwatch): string {
    const base = t(`home.color.${swatch.themeColor}` as MessageKey);
    if (swatch.percent === 0) return base;
    const amount = t(swatch.percent > 0 ? 'home.color.lighter' : 'home.color.darker').replace('{n}', String(Math.abs(swatch.percent)));
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

  function more(): void {
    home.moreColors = { initial: '000000', pick: onpick };
    session.openDialog('home.moreColors');
  }
</script>

<button class="mi" role="menuitem" onclick={() => onpick(undefined)}>
  <i class={none === 'auto' ? 'auto-chip' : 'none-chip'}></i>{none === 'auto' ? t('color.auto') : t('highlight.none')}
</button>
<div class="menu-head">{t('home.color.theme')}</div>
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
<hr />
<button class="mi" role="menuitem" onclick={more}>{t('color.more')}</button>
