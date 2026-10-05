<script lang="ts">
  /**
   * Word's colour drop-down for shapes and pens: Theme Colors grid, Standard
   * Colors, an optional "No Fill" / "No Outline" entry, More Colors…, and any
   * extra entries (Weight, Dashes, Gradient …) passed as children.
   */
  import type { Snippet } from 'svelte';
  import { themeColors } from '@office-kit/docx-editor';
  import { getSession } from '../session.svelte';
  import { t } from '../i18n/index.svelte';
  import { STANDARD_COLORS, themeGrid } from './palette';

  type Props = {
    onpick: (hex: string) => void;
    /** Label of the "no colour" entry (No Fill / No Outline); omitted when absent. */
    noneLabel?: string;
    onnone?: () => void;
    children?: Snippet;
  };
  const { onpick, noneLabel, onnone, children }: Props = $props();
  const session = getSession();
  const grid = $derived(session.model && session.version >= 0 ? themeGrid(themeColors(session.model.doc)) : []);
</script>

<div class="menu-head">{t('draw.themeColors')}</div>
<div class="wk-color-grid">
  {#each grid as row, r (r)}
    {#each row as hex, c (c)}
      <button class="chip" class:first-row={r === 0} role="menuitem" style="background:#{hex}" title={`#${hex}`} aria-label={`#${hex}`} onclick={() => onpick(hex)}></button>
    {/each}
  {/each}
</div>
<div class="menu-head">{t('draw.standardColors')}</div>
<div class="grid ten">
  {#each STANDARD_COLORS as hex (hex)}
    <button class="chip" role="menuitem" style="background:#{hex}" title={`#${hex}`} aria-label={`#${hex}`} onclick={() => onpick(hex)}></button>
  {/each}
</div>
{#if noneLabel && onnone}
  <button class="mi" role="menuitem" onclick={onnone}><span class="wk-none-chip"></span>{noneLabel}</button>
{/if}
<label class="mi">
  {t('draw.moreColors')}
  <input type="color" onchange={(e) => onpick((e.currentTarget as HTMLInputElement).value.slice(1).toUpperCase())} hidden />
</label>
{#if children}<hr />{@render children()}{/if}
