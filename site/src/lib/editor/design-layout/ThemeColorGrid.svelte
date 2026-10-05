<script lang="ts" module>
  /** A colour picked from Word's palette: RGB plus the theme slot it came from, if any. */
  export interface PickedColor {
    readonly rgb: string;
    readonly themeColor?: string;
    readonly themeShade?: string;
    readonly themeTint?: string;
  }
</script>

<script lang="ts">
  /**
   * Word's colour palette as menu content: Theme Colors (10 × 6 grid) and
   * Standard Colors, plus a More Colors… picker. Used by Page Color and the
   * colour drop-downs of the Design dialogs.
   */
  import { resolveTheme, themeSchemePalette } from '@office-kit/docx-editor';
  import { getSession } from '../session.svelte';
  import { t } from '../i18n/index.svelte';
  import { STANDARD_COLORS } from './presets';

  type Props = { onpick: (color: PickedColor) => void };
  const { onpick }: Props = $props();
  const session = getSession();
  const palette = $derived(session.tick >= 0 && session.model ? themeSchemePalette(resolveTheme(session.model.doc).colors) : []);
</script>

<div class="menu-head">{t('dsn.color.theme')}</div>
<div class="grid ten">
  {#each palette as row, r (r)}
    {#each row as cell (`${r}-${cell.themeColor}`)}
      <button
        type="button"
        class="chip"
        role="menuitem"
        style="background:#{cell.rgb}"
        title={`#${cell.rgb}`}
        aria-label={`#${cell.rgb}`}
        onclick={() => onpick(cell)}
      ></button>
    {/each}
  {/each}
</div>
<div class="menu-head">{t('dsn.color.standard')}</div>
<div class="grid ten">
  {#each STANDARD_COLORS as hex (hex)}
    <button type="button" class="chip" role="menuitem" style="background:#{hex}" title={`#${hex}`} aria-label={`#${hex}`} onclick={() => onpick({ rgb: hex })}></button>
  {/each}
</div>
<label class="mi">
  {t('dsn.color.more')}
  <input type="color" hidden onchange={(e) => onpick({ rgb: (e.currentTarget as HTMLInputElement).value.slice(1).toUpperCase() })} />
</label>
