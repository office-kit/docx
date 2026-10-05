<script lang="ts">
  /** Insert ▸ Icons: pick one or more icons from the built-in library. */
  import Dialog from '../Dialog.svelte';
  import { getSession } from '../session.svelte';
  import { t, type MessageKey } from '../i18n/index.svelte';
  import { ICON_LIBRARY, iconSvg, type LibraryIcon } from './icon-library';
  import { insertIcons } from './insert';

  const session = getSession();
  const RASTER_PX = 512;
  const CATEGORIES: ReadonlyArray<readonly [LibraryIcon['category'], MessageKey]> = [
    ['people', 'iconLib.people'],
    ['objects', 'iconLib.objects'],
    ['symbols', 'iconLib.symbols'],
    ['nature', 'iconLib.nature'],
  ];
  let picked = $state<ReadonlySet<string>>(new Set());

  /** "arrowRight" → "Arrow right": the icon's name and default alt text. */
  const label = (id: string): string => {
    const words = id.replace(/([A-Z])/g, ' $1').toLowerCase();
    return words.charAt(0).toUpperCase() + words.slice(1);
  };

  function toggle(id: string): void {
    const next = new Set(picked);
    if (!next.delete(id)) next.add(id);
    picked = next;
  }

  function ok(): void {
    const icons = ICON_LIBRARY.filter((i) => picked.has(i.id));
    picked = new Set();
    insertIcons(session, icons.map((icon) => ({ svg: iconSvg(icon, RASTER_PX), name: label(icon.id) }))).catch((err: unknown) => {
      session.status = `${t('ill.icons')}: ${(err as Error).message}`;
    });
  }
</script>

<Dialog id="picture.icons" title={t('ill.icons')} onok={ok} okLabel={t('iconLib.insert')} okDisabled={picked.size === 0}>
  {#each CATEGORIES as [category, key] (category)}
    <div class="menu-head">{t(key)}</div>
    <div class="icons-grid">
      {#each ICON_LIBRARY.filter((i) => i.category === category) as icon (icon.id)}
        <button type="button" class:picked={picked.has(icon.id)} aria-pressed={picked.has(icon.id)} title={label(icon.id)} aria-label={label(icon.id)} onclick={() => toggle(icon.id)}>
          <svg viewBox="0 0 24 24" width="32" height="32" aria-hidden="true"><path d={icon.path} fill="currentColor" /></svg>
        </button>
      {/each}
    </div>
  {/each}
</Dialog>

<style>
  .icons-grid {
    display: grid;
    grid-template-columns: repeat(8, 48px);
    gap: 6px;
  }
  .icons-grid button {
    display: grid;
    place-items: center;
    height: 48px;
    border: 1px solid var(--chrome-line);
    border-radius: 4px;
    background: #fff;
    color: #333;
  }
  .icons-grid button.picked {
    border: 2px solid #2b579a;
    background: #e8eef8;
  }
</style>
