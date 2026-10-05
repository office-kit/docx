<script lang="ts">
  /**
   * Word's split button: the face applies the last choice, the chevron opens a
   * menu. Without `onclick` the whole control opens the menu (a drop-down
   * button, like Layout ▸ Margins). Only one menu is open at a time; the
   * session tracks which.
   */
  import type { Snippet } from 'svelte';
  import RibbonIcon, { type IconName } from '../RibbonIcon.svelte';
  import { getSession } from '../session.svelte';

  type Props = {
    /** Unique id of this menu within the page. */
    id: string;
    tip: string;
    menu: Snippet;
    onclick?: () => void;
    icon?: IconName;
    /** Custom face (e.g. a color swatch icon) instead of `icon`. */
    face?: Snippet;
    label?: string;
    size?: 'small' | 'mid' | 'large';
    on?: boolean;
    disabled?: boolean;
    /** Open the menu toward the left (for controls near the right edge). */
    alignRight?: boolean;
  };
  const { id, tip, menu, onclick, icon, face, label, size = 'small', on = false, disabled = false, alignRight = false }: Props = $props();
  const session = getSession();
  const open = $derived(session.openMenu === id);
  const toggle = (): void => session.toggleMenu(id);
</script>

<span class="split" class:open>
  {#if size === 'large'}
    <button class="big" class:on {disabled} onclick={onclick ?? toggle} title={tip} aria-haspopup={onclick ? undefined : 'menu'} aria-expanded={onclick ? undefined : open}>
      {#if face}{@render face()}{:else if icon}<RibbonIcon name={icon} size={32} />{/if}
      <span class="caret">{label ?? tip}{#if !onclick}<RibbonIcon name="chevronDown" size={10} />{/if}</span>
    </button>
    {#if onclick}<button class="arrow" onclick={toggle} {disabled} aria-label={tip} aria-expanded={open}><RibbonIcon name="chevronDown" size={10} /></button>{/if}
  {:else}
    <button class={size === 'mid' ? 'mid' : 'rb'} class:on {disabled} onclick={onclick ?? toggle} title={tip} aria-label={size === 'mid' ? undefined : tip} aria-haspopup={onclick ? undefined : 'menu'} aria-expanded={onclick ? undefined : open}>
      {#if face}{@render face()}{:else if icon}<RibbonIcon name={icon} size={size === 'mid' ? 16 : 20} />{/if}
      {#if size === 'mid'}<span>{label ?? tip}</span>{/if}
    </button>
    <button class="arrow" onclick={toggle} {disabled} aria-label={tip} aria-expanded={open}><RibbonIcon name="chevronDown" size={10} /></button>
  {/if}
  {#if open}
    <div class="menu" class:right={alignRight} role="menu">{@render menu()}</div>
  {/if}
</span>
