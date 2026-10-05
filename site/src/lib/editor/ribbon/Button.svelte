<script lang="ts">
  /**
   * A ribbon button in one of Word's three sizes:
   * - `small`: a 28 px icon (or text glyph like **B**) in a ribbon row;
   * - `mid`: a small icon with its label beside it (Insert ▸ Field);
   * - `large`: a 32 px icon over its label (Insert ▸ Table).
   * `on` draws the pressed state of a toggle.
   */
  import RibbonIcon, { type IconName } from '../RibbonIcon.svelte';

  type Props = {
    tip: string;
    onclick: () => void;
    icon?: IconName;
    /** Trusted markup drawn instead of an icon (e.g. `x<sub>2</sub>`). */
    glyph?: string;
    glyphClass?: string;
    /** Visible label (`mid` / `large`); defaults to the tooltip. */
    label?: string;
    size?: 'small' | 'mid' | 'large';
    on?: boolean;
    disabled?: boolean;
  };
  const { tip, onclick, icon, glyph, glyphClass = '', label, size = 'small', on = false, disabled = false }: Props = $props();
</script>

{#if size === 'small'}
  <button class="rb {glyph === undefined ? '' : `glyph ${glyphClass}`}" class:on {disabled} {onclick} title={tip} aria-label={tip} aria-pressed={on}
    >{#if glyph !== undefined}{@html glyph}{:else if icon}<RibbonIcon name={icon} />{/if}</button
  >
{:else if size === 'mid'}
  <button class="mid" class:on {disabled} {onclick} title={tip} aria-pressed={on}
    >{#if icon}<RibbonIcon name={icon} size={16} />{/if}<span>{label ?? tip}</span></button
  >
{:else}
  <button class="big" class:on {disabled} {onclick} title={tip} aria-pressed={on}
    >{#if icon}<RibbonIcon name={icon} size={32} />{/if}<span>{label ?? tip}</span></button
  >
{/if}
