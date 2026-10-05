<script lang="ts">
  /**
   * Word's measurement box with spinner arrows: shows twips in the given unit
   * (`1"`, `2.54 cm`, `6 pt`), accepts any unit typed with the number, and
   * reports the new value in twips. Out-of-range or unreadable input puts
   * the box back without reporting anything.
   */
  import RibbonIcon from '../RibbonIcon.svelte';
  import { formatLength, parseLength, spinStep, type LengthUnit } from './units';

  type Props = {
    value: number | undefined;
    unit: LengthUnit;
    onchange: (twips: number) => void;
    label: string;
    min?: number;
    max?: number;
    disabled?: boolean;
    /** Text shown when `value` is undefined (e.g. "Auto"). */
    placeholder?: string;
    id?: string;
  };
  const { value, unit, onchange, label, min = 0, max = 31680, disabled = false, placeholder = '', id }: Props = $props();
  let input = $state<HTMLInputElement | null>(null);

  function shown(): string {
    return value === undefined ? '' : formatLength(value, unit);
  }

  function commit(next: number | undefined): void {
    if (next === undefined || next < min || next > max) {
      if (input) input.value = shown();
      return;
    }
    onchange(next);
    if (input) input.value = formatLength(next, unit);
  }

  function spin(direction: 1 | -1): void {
    const step = spinStep(unit);
    const base = value ?? min;
    // Snap to the step grid first, as Word's arrows do.
    const snapped = direction > 0 ? Math.floor(base / step) * step + step : Math.ceil(base / step) * step - step;
    commit(Math.min(max, Math.max(min, snapped)));
  }
</script>

<span class="spin">
  <input
    bind:this={input}
    {id}
    type="text"
    value={shown()}
    {placeholder}
    {disabled}
    aria-label={label}
    title={label}
    onchange={(e) => commit(parseLength((e.currentTarget as HTMLInputElement).value, unit))}
    onkeydown={(e) => {
      if (e.key === 'ArrowUp' || e.key === 'ArrowDown') {
        e.preventDefault();
        spin(e.key === 'ArrowUp' ? 1 : -1);
      }
    }}
  />
  <span class="spin-arrows">
    <button type="button" tabindex="-1" {disabled} aria-label={`${label} +`} onclick={() => spin(1)}><RibbonIcon name="chevronDown" size={8} /></button>
    <button type="button" tabindex="-1" {disabled} aria-label={`${label} −`} onclick={() => spin(-1)}><RibbonIcon name="chevronDown" size={8} /></button>
  </span>
</span>
