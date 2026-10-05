<script lang="ts">
  /** A length spinner in the locale's unit (in / cm / mm) whose value is twips. */
  import { t } from '../../i18n/index.svelte';
  import { lengthUnit, twipsToUnit, unitStep, unitToTwips } from './table-tool.svelte';

  type Props = { label: string; twips: number; disabled?: boolean; onchange?: (twips: number) => void };
  let { label, twips = $bindable(), disabled = false, onchange }: Props = $props();
  const unit = $derived(lengthUnit());

  function input(e: Event): void {
    if (!(e.currentTarget instanceof HTMLInputElement)) return;
    const value = Number(e.currentTarget.value);
    if (!Number.isFinite(value) || value < 0) return;
    twips = unitToTwips(value, unit);
    onchange?.(twips);
  }
</script>

<label class="field"
  >{label}<input type="number" min="0" step={unitStep(unit)} value={twipsToUnit(twips, unit)} {disabled} onchange={input} />{t(`tbl.unit.${unit}`)}</label
>
