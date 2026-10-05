<script lang="ts">
  /** The "Separate text at / with" choice shared by Convert to Text and Convert Text to Table. */
  import type { TableTextSeparator } from '@office-kit/docx';
  import { t, type MessageKey } from '../../i18n/index.svelte';

  type Props = { legend: MessageKey; paragraphLabel: MessageKey; separator: TableTextSeparator };
  let { legend, paragraphLabel, separator = $bindable() }: Props = $props();

  // Word's default "Other" character.
  let other = $state('-');
  const kind = $derived(typeof separator === 'string' ? separator : 'other');
  const CHOICES: readonly { value: 'paragraph' | 'tab' | 'comma'; label: MessageKey | null }[] = [
    { value: 'paragraph', label: null },
    { value: 'tab', label: 'tbl.dlg.tabs' },
    { value: 'comma', label: 'tbl.dlg.commas' },
  ];
</script>

<fieldset class="tbl-sep">
  <legend>{t(legend)}</legend>
  {#each CHOICES as c (c.value)}
    <label class="field"><input type="radio" name={legend} checked={kind === c.value} onchange={() => (separator = c.value)} />{t(c.label ?? paragraphLabel)}</label>
  {/each}
  <label class="field"
    ><input type="radio" name={legend} checked={kind === 'other'} onchange={() => (separator = { other })} />{t('tbl.dlg.other')}<input
      type="text"
      maxlength="1"
      size="2"
      bind:value={other}
      oninput={() => (separator = { other })}
    /></label
  >
</fieldset>

<style>
  .tbl-sep {
    display: grid;
    grid-template-columns: auto auto;
    gap: 6px 16px;
  }
</style>
