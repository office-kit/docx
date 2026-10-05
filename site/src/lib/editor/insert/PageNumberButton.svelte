<script lang="ts">
  /**
   * Insert ▸ Page Number: Top of Page / Bottom of Page / Page Margins, each
   * with Word's plain-number alignments and "Page X of Y"; Current Position
   * (a PAGE field at the caret); Format Page Numbers…; Remove Page Numbers.
   */
  import { commands } from '@office-kit/docx-editor';
  import SplitButton from '../ribbon/SplitButton.svelte';
  import { getSession } from '../session.svelte';
  import { t } from '../i18n/index.svelte';
  import { DIALOG } from './state.svelte';

  const session = getSession();
  const POSITIONS = [
    { position: 'top', key: 'ins.pn.top' },
    { position: 'bottom', key: 'ins.pn.bottom' },
    { position: 'margin', key: 'ins.pn.margins' },
  ] as const;
  const ALIGNS = [
    { align: 'left', key: 'ins.pn.left' },
    { align: 'center', key: 'ins.pn.center' },
    { align: 'right', key: 'ins.pn.right' },
  ] as const;
  const hasCaret = $derived(session.tick >= 0 && !!session.model?.selection);
</script>

<SplitButton id="insert.pageNumber" size="large" icon="pageNumber" tip={t('ins.pageNumber')}>
  {#snippet menu()}
    {#each POSITIONS as p (p.position)}
      <div class="menu-head">{t(p.key)}</div>
      {#each ALIGNS as a (a.align)}
        <button class="mi" role="menuitem" onclick={() => session.apply(commands.insertPageNumberCommand, { position: p.position, align: a.align })}>
          {t('ins.pn.plain')} — {t(a.key)}
        </button>
      {/each}
      {#if p.position !== 'margin'}
        <button class="mi" role="menuitem" onclick={() => session.apply(commands.insertPageNumberCommand, { position: p.position, align: 'right', style: 'pageXofY' })}>
          {t('ins.pn.pageXofY')}
        </button>
      {/if}
    {/each}
    <hr />
    <button class="mi" role="menuitem" disabled={!hasCaret} onclick={() => session.apply(commands.insertFieldAtCaretCommand, { instruction: 'PAGE', context: { page: session.currentPage } })}>
      {t('ins.pn.current')}
    </button>
    <button class="mi" role="menuitem" onclick={() => session.openDialog(DIALOG.pageNumberFormat)}>{t('ins.pn.format')}</button>
    <button class="mi" role="menuitem" onclick={() => session.apply(commands.removePageNumbersCommand, undefined)}>{t('ins.pn.remove')}</button>
  {/snippet}
</SplitButton>
