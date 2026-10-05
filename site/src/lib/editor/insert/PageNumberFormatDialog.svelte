<script lang="ts">
  /**
   * Format Page Numbers (`w:pgNumType`): number format, chapter number
   * (heading level + separator) and continue / start at.
   */
  import { CHAPTER_SEPARATORS, getPageNumberFormat, PAGE_NUMBER_FORMATS, type ChapterSeparator, type PageNumberFormat } from '@office-kit/docx';
  import { commands } from '@office-kit/docx-editor';
  import Dialog from '../Dialog.svelte';
  import { getSession } from '../session.svelte';
  import { t } from '../i18n/index.svelte';
  import { DIALOG } from './state.svelte';

  const session = getSession();
  // How each format shows 1, 2, 3 in Word's list.
  const SAMPLES: Record<PageNumberFormat, string> = {
    decimal: '1, 2, 3, …',
    upperRoman: 'I, II, III, …',
    lowerRoman: 'i, ii, iii, …',
    upperLetter: 'A, B, C, …',
    lowerLetter: 'a, b, c, …',
    numberInDash: '- 1 -, - 2 -, - 3 -, …',
    decimalFullWidth: '１, ２, ３, …',
    japaneseCounting: '一, 二, 三, …',
    aiueoFullWidth: 'ア, イ, ウ, …',
    iroha: 'イ, ロ, ハ, …',
    chineseCounting: '一, 二, 三, …',
    ideographTraditional: '甲, 乙, 丙, …',
  };
  const SEPARATOR_KEYS = {
    hyphen: 'ins.pnf.hyphen',
    period: 'ins.pnf.period',
    colon: 'ins.pnf.colon',
    emDash: 'ins.pnf.emDash',
    enDash: 'ins.pnf.enDash',
  } as const satisfies Record<ChapterSeparator, string>;
  const HEADING_LEVELS = [1, 2, 3, 4, 5, 6, 7, 8, 9] as const;

  let format = $state<PageNumberFormat>('decimal');
  let includeChapter = $state(false);
  let chapterStyle = $state(1);
  let separator = $state<ChapterSeparator>('hyphen');
  let restart = $state(false);
  let startAt = $state(1);

  $effect(() => {
    if (session.dialog !== DIALOG.pageNumberFormat || !session.model) return;
    const current = getPageNumberFormat(session.model.doc);
    format = current.format ?? 'decimal';
    includeChapter = current.chapterStyle !== undefined;
    chapterStyle = current.chapterStyle ?? 1;
    separator = current.chapterSeparator ?? 'hyphen';
    restart = current.start !== undefined;
    startAt = current.start ?? 1;
  });

  function ok(): boolean {
    session.apply(commands.formatPageNumbersCommand, {
      format,
      ...(restart ? { start: startAt } : {}),
      ...(includeChapter ? { chapterStyle, chapterSeparator: separator } : {}),
    });
    return session.status === '';
  }
</script>

<Dialog id={DIALOG.pageNumberFormat} title={t('ins.pnf.title')} onok={ok}>
  <label class="stack">{t('ins.pnf.format')}
    <select bind:value={format}>
      {#each PAGE_NUMBER_FORMATS as f (f)}<option value={f}>{SAMPLES[f]}</option>{/each}
    </select>
  </label>
  <label><input type="checkbox" bind:checked={includeChapter} /> {t('ins.pnf.includeChapter')}</label>
  <div class="dialog-row">
    <label class="stack">{t('ins.pnf.chapterStyle')}
      <select bind:value={chapterStyle} disabled={!includeChapter}>
        {#each HEADING_LEVELS as level (level)}<option value={level}>{t('ins.heading')} {level}</option>{/each}
      </select>
    </label>
    <label class="stack">{t('ins.pnf.separator')}
      <select bind:value={separator} disabled={!includeChapter}>
        {#each CHAPTER_SEPARATORS as s (s)}<option value={s}>{t(SEPARATOR_KEYS[s])}</option>{/each}
      </select>
    </label>
  </div>
  <fieldset>
    <legend>{t('ins.pnf.numbering')}</legend>
    <label><input type="radio" name="pn-restart" value={false} bind:group={restart} /> {t('ins.pnf.continue')}</label>
    <label><input type="radio" name="pn-restart" value={true} bind:group={restart} /> {t('ins.pnf.startAt')}
      <input type="number" min="0" bind:value={startAt} disabled={!restart} style="width: 72px" /></label>
  </fieldset>
  {#if session.status}<p class="error-note">{session.status}</p>{/if}
</Dialog>
