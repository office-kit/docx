<script lang="ts">
  /** Review ▸ Language ▸ Set Proofing Language: `w:lang` and `w:noProof` on the selection. */
  import { commands, runAtPath } from '@office-kit/docx-editor';
  import { getRunLanguage, getRunProp } from '@office-kit/docx';
  import Dialog from '../Dialog.svelte';
  import { getSession } from '../session.svelte';
  import { locale, t } from '../i18n/index.svelte';
  import { DEFAULT_LANGUAGE, EAST_ASIAN_LANGUAGES, LATIN_LANGUAGES, languageName } from '../languages';

  const session = getSession();
  let latin = $state(DEFAULT_LANGUAGE);
  let eastAsia = $state('');
  let noProof = $state(false);

  // Start from the language at the caret each time the dialog opens.
  $effect(() => {
    if (session.dialog !== 'language' || !session.model?.selection) return;
    const run = runAtPath(session.model.doc, session.model.selection.focus);
    const language = run ? getRunLanguage(run) : {};
    latin = language.latin ?? DEFAULT_LANGUAGE;
    eastAsia = language.eastAsia ?? '';
    const proof = run ? getRunProp(run, 'noProof') : undefined;
    noProof = !!proof?.present && proof.val !== '0' && proof.val !== 'false';
  });

  // A tag from the document that is not in the list is still offered.
  const latinChoices = $derived(LATIN_LANGUAGES.includes(latin) ? LATIN_LANGUAGES : [latin, ...LATIN_LANGUAGES]);

  function apply(): void {
    session.apply(commands.setProofingLanguageCommand, {
      language: { latin, ...(eastAsia ? { eastAsia } : {}) },
      noProof,
    });
  }
</script>

<Dialog id="language" title={t('review.language')} onok={apply}>
  <label class="lang-field">
    {t('review.langMark')}
    <select bind:value={latin} size="10">
      {#each latinChoices as tag (tag)}<option value={tag}>{languageName(tag, locale())}</option>{/each}
    </select>
  </label>
  <label class="field">
    {t('review.langEastAsia')}
    <select bind:value={eastAsia}>
      <option value="">—</option>
      {#each EAST_ASIAN_LANGUAGES as tag (tag)}<option value={tag}>{languageName(tag, locale())}</option>{/each}
    </select>
  </label>
  <label class="check-item"><input type="checkbox" bind:checked={noProof} />{t('review.langNoProof')}</label>
</Dialog>

<style>
  .lang-field { display: flex; flex-direction: column; gap: 4px; }
  .lang-field select { min-width: 300px; font: inherit; }
</style>
