<script lang="ts">
  /**
   * Insert ▸ Date & Time: Word's formats for the chosen language. "Update
   * automatically" inserts a DATE field with the `\@` picture (in that
   * language, so updates keep its month names); otherwise plain text.
   */
  import { formatDatePicture } from '@office-kit/docx';
  import { commands } from '@office-kit/docx-editor';
  import Dialog from '../Dialog.svelte';
  import { getSession } from '../session.svelte';
  import { LOCALES, locale, t, type LocaleId } from '../i18n/index.svelte';
  import { DATE_FORMATS, DATE_LANGUAGE } from './date-formats';
  import { DIALOG } from './state.svelte';

  const session = getSession();
  const LANGUAGES = Object.keys(LOCALES).filter((id): id is LocaleId => id in LOCALES);

  let language = $state<LocaleId>('en');
  let selected = $state(0);
  let auto = $state(false);
  let now = $state(new Date());

  $effect(() => {
    if (session.dialog !== DIALOG.dateTime) return;
    language = locale();
    selected = 0;
    now = new Date();
  });

  const formats = $derived(DATE_FORMATS[language]);
  const lang = $derived(DATE_LANGUAGE[language]);

  function ok(): boolean {
    const picture = formats[selected];
    if (!picture) return false;
    if (auto) {
      session.apply(commands.insertFieldAtCaretCommand, { instruction: `DATE \\@ "${picture}"`, lang, context: { now: new Date() } });
    } else {
      session.apply(commands.insertTextCommand, { text: formatDatePicture(new Date(), picture, lang) });
    }
    return session.status === '';
  }
</script>

<Dialog id={DIALOG.dateTime} title={t('ins.dt.title')} onok={ok}>
  <div class="dialog-row">
    <div class="dialog-col">
      <span>{t('ins.dt.formats')}</span>
      <div class="listbox" role="listbox" aria-label={t('ins.dt.formats')} style="height: 220px">
        {#each formats as picture, i (picture)}
          <button type="button" role="option" aria-selected={i === selected} class:selected={i === selected} onclick={() => (selected = i)} ondblclick={() => { selected = i; if (ok()) session.dialog = null; }}>
            {formatDatePicture(now, picture, lang)}
          </button>
        {/each}
      </div>
    </div>
    <div class="dialog-col">
      <label class="stack">{t('ins.dt.language')}
        <select value={language} onchange={(e) => { const next = LANGUAGES.find((id) => id === e.currentTarget.value); if (next) { language = next; selected = 0; } }}>
          {#each LANGUAGES as id (id)}<option value={id}>{LOCALES[id]}</option>{/each}
        </select>
      </label>
      <label><input type="checkbox" bind:checked={auto} /> {t('ins.dt.update')}</label>
    </div>
  </div>
  {#if session.status}<p class="error-note">{session.status}</p>{/if}
</Dialog>
