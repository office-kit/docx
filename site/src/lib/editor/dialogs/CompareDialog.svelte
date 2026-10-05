<script lang="ts">
  /**
   * Review ▸ Compare ▸ Compare Documents. The differences between the
   * original and the revised document open as a new document of tracked
   * changes, as Word's "Show changes in: New document" does; the open
   * document is left as it was.
   */
  import { compareDocuments, openDocx, type Docx } from '@office-kit/docx';
  import { editorFor } from '@office-kit/docx-editor';
  import Dialog from '../Dialog.svelte';
  import { getSession } from '../session.svelte';
  import { t } from '../i18n/index.svelte';

  const session = getSession();
  const CURRENT = 'current';

  let originalSource = $state<'current' | 'file'>(CURRENT);
  let originalFile = $state<File | null>(null);
  let revisedFile = $state<File | null>(null);
  let label = $state('');
  let error = $state('');

  $effect(() => {
    if (session.dialog !== 'compare') return;
    label = session.prefs.userName;
    error = '';
  });

  const ready = $derived(!!revisedFile && (originalSource === CURRENT ? !!session.model : !!originalFile));

  async function read(file: File): Promise<Docx> {
    return openDocx(new Uint8Array(await file.arrayBuffer()));
  }

  function picked(e: Event): File | null {
    return (e.currentTarget as HTMLInputElement).files?.[0] ?? null;
  }

  async function run(): Promise<void> {
    if (!revisedFile) return;
    try {
      const original = originalSource === CURRENT || !originalFile ? session.model?.doc : await read(originalFile);
      if (!original) return;
      const result = compareDocuments(original, await read(revisedFile), { author: label || session.prefs.userName });
      session.load(editorFor(result), `${t('review.cmpResult')}.docx`);
      session.markup = 'all';
      session.dialog = null;
    } catch (err) {
      // A file that is not a .docx: report it in the dialog and keep it open.
      error = (err as Error).message;
    }
  }
</script>

<Dialog id="compare" title={t('review.compareDocuments').replace('…', '')} onok={() => { void run(); return false; }} okDisabled={!ready}>
  <fieldset>
    <legend>{t('review.cmpOriginal')}</legend>
    <label class="check-item"><input type="radio" bind:group={originalSource} value={CURRENT} />{t('review.cmpCurrent')} — {session.fileName}</label>
    <label class="check-item">
      <input type="radio" bind:group={originalSource} value="file" />
      <input type="file" accept=".docx" onchange={(e) => { originalFile = picked(e); originalSource = 'file'; }} aria-label={t('review.cmpOriginal')} />
    </label>
  </fieldset>
  <fieldset>
    <legend>{t('review.cmpRevised')}</legend>
    <input type="file" accept=".docx" onchange={(e) => (revisedFile = picked(e))} aria-label={t('review.cmpRevised')} />
  </fieldset>
  <label class="field">{t('review.cmpLabel')} <input bind:value={label} /></label>
  {#if error}<p class="pane-note" role="alert">{error}</p>{/if}
</Dialog>
