<script lang="ts">
  /**
   * Design ▸ Paragraph Spacing ▸ Custom Paragraph Spacing…: the paragraph
   * spacing part of Word's Manage Styles ▸ Set Defaults, written to the
   * document defaults.
   */
  import { commands } from '@office-kit/docx-editor';
  import { getDefaultParagraphSpacing } from '@office-kit/docx';
  import Dialog from '../Dialog.svelte';
  import LengthField from './LengthField.svelte';
  import { getSession } from '../session.svelte';
  import { t } from '../i18n/index.svelte';

  const ID = 'design.spacing';
  const LINE_UNIT = 240;
  const LINE_CHOICES = [1, 1.07, 1.08, 1.15, 1.5, 2, 2.5, 3];
  const session = getSession();

  let open = $state(false);
  let before = $state(0);
  let after = $state(160);
  let line = $state(1.08);

  $effect(() => {
    const showing = session.dialog === ID;
    if (showing && !open && session.model) {
      const current = getDefaultParagraphSpacing(session.model.doc);
      before = current?.before ?? 0;
      after = current?.after ?? 0;
      line = current ? Math.round((current.line / LINE_UNIT) * 100) / 100 : 1;
    }
    open = showing;
  });

  function ok(): boolean | void {
    if (!(line > 0 && line <= 132)) return false;
    session.apply(commands.paragraphSpacingCommand, { spacing: { before, after, line: Math.round(line * LINE_UNIT) } });
  }
</script>

<Dialog id={ID} title={t('dsn.spacing.customTitle')} onok={ok}>
  <div class="form-grid">
    <label for="ps-before">{t('dsn.spacing.before')}</label>
    <LengthField id="ps-before" value={before} unit="pt" label={t('dsn.spacing.before')} max={31680} onchange={(v) => (before = v)} />
    <label for="ps-after">{t('dsn.spacing.after')}</label>
    <LengthField id="ps-after" value={after} unit="pt" label={t('dsn.spacing.after')} max={31680} onchange={(v) => (after = v)} />
    <label for="ps-line">{t('dsn.spacing.line')}</label>
    <input id="ps-line" type="number" min="0.06" max="132" step="0.01" list="ps-lines" bind:value={line} />
  </div>
  <datalist id="ps-lines">{#each LINE_CHOICES as v (v)}<option value={v}></option>{/each}</datalist>
</Dialog>
