<script lang="ts">
  /** Layout ▸ Text Direction ▸ Text Direction Options… ("Text Direction - Main Document"). */
  import { commands, type SectionTarget } from '@office-kit/docx-editor';
  import type { SectionTextDirection } from '@office-kit/docx';
  import Dialog from '../Dialog.svelte';
  import { getSession } from '../session.svelte';
  import { t, type MessageKey } from '../i18n/index.svelte';

  const ID = 'layout.textDirection';
  const session = getSession();
  const CHOICES: ReadonlyArray<{ value: SectionTextDirection; key: MessageKey }> = [
    { value: 'lrTb', key: 'lay.td.horizontal' },
    { value: 'tbRl', key: 'lay.td.rotate90' },
    { value: 'btLr', key: 'lay.td.rotate270' },
  ];

  let open = false;
  let direction = $state<SectionTextDirection>('lrTb');
  let applyTo = $state<SectionTarget>('section');

  $effect(() => {
    const showing = session.dialog === ID;
    if (showing && !open && session.model) {
      direction = commands.currentSectionProperties(session.model).textDirection;
      applyTo = 'section';
    }
    open = showing;
  });
</script>

<Dialog id={ID} title={t('lay.td.title')} onok={() => session.apply(commands.textDirectionCommand, { direction, target: applyTo })}>
  <fieldset>
    <legend>{t('lay.orientation')}</legend>
    <div class="row">
      {#each CHOICES as c (c.value)}
        <label class="td-choice" class:selected={direction === c.value}>
          <input type="radio" bind:group={direction} value={c.value} />
          <span class="td-sample td-{c.value}">{t('dsn.sample.text')}</span>
          {t(c.key)}
        </label>
      {/each}
    </div>
  </fieldset>
  <div class="field apply-to">
    <label for="td-apply">{t('dsn.applyTo')}</label>
    <select id="td-apply" bind:value={applyTo}>
      <option value="section">{t('dsn.applyTo.section')}</option>
      <option value="document">{t('dsn.applyTo.document')}</option>
      <option value="forward">{t('lay.forward')}</option>
    </select>
  </div>
</Dialog>
