<script lang="ts">
  /** Layout ▸ Line Numbers ▸ Line Numbering Options…: Word's Line Numbers dialog. */
  import { commands } from '@office-kit/docx-editor';
  import type { LineNumberRestart } from '@office-kit/docx';
  import Dialog from '../Dialog.svelte';
  import LengthField from './LengthField.svelte';
  import { getSession } from '../session.svelte';
  import { locale, t, type MessageKey } from '../i18n/index.svelte';
  import { lengthUnitFor } from './units';

  const ID = 'layout.lineNumbers';
  const session = getSession();
  const unit = $derived(lengthUnitFor(locale()));
  const RESTARTS: ReadonlyArray<{ value: LineNumberRestart; key: MessageKey }> = [
    { value: 'newPage', key: 'lay.ln.restartPage' },
    { value: 'newSection', key: 'lay.ln.restartSection' },
    { value: 'continuous', key: 'lay.ln.continuous' },
  ];

  let open = false;
  let enabled = $state(false);
  let start = $state(1);
  let countBy = $state(1);
  let distance = $state<number | undefined>(undefined);
  let restart = $state<LineNumberRestart>('newPage');

  $effect(() => {
    const showing = session.dialog === ID;
    if (showing && !open && session.model) {
      const ln = commands.currentSectionProperties(session.model).lineNumbering;
      enabled = !!ln;
      start = ln?.start ?? 1;
      countBy = ln?.countBy ?? 1;
      distance = ln?.distanceTwips;
      restart = ln?.restart ?? 'newPage';
    }
    open = showing;
  });

  function ok(): boolean | void {
    if (enabled && (!Number.isInteger(start) || start < 1 || !Number.isInteger(countBy) || countBy < 1 || countBy > 100)) return false;
    session.apply(commands.lineNumbersCommand, {
      lineNumbering: enabled ? { start, countBy, restart, ...(distance === undefined ? {} : { distanceTwips: distance }) } : null,
    });
  }
</script>

<Dialog id={ID} title={t('lay.ln.title')} onok={ok}>
  <label class="check"><input type="checkbox" bind:checked={enabled} />{t('lay.ln.add')}</label>
  <div class="form-grid" class:dim={!enabled}>
    <label for="ln-start">{t('lay.ln.startAt')}</label>
    <input id="ln-start" type="number" min="1" max="32767" bind:value={start} disabled={!enabled} />
    <label for="ln-from">{t('lay.ln.fromText')}</label>
    <LengthField id="ln-from" value={distance} {unit} placeholder={t('dsn.auto')} label={t('lay.ln.fromText')} disabled={!enabled} onchange={(v) => (distance = v)} />
    <label for="ln-by">{t('lay.ln.countBy')}</label>
    <input id="ln-by" type="number" min="1" max="100" bind:value={countBy} disabled={!enabled} />
  </div>
  <fieldset class:dim={!enabled}>
    <legend>{t('lay.ln.numbering')}</legend>
    {#each RESTARTS as r (r.value)}
      <label class="radio"><input type="radio" bind:group={restart} value={r.value} disabled={!enabled} />{t(r.key)}</label>
    {/each}
  </fieldset>
</Dialog>
