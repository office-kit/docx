<script lang="ts">
  /**
   * Footnote and Endnote: location, number format, custom mark, start at and
   * numbering; Insert adds a note with these settings, Apply only stores them.
   * Convert… moves notes between the footnotes and endnotes.
   */
  import { noteProperties, type NoteConversion, type NoteKind, type NotePosition, type NoteRestart, type NumberingFormat } from '@office-kit/docx';
  import { untrack } from 'svelte';
  import { commands } from '@office-kit/docx-editor';
  import Dialog from '../Dialog.svelte';
  import { getSession } from '../session.svelte';
  import { t } from '../i18n/index.svelte';
  import { NUMBER_FORMATS } from './options';

  const session = getSession();
  const ID = 'references.notes';

  let kind = $state<NoteKind>('footnote');
  let position = $state<NotePosition>('pageBottom');
  let numberFormat = $state<NumberingFormat>('decimal');
  let customMark = $state('');
  let startAt = $state(1);
  let restart = $state<NoteRestart>('continuous');
  let conversion = $state<NoteConversion>('footnotesToEndnotes');

  const POSITIONS: Readonly<Record<NoteKind, readonly NotePosition[]>> = {
    footnote: ['pageBottom', 'beneathText'],
    endnote: ['sectEnd', 'docEnd'],
  };

  // Show the document's current settings for the chosen kind, as Word does.
  function load(next: NoteKind): void {
    kind = next;
    const model = session.model;
    if (!model) return;
    const current = noteProperties(model.doc, next);
    position = current.position;
    numberFormat = current.numberFormat;
    startAt = current.startAt;
    restart = current.restart;
  }

  $effect(() => {
    if (session.dialog === ID) untrack(() => load(kind));
  });

  function apply(): void {
    session.apply(commands.noteOptionsCommand, { kind, properties: { position, numberFormat, startAt, restart }, scope: 'document' });
  }

  function insert(): void {
    apply();
    const command = kind === 'footnote' ? commands.addFootnoteCommand : commands.addEndnoteCommand;
    const mark = customMark.trim();
    if (session.apply(command, mark ? { customMark: mark } : {}) !== undefined) session.pane.right = 'notes';
  }

  function convert(): void {
    const n = session.apply(commands.convertNotesCommand, { conversion });
    if (n !== undefined) session.status = t('ref.convertedNotes').replace('{n}', String(n));
  }
</script>

<Dialog id={ID} title={t('ref.noteDialog')} okLabel={t('ref.insert')} onok={insert}>
  <fieldset>
    <legend>{t('ref.location')}</legend>
    <label><input type="radio" checked={kind === 'footnote'} onchange={() => load('footnote')} /> {t('ref.footnotes')}</label>
    <label><input type="radio" checked={kind === 'endnote'} onchange={() => load('endnote')} /> {t('ref.endnotes')}</label>
    <label class="field">{t('ref.position')}
      <select bind:value={position}>
        {#each POSITIONS[kind] as p (p)}<option value={p}>{t(`ref.pos.${p}`)}</option>{/each}
      </select>
    </label>
  </fieldset>
  <fieldset>
    <legend>{t('ref.format')}</legend>
    <label class="field">{t('ref.numberFormat')}
      <select bind:value={numberFormat}>
        {#each NUMBER_FORMATS as f (f.value)}<option value={f.value}>{f.sample}</option>{/each}
      </select>
    </label>
    <label class="field">{t('ref.customMark')} <input bind:value={customMark} maxlength="10" /></label>
    <label class="field">{t('ref.startAt')} <input type="number" min="1" bind:value={startAt} /></label>
    <label class="field">{t('ref.numbering')}
      <select bind:value={restart}>
        <option value="continuous">{t('ref.continuous')}</option>
        <option value="eachSect">{t('ref.restartSection')}</option>
        {#if kind === 'footnote'}<option value="eachPage">{t('ref.restartPage')}</option>{/if}
      </select>
    </label>
    <button type="button" class="push" onclick={apply}>{t('ref.apply')}</button>
  </fieldset>
  <fieldset>
    <legend>{t('ref.convert')}</legend>
    <select bind:value={conversion} aria-label={t('ref.convert')}>
      <option value="footnotesToEndnotes">{t('ref.convertToEndnotes')}</option>
      <option value="endnotesToFootnotes">{t('ref.convertToFootnotes')}</option>
      <option value="swap">{t('ref.swapNotes')}</option>
    </select>
    <button type="button" class="push" onclick={convert}>{t('ref.convertButton')}</button>
  </fieldset>
</Dialog>
