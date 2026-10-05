<script lang="ts">
  /** Insert Caption, with New Label and Numbering (format, chapter number). */
  import { captionLabels, type NumberingFormat } from '@office-kit/docx';
  import { commands } from '@office-kit/docx-editor';
  import Dialog from '../Dialog.svelte';
  import { getSession } from '../session.svelte';
  import { t } from '../i18n/index.svelte';
  import { NUMBER_FORMATS } from './options';

  const session = getSession();

  // Word's chapter-number separators: hyphen, period, colon, em dash, en dash.
  const SEPARATORS = ['-', '.', ':', '—', '–'] as const;
  const HEADING_LEVELS = [1, 2, 3, 4, 5, 6, 7, 8, 9] as const;

  let label = $state('Figure');
  let text = $state('');
  let position = $state<'above' | 'below'>('below');
  let excludeLabel = $state(false);
  let numberFormat = $state<NumberingFormat>('decimal');
  let chapter = $state(false);
  let headingLevel = $state(1);
  let separator = $state<string>('-');
  let newLabel = $state('');

  const labels = $derived(session.tick >= 0 && session.model ? captionLabels(session.model.doc) : []);

  function addLabel(): void {
    const name = newLabel.trim();
    if (!name) return;
    session.apply(commands.addCaptionLabelCommand, { name });
    if (!session.status) {
      label = name;
      newLabel = '';
    }
  }

  function insert(): void {
    session.apply(commands.insertCaptionCommand, {
      position,
      options: {
        label,
        excludeLabel,
        numberFormat,
        ...(text ? { text } : {}),
        ...(chapter ? { chapter: { headingLevel, separator } } : {}),
      },
    });
  }
</script>

<Dialog id="references.caption" title={t('ref.insertCaption')} onok={insert}>
  <label class="field">{t('ref.caption')}
    <span>{excludeLabel ? '' : `${label} `}1</span><input bind:value={text} size="28" />
  </label>
  <label class="field">{t('ref.label')}
    <select bind:value={label}>
      {#each labels as l (l)}<option value={l}>{l}</option>{/each}
    </select>
  </label>
  <label class="field">{t('ref.position')}
    <select bind:value={position}>
      <option value="above">{t('ref.aboveSelection')}</option>
      <option value="below">{t('ref.belowSelection')}</option>
    </select>
  </label>
  <label><input type="checkbox" bind:checked={excludeLabel} /> {t('ref.excludeLabel')}</label>
  <div class="row">
    <input bind:value={newLabel} placeholder={t('ref.newLabel')} aria-label={t('ref.newLabel')} />
    <button type="button" class="push" onclick={addLabel} disabled={!newLabel.trim()}>{t('ref.newLabel')}</button>
  </div>
  <fieldset>
    <legend>{t('ref.numberingTitle')}</legend>
    <label class="field">{t('ref.format')}
      <select bind:value={numberFormat}>
        {#each NUMBER_FORMATS.filter((f) => f.value !== 'chicago') as f (f.value)}<option value={f.value}>{f.sample}</option>{/each}
      </select>
    </label>
    <label><input type="checkbox" bind:checked={chapter} /> {t('ref.includeChapter')}</label>
    <label class="field">{t('ref.chapterStarts')}
      <select bind:value={headingLevel} disabled={!chapter}>
        {#each HEADING_LEVELS as level (level)}<option value={level}>{t('ref.heading')} {level}</option>{/each}
      </select>
    </label>
    <label class="field">{t('ref.separator')}
      <select bind:value={separator} disabled={!chapter}>
        {#each SEPARATORS as s (s)}<option value={s}>{s}</option>{/each}
      </select>
    </label>
  </fieldset>
</Dialog>
