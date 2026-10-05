<script lang="ts">
  /**
   * Create Source, Source Manager (Master List ⇄ Current List) and New
   * Placeholder. The current list is the document's bibliography part; the
   * master list is kept in this browser.
   */
  import { untrack } from 'svelte';
  import {
    bibliographySources,
    SOURCE_FIELDS,
    SOURCE_TYPES,
    suggestSourceTag,
    type BibliographySource,
    type PersonName,
    type SourceField,
    type SourceType,
  } from '@office-kit/docx';
  import { commands } from '@office-kit/docx-editor';
  import Dialog from '../Dialog.svelte';
  import { getSession } from '../session.svelte';
  import { t } from '../i18n/index.svelte';
  import { masterList, sourceEditor } from './state.svelte';

  const session = getSession();

  // The fields Word's Create Source dialog shows for each type (basic view).
  const DEFAULT_FIELDS: readonly SourceField[] = ['Title', 'Year', 'City', 'Publisher'];
  const FIELDS_BY_TYPE: Partial<Record<SourceType, readonly SourceField[]>> = {
    BookSection: ['Title', 'BookTitle', 'Year', 'Pages', 'City', 'Publisher'],
    JournalArticle: ['Title', 'JournalName', 'Year', 'Pages', 'Volume', 'Issue'],
    ArticleInAPeriodical: ['Title', 'PeriodicalTitle', 'Year', 'Month', 'Day', 'Pages'],
    ConferenceProceedings: ['Title', 'ConferenceName', 'Year', 'Pages', 'City', 'Publisher'],
    Report: ['Title', 'Year', 'Publisher', 'City'],
    InternetSite: ['Title', 'InternetSiteTitle', 'Year', 'Month', 'Day', 'URL'],
    DocumentFromInternetSite: ['Title', 'InternetSiteTitle', 'Year', 'Month', 'Day', 'URL'],
  };

  let type = $state<SourceType>('Book');
  let authors = $state('');
  let corporate = $state(false);
  let fields = $state<Partial<Record<SourceField, string>>>({});
  let tag = $state('');
  const shownFields = $derived(FIELDS_BY_TYPE[type] ?? DEFAULT_FIELDS);

  function formatPeople(people: readonly PersonName[] | undefined): string {
    return (people ?? []).map((p) => [p.last, [p.first, p.middle].filter(Boolean).join(' ')].filter(Boolean).join(', ')).join('; ');
  }

  // "Last, First Middle; Last, First" — Word's Edit Name list as one line.
  function parsePeople(text: string): PersonName[] {
    return text
      .split(';')
      .map((part) => part.trim())
      .filter(Boolean)
      .map((part) => {
        const [last = '', given = ''] = part.split(',').map((s) => s.trim());
        const [first, ...middle] = given.split(/\s+/).filter(Boolean);
        return { last, ...(first ? { first } : {}), ...(middle.length ? { middle: middle.join(' ') } : {}) };
      });
  }

  $effect(() => {
    if (session.dialog !== 'references.source') return;
    untrack(() => {
      const s = sourceEditor.editing;
      type = s?.type ?? 'Book';
      corporate = s?.corporateAuthor !== undefined;
      authors = s?.corporateAuthor ?? formatPeople(s?.authors);
      fields = { ...s?.fields };
      tag = s?.tag ?? '';
    });
  });

  function currentSources(): BibliographySource[] {
    return session.model ? bibliographySources(session.model.doc) : [];
  }

  function saveSource(): boolean {
    const filled: Partial<Record<SourceField, string>> = {};
    for (const field of SOURCE_FIELDS) {
      const value = fields[field]?.trim();
      if (value) filled[field] = value;
    }
    const body: Omit<BibliographySource, 'tag'> = {
      type,
      ...(corporate ? { corporateAuthor: authors.trim() } : { authors: parsePeople(authors) }),
      fields: filled,
    };
    const existing = currentSources().filter((s) => s.tag !== sourceEditor.editing?.tag);
    const source: BibliographySource = { tag: tag.trim() || suggestSourceTag(existing, body), ...body };
    session.apply(commands.setSourcesCommand, { sources: [...existing, source] });
    if (session.status) return false;
    masterList.save(source);
    if (sourceEditor.citeAfterSave) session.apply(commands.insertCitationCommand, { tag: source.tag });
    sourceEditor.editing = null;
    sourceEditor.citeAfterSave = false;
    return true;
  }

  // --- Source Manager -------------------------------------------------------------------
  let masterPick = $state('');
  let currentPick = $state('');
  const current = $derived(session.tick >= 0 ? currentSources() : []);

  function copyToCurrent(): void {
    const source = masterList.sources.find((s) => s.tag === masterPick);
    if (!source) return;
    session.apply(commands.setSourcesCommand, { sources: [...current.filter((s) => s.tag !== source.tag), source] });
  }

  function copyToMaster(): void {
    const source = current.find((s) => s.tag === currentPick);
    if (source) masterList.save(source);
  }

  function remove(): void {
    if (currentPick) {
      session.apply(commands.setSourcesCommand, { sources: current.filter((s) => s.tag !== currentPick) });
      currentPick = '';
    } else if (masterPick) {
      masterList.remove(masterPick);
      masterPick = '';
    }
  }

  function edit(): void {
    const source = current.find((s) => s.tag === currentPick) ?? masterList.sources.find((s) => s.tag === masterPick);
    if (!source) return;
    sourceEditor.editing = source;
    sourceEditor.citeAfterSave = false;
    session.openDialog('references.source');
  }

  function create(): void {
    sourceEditor.editing = null;
    sourceEditor.citeAfterSave = false;
    session.openDialog('references.source');
  }

  const label = (s: BibliographySource): string =>
    [s.corporateAuthor ?? s.authors?.[0]?.last, s.fields?.Year, s.fields?.Title].filter(Boolean).join(' · ') || s.tag;

  // --- Placeholder ----------------------------------------------------------------------
  let placeholder = $state('');

  function addPlaceholder(): boolean {
    const name = placeholder.trim();
    if (!/^[\p{L}\p{N}_]+$/u.test(name)) {
      session.status = t('ref.placeholderInvalid');
      return false;
    }
    if (!current.some((s) => s.tag === name)) {
      session.apply(commands.setSourcesCommand, { sources: [...current, { tag: name, type: 'Misc' }] });
    }
    session.apply(commands.insertCitationCommand, { tag: name });
    placeholder = '';
    return true;
  }
</script>

<Dialog id="references.source" title={t(sourceEditor.editing ? 'ref.editSource' : 'ref.createSource')} onok={saveSource}>
  <label class="field">{t('ref.sourceType')}
    <select bind:value={type}>
      {#each SOURCE_TYPES as st (st)}<option value={st}>{t(`ref.type.${st}`)}</option>{/each}
    </select>
  </label>
  <label class="field">{t(corporate ? 'ref.corporateAuthor' : 'ref.author')}
    <input bind:value={authors} placeholder={corporate ? '' : t('ref.authorHint')} size="36" />
  </label>
  <label><input type="checkbox" bind:checked={corporate} /> {t('ref.corporateAuthor')}</label>
  {#each shownFields as field (field)}
    <label class="field">{t(`ref.field.${field}`)} <input bind:value={fields[field]} size="36" /></label>
  {/each}
  <label class="field">{t('ref.tagName')} <input bind:value={tag} /></label>
</Dialog>

<Dialog id="references.sources" title={t('ref.manageSources')}>
  <div class="row" style="align-items: stretch; gap: 8px">
    <label class="col">{t('ref.masterList')}
      <select size="8" bind:value={masterPick} onfocus={() => (currentPick = '')} style="min-width: 220px">
        {#each masterList.sources as s (s.tag)}<option value={s.tag}>{label(s)}</option>{/each}
      </select>
    </label>
    <div class="col" style="justify-content: center">
      <button type="button" class="push" onclick={copyToCurrent} disabled={!masterPick}>{t('ref.copyRight')}</button>
      <button type="button" class="push" onclick={copyToMaster} disabled={!currentPick}>{t('ref.copyLeft')}</button>
      <button type="button" class="push" onclick={remove} disabled={!masterPick && !currentPick}>{t('ref.delete')}</button>
      <button type="button" class="push" onclick={edit} disabled={!masterPick && !currentPick}>{t('ref.edit')}</button>
      <button type="button" class="push" onclick={create}>{t('ref.new')}</button>
    </div>
    <label class="col">{t('ref.currentList')}
      <select size="8" bind:value={currentPick} onfocus={() => (masterPick = '')} style="min-width: 220px">
        {#each current as s (s.tag)}<option value={s.tag}>{label(s)}</option>{/each}
      </select>
    </label>
  </div>
</Dialog>

<Dialog id="references.placeholder" title={t('ref.placeholderTitle')} onok={addPlaceholder}>
  <label class="field">{t('ref.placeholderName')} <input bind:value={placeholder} /></label>
</Dialog>
