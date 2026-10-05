<script lang="ts">
  /**
   * File ▸ Properties (on the View tab here): Summary (core and extended
   * properties), Statistics, and Custom (docProps/custom.xml). Custom
   * properties are added and deleted at once, as Word's Add / Delete buttons do.
   */
  import { commands } from '@office-kit/docx-editor';
  import { appProperties, coreProperties, customProperties, wordCount, type CustomPropertyValue } from '@office-kit/docx';
  import Dialog from '../Dialog.svelte';
  import { getSession } from '../session.svelte';
  import { t, type MessageKey } from '../i18n/index.svelte';

  const session = getSession();
  type Tab = 'summary' | 'statistics' | 'custom';
  const TABS: readonly [Tab, MessageKey][] = [
    ['summary', 'props.summary'],
    ['statistics', 'props.statistics'],
    ['custom', 'props.custom'],
  ];
  type CustomType = 'text' | 'date' | 'number' | 'yesno';
  const TYPES: readonly [CustomType, MessageKey][] = [
    ['text', 'props.typeText'],
    ['date', 'props.typeDate'],
    ['number', 'props.typeNumber'],
    ['yesno', 'props.typeYesNo'],
  ];

  let tab = $state<Tab>('summary');
  const summary = $state({
    title: '',
    subject: '',
    creator: '',
    manager: '',
    company: '',
    category: '',
    keywords: '',
    description: '',
    hyperlinkBase: '',
  });
  let customName = $state('');
  let customType = $state<CustomType>('text');
  let customValue = $state('');
  let customError = $state('');

  $effect(() => {
    if (session.dialog !== 'properties' || !session.model) return;
    const core = coreProperties(session.model.doc);
    const app = appProperties(session.model.doc);
    summary.title = core.title ?? '';
    summary.subject = core.subject ?? '';
    summary.creator = core.creator ?? '';
    summary.category = core.category ?? '';
    summary.keywords = core.keywords ?? '';
    summary.description = core.description ?? '';
    summary.manager = app.manager ?? '';
    summary.company = app.company ?? '';
    summary.hyperlinkBase = app.hyperlinkBase ?? '';
    tab = 'summary';
  });

  const stats = $derived.by(() => {
    if (session.dialog !== 'properties' || session.tick < 0 || !session.model) return undefined;
    return { core: coreProperties(session.model.doc), counts: wordCount(session.model.doc) };
  });
  const custom = $derived(
    session.dialog === 'properties' && session.tick >= 0 && session.model ? customProperties(session.model.doc) : [],
  );

  const SUMMARY_FIELDS: readonly [keyof typeof summary, MessageKey][] = [
    ['title', 'props.title'],
    ['subject', 'props.subject'],
    ['creator', 'props.author'],
    ['manager', 'props.manager'],
    ['company', 'props.company'],
    ['category', 'props.category'],
    ['keywords', 'props.keywords'],
    ['description', 'props.comments'],
    ['hyperlinkBase', 'props.hyperlinkBase'],
  ];

  function apply(): void {
    const { manager, company, hyperlinkBase, ...core } = summary;
    session.apply(commands.setCorePropertiesCommand, core);
    session.apply(commands.setAppPropertiesCommand, { manager, company, hyperlinkBase });
  }

  type Parsed = { readonly value: CustomPropertyValue } | { readonly error: string };

  function parseCustom(): Parsed {
    switch (customType) {
      case 'text':
        return { value: customValue };
      case 'number': {
        const n = Number(customValue);
        return customValue.trim() !== '' && Number.isFinite(n) ? { value: n } : { error: t('props.invalidValue') };
      }
      case 'date': {
        const d = new Date(customValue);
        return Number.isNaN(d.getTime()) ? { error: t('props.invalidValue') } : { value: d };
      }
      case 'yesno':
        return { value: customValue === 'yes' };
    }
  }

  function addCustom(): void {
    const parsed = parseCustom();
    if ('error' in parsed) {
      customError = parsed.error;
      return;
    }
    customError = '';
    session.apply(commands.setCustomPropertyCommand, { name: customName.trim(), value: parsed.value });
    customName = '';
    customValue = '';
  }

  function show(value: CustomPropertyValue): string {
    if (value instanceof Date) return value.toLocaleString();
    if (typeof value === 'boolean') return value ? t('props.yes') : t('props.no');
    return String(value);
  }

  function date(iso: string | undefined): string {
    return iso ? new Date(iso).toLocaleString() : '—';
  }
</script>

<Dialog id="properties" title={`${session.fileName} — ${t('view.properties')}`} onok={apply}>
  <div class="tabs" role="tablist">
    {#each TABS as [id, key] (id)}
      <button type="button" role="tab" class:active={tab === id} aria-selected={tab === id} onclick={() => (tab = id)}>{t(key)}</button>
    {/each}
  </div>
  {#if tab === 'summary'}
    <div class="grid2">
      {#each SUMMARY_FIELDS as [field, key] (field)}
        <label for="prop-{field}">{t(key)}</label>
        {#if field === 'description'}
          <textarea id="prop-{field}" rows="3" bind:value={summary[field]}></textarea>
        {:else}
          <input id="prop-{field}" bind:value={summary[field]} />
        {/if}
      {/each}
    </div>
  {:else if tab === 'statistics' && stats}
    <table class="stats">
      <tbody>
        <tr><th>{t('props.created')}</th><td>{date(stats.core.created)}</td></tr>
        <tr><th>{t('props.modified')}</th><td>{date(stats.core.modified)}</td></tr>
        <tr><th>{t('props.lastSavedBy')}</th><td>{stats.core.lastModifiedBy ?? '—'}</td></tr>
        <tr><th>{t('props.revision')}</th><td>{stats.core.revision ?? '—'}</td></tr>
        <tr><th>{t('review.wcPages')}</th><td>{session.pageCount}</td></tr>
        <tr><th>{t('review.wcParagraphs')}</th><td>{stats.counts.paragraphs}</td></tr>
        <tr><th>{t('review.wcWords')}</th><td>{stats.counts.words}</td></tr>
        <tr><th>{t('review.wcChars')}</th><td>{stats.counts.characters}</td></tr>
        <tr><th>{t('review.wcCharsSpaces')}</th><td>{stats.counts.charactersWithSpaces}</td></tr>
      </tbody>
    </table>
  {:else if tab === 'custom'}
    <div class="grid2">
      <label for="custom-name">{t('props.name')}</label>
      <input id="custom-name" bind:value={customName} />
      <label for="custom-type">{t('props.type')}</label>
      <select id="custom-type" bind:value={customType} onchange={() => (customValue = customType === 'yesno' ? 'yes' : '')}>
        {#each TYPES as [v, key] (v)}<option value={v}>{t(key)}</option>{/each}
      </select>
      <label for="custom-value">{t('props.value')}</label>
      {#if customType === 'yesno'}
        <select id="custom-value" bind:value={customValue}>
          <option value="yes">{t('props.yes')}</option>
          <option value="no">{t('props.no')}</option>
        </select>
      {:else}
        <input id="custom-value" type={customType === 'date' ? 'date' : customType === 'number' ? 'number' : 'text'} bind:value={customValue} />
      {/if}
    </div>
    <div><button type="button" class="push" onclick={addCustom} disabled={!customName.trim()}>{t('props.add')}</button></div>
    {#if customError}<p class="pane-note" role="alert">{customError}</p>{/if}
    <div>{t('props.list')}</div>
    <table class="stats custom">
      <tbody>
        {#each custom as prop (prop.name)}
          <tr>
            <th>{prop.name}</th>
            <td>{show(prop.value)}</td>
            <td><button type="button" class="push" onclick={() => session.apply(commands.removeCustomPropertyCommand, { name: prop.name })}>{t('props.delete')}</button></td>
          </tr>
        {/each}
      </tbody>
    </table>
  {/if}
</Dialog>

<style>
  .tabs { display: flex; justify-content: center; gap: 2px; }
  .tabs button { padding: 3px 12px; border: 1px solid var(--control-line); background: #fff; cursor: pointer; }
  .tabs button.active { background: var(--accent); color: #fff; border-color: var(--accent); }
  .grid2 { display: grid; grid-template-columns: auto minmax(260px, 1fr); gap: 6px 10px; align-items: center; }
  .grid2 input, .grid2 select, .grid2 textarea { font: inherit; }
  .stats { border-collapse: collapse; }
  .stats th { text-align: left; font-weight: normal; padding: 2px 24px 2px 0; }
  .stats td { padding: 2px 0; }
  .custom td { padding-right: 12px; }
</style>
