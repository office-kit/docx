<script lang="ts">
  /**
   * Type a New List, Edit Recipient List (include, sort, filter), Find
   * Recipient, Match Fields, Check for Errors, and Finish & Merge ▸ Edit
   * Individual Documents.
   */
  import { untrack } from 'svelte';
  import {
    ADDRESS_FIELDS,
    autoFieldMap,
    mailMergeFieldMap,
    mergeFieldErrors,
    mergeToNewDocument,
    recipientListToCsv,
    toUint8Array,
    type RecipientList,
  } from '@office-kit/docx';
  import { commands } from '@office-kit/docx-editor';
  import Dialog from '../Dialog.svelte';
  import { getSession } from '../session.svelte';
  import { t } from '../i18n/index.svelte';
  import { merge } from './merge.svelte';

  const session = getSession();
  const DOCX_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

  function save(name: string, data: BlobPart, type: string): void {
    const url = URL.createObjectURL(new Blob([data], { type }));
    const a = document.createElement('a');
    a.href = url;
    a.download = name;
    a.click();
    URL.revokeObjectURL(url);
  }

  // --- Type a New List ------------------------------------------------------------------
  // Word's default address-list columns.
  const DEFAULT_COLUMNS = ['Title', 'First Name', 'Last Name', 'Company Name', 'Address Line 1', 'Address Line 2', 'City', 'State', 'ZIP Code', 'Country or Region', 'Home Phone', 'Work Phone', 'E-mail Address'];
  const NEW_LIST_FILE = 'My Recipients.csv';
  let columns = $state<string[]>([...DEFAULT_COLUMNS]);
  let rows = $state<Record<string, string>[]>([{}]);
  let newColumn = $state('');

  function addColumn(): void {
    const name = newColumn.trim();
    if (name && !columns.includes(name)) columns.push(name);
    newColumn = '';
  }

  function saveNewList(): boolean {
    const records = rows.filter((r) => columns.some((c) => (r[c] ?? '').trim() !== '')).map((r) => Object.fromEntries(columns.map((c) => [c, r[c] ?? ''])));
    if (records.length === 0) return false;
    const list: RecipientList = { columns: [...columns], records };
    // Word saves a typed list as a data source file; this is that file.
    save(NEW_LIST_FILE, recipientListToCsv(list), 'text/csv');
    merge.attach(session, NEW_LIST_FILE, list);
    rows = [{}];
    return !session.status;
  }

  // --- Edit Recipient List ---------------------------------------------------------------
  type Comparison = 'equals' | 'notEquals' | 'contains' | 'isBlank' | 'isNotBlank';
  const COMPARISONS: readonly Comparison[] = ['equals', 'notEquals', 'contains', 'isBlank', 'isNotBlank'];
  let draftIncluded = $state<boolean[]>([]);
  let order = $state<number[]>([]);
  let sortColumn = $state('');
  let sortDescending = $state(false);
  let filterColumn = $state('');
  let comparison = $state<Comparison>('equals');
  let filterValue = $state('');
  const list = $derived(merge.list);

  $effect(() => {
    if (session.dialog !== 'mailings.recipients') return;
    untrack(() => {
      draftIncluded = [...merge.included];
      order = (merge.list?.records ?? []).map((_, i) => i);
      sortColumn = '';
      filterColumn = merge.list?.columns[0] ?? '';
    });
  });

  function sortBy(column: string): void {
    if (!list) return;
    sortDescending = sortColumn === column ? !sortDescending : false;
    sortColumn = column;
    const sign = sortDescending ? -1 : 1;
    order = order.toSorted((a, b) => sign * (list.records[a]?.[column] ?? '').localeCompare(list.records[b]?.[column] ?? '', undefined, { numeric: true }));
  }

  function matches(value: string): boolean {
    const v = value.trim().toLowerCase();
    const want = filterValue.trim().toLowerCase();
    switch (comparison) {
      case 'equals': return v === want;
      case 'notEquals': return v !== want;
      case 'contains': return v.includes(want);
      case 'isBlank': return v === '';
      case 'isNotBlank': return v !== '';
    }
  }

  function applyFilter(): void {
    if (!list) return;
    draftIncluded = list.records.map((r) => matches(r[filterColumn] ?? ''));
  }

  function saveRecipients(): void {
    if (!list) return;
    // Sorting reorders the list itself, as Word's query ORDER BY does.
    const sorted: RecipientList = { columns: list.columns, records: order.map((i) => list.records[i] ?? {}) };
    const included = order.map((i) => draftIncluded[i] ?? true);
    session.apply(commands.recipientInclusionCommand, { list: sorted, included });
    if (session.status) return;
    merge.list = sorted;
    merge.included = included;
    if (merge.preview) merge.show(session);
  }

  // --- Find Recipient -----------------------------------------------------------------
  let findText = $state('');
  let findColumn = $state('');

  function find(): boolean {
    if (!list) return false;
    const needle = findText.trim().toLowerCase();
    const n = list.records.length;
    for (let step = 1; step <= n; step++) {
      const i = (merge.index + step) % n;
      const record = list.records[i] ?? {};
      const values = findColumn ? [record[findColumn] ?? ''] : Object.values(record);
      if (values.some((v) => v.toLowerCase().includes(needle))) {
        merge.go(session, i);
        return false;
      }
    }
    session.status = t('mail.notFound');
    return false;
  }

  // --- Match Fields ---------------------------------------------------------------------
  let fieldMap = $state<Record<string, string>>({});

  $effect(() => {
    if (session.dialog !== 'mailings.matchFields') return;
    untrack(() => {
      const model = session.model;
      if (!model || !merge.list) return;
      const stored = mailMergeFieldMap(model.doc);
      fieldMap = Object.keys(stored).length > 0 ? stored : autoFieldMap(merge.list.columns);
    });
  });

  function saveFieldMap(): void {
    if (!list) return;
    const map = Object.fromEntries(Object.entries(fieldMap).filter(([, column]) => column !== ''));
    session.apply(commands.matchFieldsCommand, { columns: list.columns, map });
    if (merge.preview) merge.show(session);
  }

  // --- Check for Errors / Merge -----------------------------------------------------------
  const errors = $derived(session.dialog === 'mailings.errors' && session.model && list ? mergeFieldErrors(session.model.doc, list) : []);
  let range = $state<'all' | 'current' | 'fromTo'>('all');
  let from = $state(1);
  let to = $state(1);

  function mergeToNew(): boolean {
    const model = session.model;
    if (!model || !list) return false;
    const bounds = range === 'current' ? { from: merge.index + 1, to: merge.index + 1 } : range === 'fromTo' ? { from, to } : {};
    try {
      const merged = mergeToNewDocument(model.doc, list, { ...bounds, included: merge.included, answers: merge.answers });
      save('Letters1.docx', toUint8Array(merged).slice().buffer, DOCX_TYPE);
      return true;
    } catch (err) {
      // An empty range is the user's choice: report it and keep the dialog open.
      session.status = `${t('mail.finish')}: ${(err as Error).message}`;
      return false;
    }
  }
</script>

<Dialog id="mailings.newList" title={t('mail.newListTitle')} onok={saveNewList}>
  <div class="grid-wrap">
    <table>
      <thead><tr>{#each columns as c (c)}<th>{c}</th>{/each}<th></th></tr></thead>
      <tbody>
        {#each rows as row, r (r)}
          <tr>
            {#each columns as c (c)}<td><input bind:value={row[c]} aria-label={c} /></td>{/each}
            <td><button type="button" class="push" onclick={() => rows.splice(r, 1)} disabled={rows.length === 1}>{t('mail.deleteEntry')}</button></td>
          </tr>
        {/each}
      </tbody>
    </table>
  </div>
  <div class="row">
    <button type="button" class="push" onclick={() => rows.push({})}>{t('mail.newEntry')}</button>
    <input bind:value={newColumn} placeholder={t('mail.addColumn')} aria-label={t('mail.addColumn')} />
    <button type="button" class="push" onclick={addColumn} disabled={!newColumn.trim()}>{t('mail.addColumn')}</button>
  </div>
</Dialog>

<Dialog id="mailings.recipients" title={t('mail.editRecipients')} onok={saveRecipients}>
  {#if list}
    <div class="grid-wrap">
      <table>
        <thead>
          <tr>
            <th></th>
            {#each list.columns as c (c)}<th><button type="button" class="sort" onclick={() => sortBy(c)}>{c}{sortColumn === c ? (sortDescending ? ' ▼' : ' ▲') : ''}</button></th>{/each}
          </tr>
        </thead>
        <tbody>
          {#each order as i (i)}
            <tr>
              <td><input type="checkbox" bind:checked={draftIncluded[i]} aria-label={t('mail.include')} /></td>
              {#each list.columns as c (c)}<td>{list.records[i]?.[c] ?? ''}</td>{/each}
            </tr>
          {/each}
        </tbody>
      </table>
    </div>
    <div class="row">
      <button type="button" class="push" onclick={() => (draftIncluded = list.records.map(() => true))}>{t('mail.selectAll')}</button>
      <button type="button" class="push" onclick={() => (draftIncluded = list.records.map(() => false))}>{t('mail.clearAll')}</button>
    </div>
    <fieldset>
      <legend>{t('mail.filter')}</legend>
      <div class="row">
        <select bind:value={filterColumn} aria-label={t('mail.field')}>
          {#each list.columns as c (c)}<option value={c}>{c}</option>{/each}
        </select>
        <select bind:value={comparison} aria-label={t('mail.comparison')}>
          {#each COMPARISONS as c (c)}<option value={c}>{t(`mail.cmp.${c}`)}</option>{/each}
        </select>
        <input bind:value={filterValue} disabled={comparison === 'isBlank' || comparison === 'isNotBlank'} aria-label={t('mail.compareTo')} />
        <button type="button" class="push" onclick={applyFilter}>{t('mail.applyFilter')}</button>
      </div>
    </fieldset>
  {/if}
</Dialog>

<Dialog id="mailings.find" title={t('mail.findRecipient')} okLabel={t('mail.findNext')} onok={find} okDisabled={!findText.trim()}>
  <label class="field">{t('mail.find')} <input bind:value={findText} /></label>
  <label class="field">{t('mail.lookIn')}
    <select bind:value={findColumn}>
      <option value="">{t('mail.allFields')}</option>
      {#each list?.columns ?? [] as c (c)}<option value={c}>{c}</option>{/each}
    </select>
  </label>
</Dialog>

<Dialog id="mailings.matchFields" title={t('mail.matchFields')} onok={saveFieldMap}>
  <div class="grid-wrap">
    {#each ADDRESS_FIELDS as field (field)}
      <label class="field match">{field}
        <select bind:value={fieldMap[field]}>
          <option value="">{t('mail.notMatched')}</option>
          {#each list?.columns ?? [] as c (c)}<option value={c}>{c}</option>{/each}
        </select>
      </label>
    {/each}
  </div>
</Dialog>

<Dialog id="mailings.errors" title={t('mail.checkErrors')}>
  {#if errors.length === 0}
    <p>{t('mail.noErrors')}</p>
  {:else}
    <p>{t('mail.invalidFields')}</p>
    <ul>{#each errors as e (e)}<li>{e}</li>{/each}</ul>
  {/if}
</Dialog>

<Dialog id="mailings.merge" title={t('mail.mergeToNew')} onok={mergeToNew}>
  <label><input type="radio" bind:group={range} value="all" /> {t('mail.all')}</label>
  <label><input type="radio" bind:group={range} value="current" /> {t('mail.currentRecord')}</label>
  <div class="row">
    <label><input type="radio" bind:group={range} value="fromTo" /> {t('mail.from')}</label>
    <input type="number" min="1" bind:value={from} disabled={range !== 'fromTo'} aria-label={t('mail.from')} style="width: 56px" />
    <label class="field">{t('mail.to')} <input type="number" min="1" bind:value={to} disabled={range !== 'fromTo'} style="width: 56px" /></label>
  </div>
</Dialog>

<style>
  .grid-wrap {
    max-height: 50vh;
    max-width: 680px;
    overflow: auto;
  }
  table {
    border-collapse: collapse;
    font-size: 12px;
  }
  th,
  td {
    padding: 2px 4px;
    border: 1px solid var(--chrome-line);
    white-space: nowrap;
  }
  td input:not([type='checkbox']) {
    width: 110px;
    border: none;
    font: inherit;
  }
  .sort {
    border: none;
    background: none;
    font-weight: 600;
    cursor: pointer;
  }
  .match {
    display: flex;
    justify-content: space-between;
    margin-bottom: 4px;
  }
</style>
