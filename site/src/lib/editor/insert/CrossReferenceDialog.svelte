<script lang="ts">
  /**
   * Insert ▸ Cross-reference: pick a reference type and an item, then what
   * to insert — a REF / PAGEREF / NOTEREF field with Word's switches (`\r`,
   * `\n`, `\w`, `\p`, `\f`, and `\h` for "Insert as hyperlink").
   */
  import {
    commands,
    documentBookmarks,
    documentCaptions,
    documentHeadings,
    documentNotes,
    documentNumberedItems,
    type DocumentTarget,
    type NoteTarget,
  } from '@office-kit/docx-editor';
  import Dialog from '../Dialog.svelte';
  import { getSession } from '../session.svelte';
  import { t } from '../i18n/index.svelte';
  import { DIALOG } from './state.svelte';

  const session = getSession();

  type RefType = 'heading' | 'bookmark' | 'footnote' | 'endnote' | 'numberedItem' | 'figure' | 'table' | 'equation';
  type RefKind =
    | 'text'
    | 'page'
    | 'number'
    | 'numberNoContext'
    | 'numberFullContext'
    | 'aboveBelow'
    | 'noteNumber'
    | 'noteNumberFormatted'
    | 'caption';

  const TYPES = ['heading', 'bookmark', 'footnote', 'endnote', 'numberedItem', 'figure', 'table', 'equation'] as const satisfies readonly RefType[];
  // Word's "Insert reference to" choices per type, first one the default.
  const PARAGRAPH_KINDS = ['text', 'page', 'number', 'numberNoContext', 'numberFullContext', 'aboveBelow'] as const;
  const KINDS: Record<RefType, readonly RefKind[]> = {
    heading: PARAGRAPH_KINDS,
    bookmark: PARAGRAPH_KINDS,
    numberedItem: ['number', 'page', 'text', 'numberNoContext', 'numberFullContext', 'aboveBelow'],
    footnote: ['noteNumber', 'page', 'aboveBelow', 'noteNumberFormatted'],
    endnote: ['noteNumber', 'page', 'aboveBelow', 'noteNumberFormatted'],
    figure: ['caption', 'page', 'aboveBelow'],
    table: ['caption', 'page', 'aboveBelow'],
    equation: ['caption', 'page', 'aboveBelow'],
  };
  const KIND_KEYS = {
    text: 'ins.xref.r.text',
    page: 'ins.xref.r.page',
    number: 'ins.xref.r.number',
    numberNoContext: 'ins.xref.r.numberNoContext',
    numberFullContext: 'ins.xref.r.numberFullContext',
    aboveBelow: 'ins.xref.r.aboveBelow',
    noteNumber: 'ins.xref.r.noteNumber',
    noteNumberFormatted: 'ins.xref.r.noteNumberFormatted',
    caption: 'ins.xref.r.caption',
  } as const satisfies Record<RefKind, string>;
  // SEQ identifiers Word's captions use for each caption label.
  const CAPTION_LABEL = { figure: 'Figure', table: 'Table', equation: 'Equation' } as const;

  let type = $state<RefType>('heading');
  let kind = $state<RefKind>('text');
  let asLink = $state(true);
  let aboveBelow = $state(false);
  let selected = $state(0);

  type Item = DocumentTarget | NoteTarget | { label: string; name: string };
  const items = $derived.by((): Item[] => {
    const doc = session.dialog === DIALOG.crossReference && session.version >= 0 ? session.model?.doc : undefined;
    if (!doc) return [];
    switch (type) {
      case 'heading':
        return documentHeadings(doc);
      case 'bookmark':
        return documentBookmarks(doc).filter((b) => !b.hidden);
      case 'footnote':
      case 'endnote':
        return documentNotes(doc, type);
      case 'numberedItem':
        return documentNumberedItems(doc);
      case 'figure':
      case 'table':
      case 'equation':
        return documentCaptions(doc, CAPTION_LABEL[type]);
    }
  });

  function chooseType(next: RefType): void {
    type = next;
    kind = KINDS[next][0] ?? 'text';
    selected = 0;
  }

  /** The field and switches a choice inserts. */
  function field(): { field: 'REF' | 'PAGEREF' | 'NOTEREF'; switches: string[] } {
    switch (kind) {
      case 'page':
        return { field: 'PAGEREF', switches: [] };
      case 'number':
        return { field: 'REF', switches: ['\\r'] };
      case 'numberNoContext':
        return { field: 'REF', switches: ['\\n'] };
      case 'numberFullContext':
        return { field: 'REF', switches: ['\\w'] };
      case 'aboveBelow':
        return type === 'footnote' || type === 'endnote' ? { field: 'NOTEREF', switches: ['\\p'] } : { field: 'REF', switches: ['\\p'] };
      case 'noteNumber':
        return { field: 'NOTEREF', switches: [] };
      case 'noteNumberFormatted':
        return { field: 'NOTEREF', switches: ['\\f'] };
      case 'text':
      case 'caption':
        return { field: 'REF', switches: [] };
    }
  }

  function insert(): boolean {
    const item = items[selected];
    if (!item) return false;
    const target =
      'name' in item
        ? { kind: 'bookmark' as const, name: item.name }
        : 'run' in item
          ? { kind: 'note' as const, paragraph: item.paragraph, run: item.run }
          : { kind: 'paragraph' as const, paragraph: item.paragraph };
    session.apply(commands.insertCrossReferenceCommand, {
      target,
      ...field(),
      hyperlink: asLink,
      aboveBelow: aboveBelow && kind !== 'aboveBelow',
    });
    return session.status === '';
  }
</script>

<Dialog id={DIALOG.crossReference} title={t('ins.crossReference')} onok={insert} okLabel={t('ins.sym.insert')} okDisabled={items.length === 0}>
  <div class="dialog-row">
    <label class="stack">{t('ins.xref.type')}
      <select value={type} onchange={(e) => { const next = TYPES.find((ty) => ty === e.currentTarget.value); if (next) chooseType(next); }}>
        {#each TYPES as ty (ty)}<option value={ty}>{t(`ins.xref.t.${ty}`)}</option>{/each}
      </select>
    </label>
    <label class="stack">{t('ins.xref.insertRef')}
      <select bind:value={kind}>
        {#each KINDS[type] as k (k)}<option value={k}>{t(KIND_KEYS[k])}</option>{/each}
      </select>
    </label>
  </div>
  <div class="dialog-row">
    <label><input type="checkbox" bind:checked={asLink} /> {t('ins.xref.asLink')}</label>
    <label><input type="checkbox" bind:checked={aboveBelow} disabled={kind === 'aboveBelow'} /> {t('ins.xref.aboveBelow')}</label>
  </div>
  <span>{t('ins.xref.forWhich')}</span>
  <div class="listbox" role="listbox" aria-label={t('ins.xref.forWhich')}>
    {#each items as item, i (i)}
      <button type="button" role="option" aria-selected={i === selected} class:selected={i === selected} onclick={() => (selected = i)} ondblclick={() => { selected = i; insert(); }}>{item.label}</button>
    {:else}
      <div class="group-label">{t('ins.xref.none')}</div>
    {/each}
  </div>
  {#if session.status}<p class="error-note">{session.status}</p>{/if}
</Dialog>
