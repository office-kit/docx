<script lang="ts">
  /**
   * Insert ▸ Link (and Edit Hyperlink): text to display, and a target that is
   * a web address, a place in this document (top, a heading, a bookmark) or
   * an email address, plus ScreenTip and target frame (`w:tooltip`,
   * `w:tgtFrame`). Opened on a link, it edits that link and offers Remove Link.
   */
  import {
    commands,
    documentBookmarks,
    documentHeadings,
    hyperlinkAtCaret,
    selectionText,
    type HeadingTarget,
  } from '@office-kit/docx-editor';
  import Dialog from '../Dialog.svelte';
  import { getSession } from '../session.svelte';
  import { t } from '../i18n/index.svelte';
  import { DIALOG } from './state.svelte';

  const session = getSession();
  type Tab = 'web' | 'document' | 'email';
  // Word's place for the start of the document.
  const TOP_ANCHOR = '_top';
  const FRAMES = [
    { value: '', key: 'ins.frame.none' },
    { value: '_blank', key: 'ins.frame.blank' },
    { value: '_self', key: 'ins.frame.self' },
    { value: '_top', key: 'ins.frame.top' },
    { value: '_parent', key: 'ins.frame.parent' },
  ] as const;

  let tab = $state<Tab>('web');
  let text = $state('');
  let address = $state('https://');
  let email = $state('');
  let subject = $state('');
  let tooltip = $state('');
  let frame = $state('');
  let place = $state<{ bookmark: string } | { heading: HeadingTarget } | null>(null);
  let editing = $state<number | null>(null);
  let error = $state('');

  const headings = $derived(session.dialog === DIALOG.link && session.model ? documentHeadings(session.model.doc) : []);
  const bookmarks = $derived(
    session.dialog === DIALOG.link && session.model ? documentBookmarks(session.model.doc).filter((b) => !b.hidden) : [],
  );

  // Fill the form each time the dialog opens: from the link at the caret when
  // there is one, otherwise from the selected text.
  $effect(() => {
    if (session.dialog !== DIALOG.link) return;
    const model = session.model;
    if (!model) return;
    error = '';
    const found = hyperlinkAtCaret(model);
    editing = found ? found.index : null;
    text = found ? found.link.text : selectionText(model);
    tooltip = found?.link.tooltip ?? '';
    frame = found?.link.targetFrame ?? '';
    place = null;
    const url = found?.link.url;
    if (url?.startsWith('mailto:')) {
      tab = 'email';
      const [to = '', query = ''] = url.slice('mailto:'.length).split('?');
      email = decodeURIComponent(to);
      subject = new URLSearchParams(query).get('subject') ?? '';
    } else if (url) {
      tab = 'web';
      address = url;
    } else if (found?.link.bookmark) {
      tab = 'document';
      place = { bookmark: found.link.bookmark };
    } else {
      tab = 'web';
      address = 'https://';
    }
  });

  function target(): { target: { url?: string; bookmark?: string; tooltip?: string; targetFrame?: string }; heading?: HeadingTarget['paragraph'] } {
    const extra = { ...(tooltip ? { tooltip } : {}), ...(frame ? { targetFrame: frame } : {}) };
    if (tab === 'email') {
      const query = subject ? `?subject=${encodeURIComponent(subject)}` : '';
      return { target: { url: `mailto:${email.trim()}${query}`, ...extra } };
    }
    if (tab === 'document') {
      if (place && 'heading' in place) return { target: { ...extra }, heading: place.heading.paragraph };
      return { target: { bookmark: place?.bookmark ?? TOP_ANCHOR, ...extra } };
    }
    return { target: { url: address.trim(), ...extra } };
  }

  function display(): string {
    if (text.trim()) return text;
    if (tab === 'email') return email.trim();
    if (tab === 'web') return address.trim();
    if (place && 'heading' in place) return place.heading.label;
    return place?.bookmark ?? TOP_ANCHOR;
  }

  function ok(): boolean {
    const { target: tgt, heading } = target();
    const params = { text: display(), target: tgt, ...(heading ? { heading } : {}) };
    if (editing === null) session.apply(commands.insertLinkCommand, params);
    else session.apply(commands.editLinkCommand, { ...params, index: editing });
    // apply() reports a rejected command in the status bar; keep the dialog
    // open and show the reason instead of closing on bad input.
    error = session.status;
    return error === '';
  }

  function removeLink(): void {
    if (editing === null) return;
    session.apply(commands.removeLinkCommand, { index: editing });
    session.dialog = null;
  }

  function placeKey(p: typeof place): string {
    if (p === null) return 'top';
    return 'heading' in p ? `h:${p.heading.block}` : `b:${p.bookmark}`;
  }
  const placeSelected = (p: typeof place): boolean => placeKey(p) === placeKey(place);
</script>

<Dialog id={DIALOG.link} title={editing === null ? t('ins.link.insertTitle') : t('ins.link.editTitle')} onok={ok}>
  <label class="stack">{t('ins.link.textToDisplay')}<input type="text" bind:value={text} /></label>
  <div class="dialog-tabs" role="tablist">
    <button type="button" role="tab" class:on={tab === 'web'} aria-selected={tab === 'web'} onclick={() => (tab = 'web')}>{t('ins.link.webPage')}</button>
    <button type="button" role="tab" class:on={tab === 'document'} aria-selected={tab === 'document'} onclick={() => (tab = 'document')}>{t('ins.link.thisDocument')}</button>
    <button type="button" role="tab" class:on={tab === 'email'} aria-selected={tab === 'email'} onclick={() => (tab = 'email')}>{t('ins.link.email')}</button>
  </div>
  {#if tab === 'web'}
    <label class="stack">{t('ins.link.address')}<input type="text" bind:value={address} /></label>
  {:else if tab === 'document'}
    <span>{t('ins.link.place')}</span>
    <div class="listbox" role="listbox" aria-label={t('ins.link.place')}>
      <button type="button" role="option" aria-selected={place === null} class:selected={place === null} onclick={() => (place = null)}>{t('ins.link.top')}</button>
      {#if headings.length}<div class="group-label">{t('ins.link.headings')}</div>{/if}
      {#each headings as h (h.block)}
        {@const p = { heading: h }}
        <button type="button" role="option" aria-selected={placeSelected(p)} class:selected={placeSelected(p)} style="padding-left: {h.level * 12}px" onclick={() => (place = p)}>{h.label}</button>
      {/each}
      {#if bookmarks.length}<div class="group-label">{t('ins.link.bookmarks')}</div>{/if}
      {#each bookmarks as b (b.name)}
        {@const p = { bookmark: b.name }}
        <button type="button" role="option" aria-selected={placeSelected(p)} class:selected={placeSelected(p)} onclick={() => (place = p)}>{b.name}</button>
      {/each}
    </div>
  {:else}
    <label class="stack">{t('ins.link.emailAddress')}<input type="email" bind:value={email} /></label>
    <label class="stack">{t('ins.link.subject')}<input type="text" bind:value={subject} /></label>
  {/if}
  <div class="dialog-row">
    <label class="stack">{t('ins.link.screenTip')}<input type="text" bind:value={tooltip} /></label>
    <label class="stack">{t('ins.link.targetFrame')}
      <select bind:value={frame}>
        {#each FRAMES as f (f.value)}<option value={f.value}>{t(f.key)}</option>{/each}
      </select>
    </label>
  </div>
  {#if error}<p class="error-note">{error}</p>{/if}
  {#if editing !== null}<div><button type="button" class="push" onclick={removeLink}>{t('ins.link.remove')}</button></div>{/if}
</Dialog>
