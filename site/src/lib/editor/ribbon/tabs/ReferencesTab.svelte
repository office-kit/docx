<script lang="ts">
  /**
   * Word's References tab: Table of Contents, Footnotes, Citations &
   * Bibliography, Captions, Index and Table of Authorities. The dialogs live
   * beside it under `../../references/`.
   */
  import {
    bibliographySources,
    bibliographyStyle,
    CITATION_STYLES,
    tocLevel,
    type CitationStyle,
    type NoteKind,
    type TableFieldType,
    type TocLevel,
  } from '@office-kit/docx';
  import { caretAt, commands, paragraphAt } from '@office-kit/docx-editor';
  import Button from '../Button.svelte';
  import Group from '../Group.svelte';
  import SplitButton from '../SplitButton.svelte';
  import { getSession } from '../../session.svelte';
  import { t } from '../../i18n/index.svelte';
  import { pageProvider, sourceEditor } from '../../references/state.svelte';
  import ReferencesDialogs from '../../references/ReferencesDialogs.svelte';

  const session = getSession();
  const hasCaret = $derived(session.tick >= 0 && !!session.model?.selection);

  const caretLevel = $derived.by((): TocLevel | undefined => {
    const model = session.model;
    const focus = session.tick >= 0 ? model?.selection?.focus : undefined;
    const paragraph = model && focus ? paragraphAt(model.doc, focus) : undefined;
    return model && paragraph ? tocLevel(model.doc, paragraph) : undefined;
  });

  const sources = $derived(session.tick >= 0 && session.model ? bibliographySources(session.model.doc) : []);
  const style = $derived(session.tick >= 0 && session.model ? bibliographyStyle(session.model.doc) : 'APA');
  const STYLE_NAMES = Object.keys(CITATION_STYLES);
  const BUILT_IN_TOCS = [
    { key: 'ref.autoTable1', title: 'Contents' },
    { key: 'ref.autoTable2', title: 'Table of Contents' },
  ] as const;
  const BIBLIOGRAPHY_GALLERY = [
    { key: 'ref.bibliography', title: 'Bibliography' },
    { key: 'ref.references', title: 'References' },
    { key: 'ref.worksCited', title: 'Works Cited' },
  ] as const;

  function update(types: readonly TableFieldType[]): void {
    const pageOf = pageProvider(session);
    session.apply(commands.updateTablesCommand, { types, ...(pageOf ? { pageOf } : {}) });
  }

  function insertToc(title: string): void {
    const pageOf = pageProvider(session);
    session.apply(commands.insertTocCommand, { title, ...(pageOf ? { pageOf } : {}) });
  }

  function insertNote(kind: NoteKind): void {
    const command = kind === 'footnote' ? commands.addFootnoteCommand : commands.addEndnoteCommand;
    if (session.apply(command, {}) === undefined) return;
    // The note text is typed in the Notes pane until the canvas edits note stories in place.
    session.pane.right = 'notes';
  }

  function goToNote(kind: NoteKind, direction: 'next' | 'previous'): void {
    session.openMenu = null;
    const model = session.model;
    if (!model) return;
    const at = commands.findNoteReference(model, kind, direction);
    if (!at) {
      session.status = t(kind === 'footnote' ? 'ref.noMoreFootnotes' : 'ref.noMoreEndnotes');
      return;
    }
    model.setSelection(caretAt(at));
    session.changed();
  }

  function cite(tag: string): void {
    session.apply(commands.insertCitationCommand, { tag });
  }

  function newSource(citeAfterSave: boolean): void {
    sourceEditor.editing = null;
    sourceEditor.citeAfterSave = citeAfterSave;
    session.openDialog('references.source');
  }

  const isCitationStyle = (value: string): value is CitationStyle => Object.hasOwn(CITATION_STYLES, value);

  function setStyle(e: Event): void {
    const value = (e.currentTarget as HTMLSelectElement).value;
    if (isCitationStyle(value)) session.apply(commands.citationStyleCommand, { style: value });
  }
</script>

<Group label={t('group.toc')}>
  <SplitButton id="ref-toc" size="large" icon="toc" tip={t('ref.toc')}>
    {#snippet menu()}
      <div class="menu-head">{t('ref.builtIn')}</div>
      {#each BUILT_IN_TOCS as item (item.key)}
        <button class="mi" onclick={() => insertToc(item.title)}>{t(item.key)}</button>
      {/each}
      <hr />
      <button class="mi" onclick={() => session.openDialog('references.toc')}>{t('ref.customToc')}</button>
      <button class="mi" onclick={() => session.apply(commands.removeTocCommand, undefined)}>{t('ref.removeToc')}</button>
    {/snippet}
  </SplitButton>
  <div class="col">
    <SplitButton id="ref-addText" size="mid" icon="addText" tip={t('ref.addText')} disabled={!hasCaret}>
      {#snippet menu()}
        <button class="mi check" class:checked={caretLevel === 'none'} onclick={() => session.apply(commands.addTextCommand, { level: 'none' })}>{t('ref.doNotShow')}</button>
        {#each [1, 2, 3] as level (level)}
          <button class="mi check" class:checked={caretLevel === level} onclick={() => session.apply(commands.addTextCommand, { level })}>{t('ref.level')} {level}</button>
        {/each}
      {/snippet}
    </SplitButton>
    <Button size="mid" icon="updateTable" tip={t('ref.updateTable')} onclick={() => update(['TOC'])} />
  </div>
</Group>

<Group label={t('group.footnotes')}>
  <Button size="large" icon="footnote" tip={t('ref.footnote')} onclick={() => insertNote('footnote')} disabled={!hasCaret} />
  <Button size="large" icon="endnote" tip={t('ref.endnote')} onclick={() => insertNote('endnote')} disabled={!hasCaret} />
  <div class="col">
    <SplitButton id="ref-nextNote" size="mid" icon="nextFootnote" tip={t('ref.nextFootnote')} onclick={() => goToNote('footnote', 'next')}>
      {#snippet menu()}
        <button class="mi" onclick={() => goToNote('footnote', 'next')}>{t('ref.nextFootnote')}</button>
        <button class="mi" onclick={() => goToNote('footnote', 'previous')}>{t('ref.previousFootnote')}</button>
        <button class="mi" onclick={() => goToNote('endnote', 'next')}>{t('ref.nextEndnote')}</button>
        <button class="mi" onclick={() => goToNote('endnote', 'previous')}>{t('ref.previousEndnote')}</button>
        <!-- Word for Mac has no ribbon button for the dialog; it sits with the note commands. -->
        <hr />
        <button class="mi" onclick={() => { session.openMenu = null; session.openDialog('references.notes'); }}>{t('ref.noteDialog')}</button>
      {/snippet}
    </SplitButton>
    <Button size="mid" icon="showNotes" tip={t('ref.showNotes')} on={session.pane.right === 'notes'} onclick={() => session.togglePane('right', 'notes')} />
  </div>
</Group>

<Group label={t('group.citations')}>
  <SplitButton id="ref-cite" size="large" icon="citation" tip={t('ref.insertCitation')} disabled={!hasCaret}>
    {#snippet menu()}
      {#each sources as source (source.tag)}
        <button class="mi" onclick={() => cite(source.tag)}>{source.fields?.Title ?? source.tag}</button>
      {/each}
      {#if sources.length > 0}<hr />{/if}
      <button class="mi" onclick={() => newSource(true)}>{t('ref.addNewSource')}</button>
      <button class="mi" onclick={() => session.openDialog('references.placeholder')}>{t('ref.addPlaceholder')}</button>
      <hr />
      <button class="mi" onclick={() => session.openDialog('references.sources')}>{t('ref.manageSources')}</button>
    {/snippet}
  </SplitButton>
  <Button size="large" icon="citations" tip={t('ref.citations')} on={session.pane.right === 'citations'} onclick={() => session.togglePane('right', 'citations')} />
  <div class="col">
    <label class="field" title={t('ref.style')}>
      <select value={style} onchange={setStyle} aria-label={t('ref.style')} style="width: 120px">
        {#each STYLE_NAMES as name (name)}<option value={name}>{name}</option>{/each}
      </select>
    </label>
    <SplitButton id="ref-bibliography" size="mid" icon="bibliography" tip={t('ref.bibliographyMenu')}>
      {#snippet menu()}
        <div class="menu-head">{t('ref.builtIn')}</div>
        {#each BIBLIOGRAPHY_GALLERY as item (item.key)}
          <button class="mi" onclick={() => session.apply(commands.insertBibliographyCommand, { title: item.title })}>{t(item.key)}</button>
        {/each}
        <hr />
        <button class="mi" onclick={() => session.apply(commands.insertBibliographyCommand, {})}>{t('ref.insertBibliography')}</button>
      {/snippet}
    </SplitButton>
  </div>
</Group>

<Group label={t('group.captions')}>
  <Button size="large" icon="caption" tip={t('ref.insertCaption')} onclick={() => session.openDialog('references.caption')} disabled={!hasCaret} />
  <Button size="large" icon="tableOfFigures" tip={t('ref.tableOfFigures')} onclick={() => session.openDialog('references.tof')} />
  <div class="col">
    <Button size="mid" icon="updateTable" tip={t('ref.updateTable')} onclick={() => update(['TOC'])} />
    <Button size="mid" icon="crossReference" tip={t('ref.crossReference')} onclick={() => session.openDialog('insert.crossReference')} disabled={!hasCaret} />
  </div>
</Group>

<Group label={t('group.index')}>
  <Button size="large" icon="markEntry" tip={t('ref.markEntry')} onclick={() => session.openDialog('references.markEntry')} disabled={!hasCaret} />
  <div class="col">
    <Button size="mid" icon="insertIndex" tip={t('ref.insertIndex')} onclick={() => session.openDialog('references.index')} />
    <Button size="mid" icon="updateTable" tip={t('ref.updateIndex')} onclick={() => update(['INDEX'])} />
  </div>
</Group>

<Group label={t('group.toa')}>
  <Button size="large" icon="markCitation" tip={t('ref.markCitation')} onclick={() => session.openDialog('references.markCitation')} disabled={!hasCaret} />
  <div class="col">
    <Button size="mid" icon="toa" tip={t('ref.insertToa')} onclick={() => session.openDialog('references.toa')} />
    <Button size="mid" icon="updateTable" tip={t('ref.updateToa')} onclick={() => update(['TOA'])} />
  </div>
</Group>

<ReferencesDialogs />

<style>
  /* Keep icons whole when the ribbon is narrower than the tab. */
  .col :global(svg) {
    flex-shrink: 0;
  }
</style>
