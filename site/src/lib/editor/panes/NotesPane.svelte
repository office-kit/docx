<script lang="ts">
  /**
   * Show Notes: every footnote and endnote with its reference mark, editable
   * in place. Word scrolls to the notes area of the page; this pane is the
   * same list beside the page.
   */
  import { noteMarks, noteText, type NoteKind } from '@office-kit/docx';
  import { commands } from '@office-kit/docx-editor';
  import RibbonIcon from '../RibbonIcon.svelte';
  import { getSession } from '../session.svelte';
  import { t } from '../i18n/index.svelte';

  const session = getSession();
  const KINDS: readonly NoteKind[] = ['footnote', 'endnote'];

  const notes = $derived.by(() => {
    const model = session.model;
    if (session.tick < 0 || !model) return [];
    return KINDS.flatMap((kind) =>
      noteMarks(model.doc, kind).map((m) => ({ kind, id: m.id, mark: m.mark, text: noteText(model.doc, kind, m.id) ?? '' })),
    );
  });

  function save(kind: NoteKind, id: number, e: Event): void {
    const text = (e.currentTarget as HTMLTextAreaElement).value;
    session.apply(commands.noteTextCommand, { kind, id, text });
  }
</script>

<div class="pane-head">
  <span>{t('ref.showNotes')}</span>
  <button class="pane-close" onclick={() => (session.pane.right = null)} aria-label={t('ref.closePane')}><RibbonIcon name="close" size={14} /></button>
</div>
{#each KINDS as kind (kind)}
  {@const list = notes.filter((n) => n.kind === kind)}
  {#if list.length > 0}
    <div class="menu-head">{t(kind === 'footnote' ? 'ref.footnotes' : 'ref.endnotes')}</div>
    {#each list as note (note.id)}
      <label class="note">
        <sup>{note.mark}</sup>
        <textarea class="pane-input" rows="2" value={note.text} onchange={(e) => save(kind, note.id, e)}></textarea>
      </label>
    {/each}
  {/if}
{/each}
{#if notes.length === 0}<p class="pane-note">{t('ref.noNotes')}</p>{/if}

<style>
  .note {
    display: flex;
    gap: 6px;
    align-items: flex-start;
  }
  .note textarea {
    flex: 1;
    height: auto;
    padding: 4px 6px;
    resize: vertical;
  }
</style>
