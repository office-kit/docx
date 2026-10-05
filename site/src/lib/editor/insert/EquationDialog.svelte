<script lang="ts">
  /**
   * Insert New Equation / Edit Equation: the equation in linear format
   * (UnicodeMath, as Word's Linear mode), a live preview, and the structure
   * templates of Word's Equation tab inserted at the text cursor.
   */
  import { buildEquation, paragraphText } from '@office-kit/docx';
  import { commands, paragraphAt, renderMath } from '@office-kit/docx-editor';
  import Dialog from '../Dialog.svelte';
  import { getSession } from '../session.svelte';
  import { t } from '../i18n/index.svelte';
  import { STRUCTURES } from './equations';
  import { DIALOG, insertUi } from './state.svelte';

  const session = getSession();
  let linear = $state('');
  let display = $state(true);
  let structure = $state(0);
  let input = $state<HTMLTextAreaElement | null>(null);

  $effect(() => {
    if (session.dialog !== DIALOG.equation) return;
    linear = insertUi.equation?.linear ?? '';
    // Word makes a display equation when the paragraph is empty, inline otherwise.
    const model = session.model;
    const para = model?.selection ? paragraphAt(model.doc, model.selection.focus) : undefined;
    display = !para || paragraphText(para) === '';
  });

  // The preview is markup the library renders from its own OMML (escaped
  // there), so it is safe for {@html}. A half-typed equation that does not
  // parse yet shows the parser's message instead.
  const preview = $derived.by((): { html: string } | { error: string } => {
    if (!linear.trim()) return { html: '' };
    try {
      return { html: renderMath(buildEquation(linear, { display: true })) };
    } catch (err) {
      return { error: err instanceof Error ? err.message : String(err) };
    }
  });

  function insertTemplate(template: string): void {
    const el = input;
    const start = el?.selectionStart ?? linear.length;
    const end = el?.selectionEnd ?? linear.length;
    linear = linear.slice(0, start) + template + linear.slice(end);
    queueMicrotask(() => {
      el?.focus();
      el?.setSelectionRange(start + template.length, start + template.length);
    });
  }

  function ok(): boolean {
    const target = insertUi.equation;
    if (target) session.apply(commands.editEquationCommand, { ...target, linear });
    else session.apply(commands.insertEquationCommand, { linear, display });
    if (session.status !== '') return false;
    insertUi.equation = null;
    return true;
  }
</script>

<Dialog id={DIALOG.equation} title={insertUi.equation ? t('ins.eq.edit') : t('ins.eq.insertNew')} onok={ok} okDisabled={!linear.trim() || 'error' in preview}>
  <div class="structure-bar" role="tablist" aria-label={t('ins.eq.structures')}>
    {#each STRUCTURES as s, i (s.key)}
      <button type="button" role="tab" class="push" class:on={i === structure} aria-selected={i === structure} onclick={() => (structure = i)}>{t(s.key)}</button>
    {/each}
  </div>
  <div class="structure-bar">
    {#each STRUCTURES[structure]?.templates ?? [] as template (template)}
      <button type="button" class="push" onclick={() => insertTemplate(template)}>{template}</button>
    {/each}
  </div>
  <label class="stack">{t('ins.eq.linear')}<textarea rows="3" bind:value={linear} bind:this={input} spellcheck="false"></textarea></label>
  <span>{t('ins.eq.preview')}</span>
  <div class="math-preview">
    {#if 'html' in preview}{@html preview.html}{:else}<p class="error-note">{preview.error}</p>{/if}
  </div>
  {#if !insertUi.equation}<label><input type="checkbox" bind:checked={display} /> {t('ins.eq.display')}</label>{/if}
  {#if session.status}<p class="error-note">{session.status}</p>{/if}
</Dialog>
