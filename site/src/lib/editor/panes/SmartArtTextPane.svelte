<script lang="ts">
  /**
   * SmartArt's Text Pane: the diagram's bullets as an outline. Enter adds a
   * shape, Tab / Shift+Tab demote and promote, Backspace in an empty row
   * removes it; every change rewrites the diagram's data model.
   */
  import { getSmartArt } from '@office-kit/docx';
  import { applyOutlineOp, commands, outlineRows, outlineToNodes, type OutlineOp, type OutlineRow } from '@office-kit/docx-editor';
  import RibbonIcon from '../RibbonIcon.svelte';
  import { getSession } from '../session.svelte';
  import { t } from '../i18n/index.svelte';
  import { selectedSmartArt } from '../shapes/selection';
  import { shapeTools } from '../shapes/tools.svelte';

  const session = getSession();
  const INDENT_PX = 16;
  const current = $derived(selectedSmartArt(session));
  const rows = $derived.by((): OutlineRow[] => {
    const model = session.model;
    return current && model ? outlineRows(getSmartArt(model.doc, current.ref).nodes) : [];
  });
  let inputs = $state<HTMLInputElement[]>([]);

  function save(next: readonly OutlineRow[], focus?: number): void {
    if (!current) return;
    session.apply(commands.smartArtNodesCommand, { at: current.at, nodes: outlineToNodes(next) });
    if (focus === undefined) return;
    shapeTools.smartArtRow = focus;
    requestAnimationFrame(() => inputs[focus]?.focus());
  }

  function run(index: number, op: OutlineOp): void {
    const result = applyOutlineOp(rows, index, op, '');
    save(result.rows, result.index);
  }

  function edit(index: number, text: string): void {
    save(rows.map((r, i) => (i === index ? { ...r, text } : r)));
  }

  function onkeydown(e: KeyboardEvent, index: number): void {
    if (e.isComposing) return;
    if (e.key === 'Enter') {
      e.preventDefault();
      run(index, 'addAfter');
    } else if (e.key === 'Tab') {
      e.preventDefault();
      run(index, e.shiftKey ? 'promote' : 'demote');
    } else if (e.key === 'Backspace' && rows[index]?.text === '' && rows.length > 1) {
      e.preventDefault();
      save(
        rows.filter((_, i) => i !== index),
        Math.max(0, index - 1),
      );
    }
  }
</script>

<div class="pane-head">
  <span>{t('draw.smartArt.textPane')}</span>
  <button class="pane-close" onclick={() => (session.pane.left = null)} aria-label={t('draw.close')}><RibbonIcon name="close" size={14} /></button>
</div>
{#if current}
  <p class="pane-note">{t('draw.smartArt.typeHere')}</p>
  <ul class="wk-smartart-outline">
    {#each rows as row, i (i)}
      <li style="padding-left:{row.level * INDENT_PX}px">
        <span aria-hidden="true">•</span>
        <input
          class="pane-input"
          bind:this={inputs[i]}
          value={row.text}
          placeholder={t('draw.smartArt.placeholder')}
          aria-label={t('draw.smartArt.textPane')}
          onfocus={() => (shapeTools.smartArtRow = i)}
          onchange={(e) => edit(i, e.currentTarget.value)}
          onkeydown={(e) => onkeydown(e, i)}
        />
      </li>
    {/each}
  </ul>
  {#if rows.length === 0}
    <button class="pane-button" onclick={() => save([{ level: 0, text: '' }], 0)}>{t('draw.smartArt.addShape')}</button>
  {/if}
{:else}
  <p class="pane-note">{t('draw.smartArt.selectFirst')}</p>
{/if}

<style>
  .wk-smartart-outline {
    margin: 0;
    padding: 0 8px;
    list-style: none;
  }
  .wk-smartart-outline li {
    display: flex;
    align-items: center;
    gap: 4px;
  }
  .wk-smartart-outline .pane-input {
    flex: 1;
    margin: 2px 0;
  }
</style>
