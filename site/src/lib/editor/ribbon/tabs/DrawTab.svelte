<script lang="ts">
  /**
   * Draw: Draw (select) and Lasso Select, Eraser (stroke eraser), the pen
   * gallery (each pen's colour and thickness in its menu), Add Pen, Ink to
   * Shape and Drawing Canvas. Strokes are saved as VML ink shapes.
   */
  import { commands } from '@office-kit/docx-editor';
  import Button from '../Button.svelte';
  import Group from '../Group.svelte';
  import SplitButton from '../SplitButton.svelte';
  import { getSession } from '../../session.svelte';
  import { t, type MessageKey } from '../../i18n/index.svelte';
  import ColorMenu from '../../ColorMenu.svelte';
  import { shapeTools, type Pen, type PenKind } from '../../shapes/tools.svelte';

  const session = getSession();
  // Word's thickness steps per pen type (points).
  const PEN_WEIGHTS = [0.25, 0.5, 1, 2, 3.5];
  const HIGHLIGHTER_WEIGHTS = [4, 6, 8, 10, 12];
  // Word's new drawing canvas: 6" × 3".
  const CANVAS_WIDTH = 432;
  const CANVAS_HEIGHT = 216;
  const PEN_LABEL: Record<PenKind, MessageKey> = { pen: 'draw.pen', pencil: 'draw.pencil', highlighter: 'draw.highlighter' };
  const ADDABLE: readonly PenKind[] = ['pen', 'pencil', 'highlighter'];

  const isOn = (pen: Pen): boolean => shapeTools.tool.kind === 'pen' && shapeTools.tool.pen === pen.id;
  const weights = (pen: Pen): number[] => (pen.kind === 'highlighter' ? HIGHLIGHTER_WEIGHTS : PEN_WEIGHTS);

  function insertCanvas(): void {
    const at = session.apply(commands.insertDrawingCanvasCommand, { width: CANVAS_WIDTH, height: CANVAS_HEIGHT });
    if (at) session.selectedObject = { kind: 'shape', at };
  }
</script>

{#snippet penFace(pen: Pen)}
  <svg width="26" height="56" viewBox="0 0 26 56" aria-hidden="true">
    {#if pen.kind === 'pencil'}
      <path d="M6 2h14v38l-7 14-7-14z" fill="#{pen.color}" stroke="#333" />
      <path d="M6 40h14M9 2v38M17 2v38" stroke="#333" fill="none" />
    {:else if pen.kind === 'highlighter'}
      <path d="M5 2h16v36l-3 6h-10l-3-6z" fill="#{pen.color}" stroke="#333" />
      <path d="M9 44h8v8h-8z" fill="#{pen.color}" stroke="#333" />
    {:else}
      <path d="M6 2h14v36l-7 16-7-16z" fill="#{pen.color}" stroke="#333" />
      <path d="M6 38h14" stroke="#333" />
    {/if}
  </svg>
{/snippet}

<Group label={t('draw.tools')}>
  <Button size="large" icon="drawSelect" tip={t('draw.draw')} on={shapeTools.tool.kind === 'none'} onclick={() => shapeTools.setTool({ kind: 'none' })} />
  <Button size="large" icon="lasso" tip={t('draw.lassoSelect')} on={shapeTools.tool.kind === 'lasso'} onclick={() => shapeTools.toggle({ kind: 'lasso' })} />
  <SplitButton id="draw.eraser" size="large" icon="eraser" tip={t('draw.eraser')} label={t('draw.eraser')} on={shapeTools.tool.kind === 'eraser'} onclick={() => shapeTools.toggle({ kind: 'eraser' })}>
    {#snippet menu()}
      <button class="mi check checked" role="menuitem" onclick={() => { session.openMenu = null; shapeTools.setTool({ kind: 'eraser' }); }}>{t('draw.strokeEraser')}</button>
    {/snippet}
  </SplitButton>
</Group>

<Group label={t('draw.pens')}>
  <div class="wk-pens">
    {#each shapeTools.pens as pen (pen.id)}
      <SplitButton id="draw.pen.{pen.id}" tip={t(PEN_LABEL[pen.kind])} on={isOn(pen)} onclick={() => shapeTools.toggle({ kind: 'pen', pen: pen.id })}>
        {#snippet face()}{@render penFace(pen)}{/snippet}
        {#snippet menu()}
          <div class="menu-head">{t('draw.thickness')}</div>
          {#each weights(pen) as w (w)}
            <button class="mi check" class:checked={pen.weight === w} role="menuitem" onclick={() => shapeTools.updatePen(pen.id, { weight: w })}>
              <span class="wk-line-sample" style="border-top:{Math.max(1, w * 1.33)}px solid #{pen.color}"></span>{w} pt
            </button>
          {/each}
          <ColorMenu onpick={(c) => { shapeTools.updatePen(pen.id, { color: c.rgb }); session.openMenu = null; }}>
            <button class="mi" role="menuitem" disabled={shapeTools.pens.length <= 1} onclick={() => { session.openMenu = null; shapeTools.removePen(pen.id); }}>{t('draw.deletePen')}</button>
          </ColorMenu>
        {/snippet}
      </SplitButton>
    {/each}
  </div>
  <SplitButton id="draw.addPen" size="large" icon="addPen" tip={t('draw.addPen')} label={t('draw.add')}>
    {#snippet menu()}
      {#each ADDABLE as kind (kind)}
        <button class="mi" role="menuitem" onclick={() => { session.openMenu = null; shapeTools.addPen(kind); }}>{t(PEN_LABEL[kind])}</button>
      {/each}
    {/snippet}
  </SplitButton>
</Group>

<Group label={t('draw.convert')}>
  <Button size="large" icon="inkToShape" tip={t('draw.inkToShape')} on={shapeTools.inkToShape} onclick={() => (shapeTools.inkToShape = !shapeTools.inkToShape)} />
</Group>

<Group label={t('draw.insert')}>
  <Button size="large" icon="drawingCanvas" tip={t('draw.drawingCanvas')} onclick={insertCanvas} />
</Group>
