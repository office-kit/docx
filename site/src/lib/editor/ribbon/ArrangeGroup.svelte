<script lang="ts">
  /**
   * Word's Arrange group, shared by Picture Format, Shape Format and Layout.
   * Every command goes through the adapter for the selected object's kind
   * (`ARRANGE_OPS`); with nothing selected only Selection Pane is available.
   */
  import Button from './Button.svelte';
  import Group from './Group.svelte';
  import SplitButton from './SplitButton.svelte';
  import LayoutOptionsDialog from '../picture/LayoutOptionsDialog.svelte';
  import { getSession } from '../session.svelte';
  import { t, type MessageKey } from '../i18n/index.svelte';
  import { arrangeOpsFor, type AlignOp, type ArrangeOrder, type ArrangeWrap, type PositionPreset, type RotateOp } from './arrange';

  const session = getSession();
  const ops = $derived(session.tick >= 0 ? arrangeOpsFor(session) : undefined);
  const wrap = $derived(ops?.wrap(session));
  const moveWithText = $derived(ops?.moveWithText?.(session));
  // Align ▸ Align to Page / Align to Margin is a mode, remembered like Word does.
  let alignTo = $state<'page' | 'margin'>('margin');

  const WRAPS: ReadonlyArray<readonly [ArrangeWrap, MessageKey]> = [
    ['inline', 'arr.wrapInline'],
    ['square', 'arr.wrapSquare'],
    ['tight', 'arr.wrapTight'],
    ['through', 'arr.wrapThrough'],
    ['topAndBottom', 'arr.wrapTopBottom'],
    ['behindText', 'arr.wrapBehind'],
    ['inFrontOfText', 'arr.wrapInFront'],
  ];
  const PRESETS: readonly Exclude<PositionPreset, 'inline'>[] = [
    'topLeft', 'topCenter', 'topRight', 'middleLeft', 'middleCenter', 'middleRight', 'bottomLeft', 'bottomCenter', 'bottomRight',
  ];
  const FORWARD: ReadonlyArray<readonly [ArrangeOrder, MessageKey]> = [
    ['bringForward', 'arr.bringForward'],
    ['bringToFront', 'arr.bringToFront'],
    ['bringInFrontOfText', 'arr.bringInFrontOfText'],
  ];
  const BACKWARD: ReadonlyArray<readonly [ArrangeOrder, MessageKey]> = [
    ['sendBackward', 'arr.sendBackward'],
    ['sendToBack', 'arr.sendToBack'],
    ['sendBehindText', 'arr.sendBehindText'],
  ];
  const ALIGNS: ReadonlyArray<readonly [AlignOp, MessageKey]> = [
    ['left', 'arr.alignLeft'],
    ['center', 'arr.alignCenter'],
    ['right', 'arr.alignRight'],
    ['top', 'arr.alignTop'],
    ['middle', 'arr.alignMiddle'],
    ['bottom', 'arr.alignBottom'],
    ['distributeH', 'arr.distributeH'],
    ['distributeV', 'arr.distributeV'],
  ];
  const ROTATES: ReadonlyArray<readonly [RotateOp, MessageKey]> = [
    ['right90', 'arr.rotateRight'],
    ['left90', 'arr.rotateLeft'],
    ['flipV', 'arr.flipV'],
    ['flipH', 'arr.flipH'],
  ];

  /** A miniature page with the object's square in the preset's cell. */
  function presetCell(p: Exclude<PositionPreset, 'inline'>): { x: number; y: number } {
    const i = PRESETS.indexOf(p);
    return { x: 3 + (i % 3) * 4.5, y: 3 + Math.floor(i / 3) * 5.5 };
  }

  function run(f: (o: NonNullable<typeof ops>) => void): void {
    session.openMenu = null;
    if (ops) f(ops);
  }
</script>

<Group label={t('arr.group')}>
  <SplitButton id="arr.position" size="large" icon="arrPosition" tip={t('arr.position')} disabled={!ops}>
    {#snippet menu()}
      <div class="menu-head">{t('arr.inLine')}</div>
      <button class="mi" role="menuitem" onclick={() => run((o) => o.position(session, 'inline'))}>{t('arr.wrapInline')}</button>
      <div class="menu-head">{t('arr.withWrapping')}</div>
      <div class="grid arr-presets">
        {#each PRESETS as p (p)}
          {@const c = presetCell(p)}
          <button class="chip arr-preset" title={t(`arr.pos.${p}`)} aria-label={t(`arr.pos.${p}`)} onclick={() => run((o) => o.position(session, p))}>
            <svg viewBox="0 0 20 24" width="30" height="36" aria-hidden="true">
              <rect x="1" y="1" width="18" height="22" fill="#fff" stroke="#999" />
              {#each [5, 8, 11, 14, 17, 20] as y (y)}<path d="M3 {y}h14" stroke="#c8c8c8" />{/each}
              <rect x={c.x} y={c.y} width="5" height="5" fill="#4472C4" />
            </svg>
          </button>
        {/each}
      </div>
      <hr />
      <button class="mi" role="menuitem" disabled={!ops?.layoutOptions} onclick={() => run((o) => o.layoutOptions?.(session))}>{t('arr.moreLayout')}</button>
    {/snippet}
  </SplitButton>
  <SplitButton id="arr.wrap" size="large" icon="arrWrapText" tip={t('arr.wrapText')} disabled={!ops}>
    {#snippet menu()}
      {#each WRAPS as [w, key] (w)}
        <button class="mi check" class:checked={wrap === w} role="menuitemradio" aria-checked={wrap === w} onclick={() => run((o) => o.setWrap(session, w))}>{t(key)}</button>
      {/each}
      <hr />
      <button class="mi" role="menuitem" disabled={!ops?.editWrapPoints || !wrap || wrap === 'inline'} onclick={() => run((o) => o.editWrapPoints?.(session))}>{t('arr.editWrapPoints')}</button>
      <button class="mi check" class:checked={moveWithText === true} role="menuitemradio" aria-checked={moveWithText === true} disabled={moveWithText === undefined} onclick={() => run((o) => o.setMoveWithText?.(session, true))}>{t('arr.moveWithText')}</button>
      <button class="mi check" class:checked={moveWithText === false} role="menuitemradio" aria-checked={moveWithText === false} disabled={moveWithText === undefined} onclick={() => run((o) => o.setMoveWithText?.(session, false))}>{t('arr.fixPosition')}</button>
      <hr />
      <button class="mi" role="menuitem" disabled={!ops?.layoutOptions} onclick={() => run((o) => o.layoutOptions?.(session))}>{t('arr.moreLayout')}</button>
    {/snippet}
  </SplitButton>
  <div class="col">
    <SplitButton id="arr.forward" size="mid" icon="arrBringForward" tip={t('arr.bringForward')} disabled={!ops} onclick={() => run((o) => o.order(session, 'bringForward'))}>
      {#snippet menu()}
        {#each FORWARD as [op, key] (op)}<button class="mi" role="menuitem" onclick={() => run((o) => o.order(session, op))}>{t(key)}</button>{/each}
      {/snippet}
    </SplitButton>
    <SplitButton id="arr.backward" size="mid" icon="arrSendBackward" tip={t('arr.sendBackward')} disabled={!ops} onclick={() => run((o) => o.order(session, 'sendBackward'))}>
      {#snippet menu()}
        {#each BACKWARD as [op, key] (op)}<button class="mi" role="menuitem" onclick={() => run((o) => o.order(session, op))}>{t(key)}</button>{/each}
      {/snippet}
    </SplitButton>
    <Button size="mid" icon="arrSelectionPane" tip={t('arr.selectionPane')} on={session.pane.right === 'selection'} onclick={() => session.togglePane('right', 'selection')} />
  </div>
  <div class="col">
    <SplitButton id="arr.align" size="mid" icon="arrAlign" tip={t('arr.align')} disabled={!ops}>
      {#snippet menu()}
        {#each ALIGNS as [op, key] (op)}
          <button class="mi" role="menuitem" onclick={() => run((o) => o.align(session, op, alignTo))}>{t(key)}</button>
          {#if op === 'right' || op === 'bottom'}<hr />{/if}
        {/each}
        <hr />
        <button class="mi check" class:checked={alignTo === 'page'} role="menuitemradio" aria-checked={alignTo === 'page'} onclick={() => (alignTo = 'page')}>{t('arr.alignToPage')}</button>
        <button class="mi check" class:checked={alignTo === 'margin'} role="menuitemradio" aria-checked={alignTo === 'margin'} onclick={() => (alignTo = 'margin')}>{t('arr.alignToMargin')}</button>
      {/snippet}
    </SplitButton>
    <SplitButton id="arr.group" size="mid" icon="arrGroup" tip={t('arr.groupObjects')} disabled={!ops?.group && !ops?.ungroup}>
      {#snippet menu()}
        <button class="mi" role="menuitem" disabled={!ops?.group} onclick={() => run((o) => o.group?.(session))}>{t('arr.groupObjects')}</button>
        <button class="mi" role="menuitem" disabled={!ops?.ungroup} onclick={() => run((o) => o.ungroup?.(session))}>{t('arr.ungroup')}</button>
      {/snippet}
    </SplitButton>
    <SplitButton id="arr.rotate" size="mid" icon="arrRotate" tip={t('arr.rotate')} disabled={!ops}>
      {#snippet menu()}
        {#each ROTATES as [op, key] (op)}<button class="mi" role="menuitem" onclick={() => run((o) => o.rotate(session, op))}>{t(key)}</button>{/each}
        <hr />
        <button class="mi" role="menuitem" disabled={!ops?.layoutOptions} onclick={() => run((o) => o.layoutOptions?.(session))}>{t('arr.moreRotation')}</button>
      {/snippet}
    </SplitButton>
  </div>
</Group>
<LayoutOptionsDialog />

<style>
  .arr-presets {
    grid-template-columns: repeat(3, 36px);
  }
  .arr-preset {
    aspect-ratio: auto;
    height: 40px;
    background: #fff;
  }
</style>
