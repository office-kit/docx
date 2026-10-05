<script lang="ts">
  /** Drop Cap Options: position (None / Dropped / In margin), font, lines to drop, distance from text. */
  import { getDropCap, type DropCapOptions } from '@office-kit/docx';
  import { commands, paragraphAt } from '@office-kit/docx-editor';
  import Dialog from '../Dialog.svelte';
  import { getSession } from '../session.svelte';
  import { t } from '../i18n/index.svelte';
  import { DIALOG } from './state.svelte';

  const session = getSession();
  const TWIPS_PER_POINT = 20;
  const DEFAULT_LINES = 3;
  const POSITIONS = [
    { value: 'none', key: 'ins.dropCap.none' },
    { value: 'drop', key: 'ins.dropCap.dropped' },
    { value: 'margin', key: 'ins.dropCap.inMargin' },
  ] as const;
  // A few fonts every platform Word runs on ships; any other name can be typed.
  const FONTS = ['Calibri', 'Cambria', 'Times New Roman', 'Arial', 'Georgia', 'Garamond', 'Helvetica'] as const;

  let position = $state<DropCapOptions['position']>('drop');
  let font = $state('');
  let lines = $state(DEFAULT_LINES);
  let distancePt = $state(0);

  $effect(() => {
    if (session.dialog !== DIALOG.dropCap || !session.model) return;
    const focus = session.model.selection?.focus;
    const para = focus ? paragraphAt(session.model.doc, focus) : undefined;
    const current: DropCapOptions = para ? getDropCap(session.model.doc, para) : { position: 'none' };
    position = current.position;
    font = current.font ?? '';
    lines = current.lines ?? DEFAULT_LINES;
    distancePt = (current.distanceTwips ?? 0) / TWIPS_PER_POINT;
  });

  function ok(): boolean {
    session.apply(commands.dropCapCommand, {
      position,
      lines,
      distanceTwips: Math.round(distancePt * TWIPS_PER_POINT),
      ...(font.trim() ? { font: font.trim() } : {}),
    });
    return session.status === '';
  }
</script>

<Dialog id={DIALOG.dropCap} title={t('ins.dropCap')} onok={ok}>
  <fieldset>
    <legend>{t('ins.dropCap.position')}</legend>
    {#each POSITIONS as p (p.value)}
      <label style="margin-right: 12px"><input type="radio" name="dropcap-pos" value={p.value} bind:group={position} /> {t(p.key)}</label>
    {/each}
  </fieldset>
  <label class="stack">{t('ins.dropCap.font')}
    <input type="text" list="wk-dropcap-fonts" bind:value={font} disabled={position === 'none'} />
    <datalist id="wk-dropcap-fonts">{#each FONTS as f (f)}<option value={f}></option>{/each}</datalist>
  </label>
  <div class="dialog-row">
    <label class="stack">{t('ins.dropCap.lines')}<input type="number" min="1" max="10" bind:value={lines} disabled={position === 'none'} /></label>
    <label class="stack">{t('ins.dropCap.distance')}<input type="number" min="0" step="0.5" bind:value={distancePt} disabled={position === 'none'} /> pt</label>
  </div>
  {#if session.status}<p class="error-note">{session.status}</p>{/if}
</Dialog>
