<script lang="ts">
  /** Envelopes and Labels (Mailings ▸ Create, and Start Mail Merge ▸ Envelopes / Labels). */
  import { untrack } from 'svelte';
  import {
    createLabelDocument,
    ENVELOPE_SIZES,
    LABEL_PRODUCTS,
    mailMergeSettings,
    type EnvelopeSize,
  } from '@office-kit/docx';
  import { commands, editorFor } from '@office-kit/docx-editor';
  import Dialog from '../Dialog.svelte';
  import { getSession } from '../session.svelte';
  import { t } from '../i18n/index.svelte';
  import { merge } from './merge.svelte';

  const session = getSession();
  const SIZES = Object.keys(ENVELOPE_SIZES).filter((s): s is EnvelopeSize => Object.hasOwn(ENVELOPE_SIZES, s));

  // Both dialogs start from the selected text, as Word's do.
  $effect(() => {
    const id = session.dialog;
    if (id !== 'mailings.envelopes' && id !== 'mailings.labels') return;
    untrack(() => {
      const selected = session.model ? commands.selectedText(session.model).trim() : '';
      if (!selected) return;
      if (id === 'mailings.envelopes') delivery = selected;
      else labelText = selected;
    });
  });

  // --- Envelopes -----------------------------------------------------------------------
  let delivery = $state('');
  let returnAddress = $state('');
  let omitReturn = $state(false);
  let size = $state<EnvelopeSize>('Size 10');

  function addEnvelope(): boolean {
    session.apply(commands.addEnvelopeCommand, {
      deliveryAddress: delivery,
      size,
      omitReturnAddress: omitReturn,
      ...(returnAddress.trim() && !omitReturn ? { returnAddress } : {}),
    });
    return !session.status;
  }

  // --- Labels -----------------------------------------------------------------------
  let labelText = $state('');
  let productName = $state(LABEL_PRODUCTS[0]?.name ?? '');
  let single = $state(false);
  let row = $state(1);
  let column = $state(1);
  const product = $derived(LABEL_PRODUCTS.find((p) => p.name === productName));
  const forMerge = $derived(session.tick >= 0 && session.model ? mailMergeSettings(session.model.doc)?.type === 'mailingLabels' : false);

  function newLabelDocument(): boolean {
    if (!product) return false;
    try {
      const doc = createLabelDocument({
        text: labelText,
        product,
        ...(single && !forMerge ? { single: { row, column } } : {}),
        ...(forMerge ? { mailMerge: true } : {}),
      });
      const model = editorFor(doc);
      session.load(model, 'Labels1.docx');
      if (forMerge) {
        session.apply(commands.startMailMergeCommand, { type: 'mailingLabels' });
        // The label document replaces the main document; keep its recipients.
        if (merge.list) merge.attach(session, merge.path, merge.list);
      }
      return true;
    } catch (err) {
      // Out-of-range label positions are the user's input: report and keep the dialog open.
      session.status = `${t('mail.labels')}: ${(err as Error).message}`;
      return false;
    }
  }
</script>

<Dialog id="mailings.envelopes" title={t('mail.envelopes')} okLabel={t('mail.addToDocument')} onok={addEnvelope} okDisabled={!delivery.trim()}>
  <label class="field">{t('mail.deliveryAddress')}<textarea bind:value={delivery} rows="4" cols="36"></textarea></label>
  <label class="field">{t('mail.returnAddress')}<textarea bind:value={returnAddress} rows="3" cols="36" disabled={omitReturn}></textarea></label>
  <label><input type="checkbox" bind:checked={omitReturn} /> {t('mail.omit')}</label>
  <label class="field">{t('mail.envelopeSize')}
    <select bind:value={size}>
      {#each SIZES as s (s)}<option value={s}>{s}</option>{/each}
    </select>
  </label>
  <button type="button" class="push" onclick={() => { session.apply(commands.removeEnvelopeCommand, undefined); session.dialog = null; }}>{t('mail.removeEnvelope')}</button>
</Dialog>

<Dialog id="mailings.labels" title={t('mail.labels')} okLabel={t('mail.newDocument')} onok={newLabelDocument}>
  {#if !forMerge}
    <label class="field">{t('mail.address')}<textarea bind:value={labelText} rows="4" cols="36"></textarea></label>
  {/if}
  <label class="field">{t('mail.product')}
    <select bind:value={productName}>
      {#each LABEL_PRODUCTS as p (p.name)}<option value={p.name}>{p.name}</option>{/each}
    </select>
  </label>
  {#if product}<p class="pane-note">{product.across} × {product.down}</p>{/if}
  {#if !forMerge}
    <fieldset>
      <legend>{t('mail.print')}</legend>
      <label><input type="radio" bind:group={single} value={false} /> {t('mail.fullPage')}</label>
      <label><input type="radio" bind:group={single} value={true} /> {t('mail.singleLabel')}</label>
      <label class="field">{t('mail.row')} <input type="number" min="1" max={product?.down} bind:value={row} disabled={!single} /></label>
      <label class="field">{t('mail.column')} <input type="number" min="1" max={product?.across} bind:value={column} disabled={!single} /></label>
    </fieldset>
  {/if}
</Dialog>
