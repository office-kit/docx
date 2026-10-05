<script lang="ts">
  /** Layout ▸ Hyphenation ▸ Hyphenation Options…: Word's Hyphenation dialog. */
  import { commands } from '@office-kit/docx-editor';
  import { getDocumentSetting } from '@office-kit/docx';
  import Dialog from '../Dialog.svelte';
  import LengthField from './LengthField.svelte';
  import { getSession } from '../session.svelte';
  import { locale, t } from '../i18n/index.svelte';
  import { lengthUnitFor } from './units';

  const ID = 'layout.hyphenation';
  const session = getSession();
  const unit = $derived(lengthUnitFor(locale()));
  // Word's default hyphenation zone is 0.25 in.
  const DEFAULT_ZONE = 360;

  let open = false;
  let automatic = $state(false);
  let caps = $state(true);
  let zone = $state(DEFAULT_ZONE);
  let limit = $state(0);

  $effect(() => {
    const showing = session.dialog === ID;
    if (showing && !open && session.model) {
      const doc = session.model.doc;
      automatic = getDocumentSetting(doc, 'autoHyphenation').present;
      caps = !getDocumentSetting(doc, 'doNotHyphenateCaps').present;
      zone = Number(getDocumentSetting(doc, 'hyphenationZone').val ?? DEFAULT_ZONE);
      limit = Number(getDocumentSetting(doc, 'consecutiveHyphenLimit').val ?? 0);
    }
    open = showing;
  });

  function ok(): boolean | void {
    if (!Number.isInteger(limit) || limit < 0) return false;
    session.apply(commands.hyphenationCommand, { automatic, hyphenateCaps: caps, zoneTwips: zone, consecutiveLimit: limit });
  }
</script>

<Dialog id={ID} title={t('lay.hy.title')} onok={ok}>
  <label class="check"><input type="checkbox" bind:checked={automatic} />{t('lay.hy.auto')}</label>
  <label class="check"><input type="checkbox" bind:checked={caps} />{t('lay.hy.caps')}</label>
  <div class="form-grid">
    <label for="hy-zone">{t('lay.hy.zone')}</label>
    <LengthField id="hy-zone" value={zone} {unit} label={t('lay.hy.zone')} min={1} onchange={(v) => (zone = v)} />
    <label for="hy-limit">{t('lay.hy.limit')}</label>
    <input id="hy-limit" type="number" min="0" max="32767" placeholder={t('lay.hy.noLimit')} value={limit === 0 ? '' : limit} onchange={(e) => (limit = Number((e.currentTarget as HTMLInputElement).value || 0))} />
  </div>
</Dialog>
