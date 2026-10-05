<script lang="ts">
  /**
   * Compress Pictures: re-encode the selected picture (or every picture) at a
   * target resolution through a canvas, optionally deleting cropped areas.
   */
  import { imageDrawings, readDrawing } from '@office-kit/docx';
  import Dialog from '../Dialog.svelte';
  import { getSession } from '../session.svelte';
  import { t, type MessageKey } from '../i18n/index.svelte';
  import { compressPictures } from './insert';
  import { selectedIndex } from './state';

  const session = getSession();
  // Word's resolution choices (ppi); `undefined` keeps the original pixels.
  const RESOLUTIONS: ReadonlyArray<readonly [number | undefined, MessageKey]> = [
    [330, 'compress.hd'],
    [220, 'compress.print'],
    [150, 'compress.web'],
    [96, 'compress.email'],
    [undefined, 'compress.original'],
  ];
  let onlySelected = $state(true);
  let deleteCropped = $state(true);
  let ppi = $state<number | undefined>(220);

  function ok(): void {
    const model = session.model;
    if (!model) return;
    const doc = model.doc;
    const indices = onlySelected
      ? [selectedIndex(session)].filter((i) => i >= 0)
      : imageDrawings(doc).flatMap((d, i) => (readDrawing(doc, d).kind === 'picture' ? [i] : []));
    compressPictures(session, indices, ppi, deleteCropped).catch((err: unknown) => {
      session.status = `${t('pic.compress')}: ${(err as Error).message}`;
    });
  }
</script>

<Dialog id="picture.compress" title={t('pic.compress')} onok={ok}>
  <fieldset>
    <legend>{t('compress.options')}</legend>
    <label class="field"><input type="checkbox" bind:checked={onlySelected} />{t('compress.onlySelected')}</label>
    <label class="field"><input type="checkbox" bind:checked={deleteCropped} />{t('compress.deleteCropped')}</label>
  </fieldset>
  <fieldset>
    <legend>{t('compress.resolution')}</legend>
    {#each RESOLUTIONS as [value, key] (key)}
      <label class="field"><input type="radio" bind:group={ppi} {value} />{t(key)}</label>
    {/each}
  </fieldset>
</Dialog>
