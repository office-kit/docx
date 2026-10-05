<script lang="ts">
  /**
   * Editing WordArt text. VML WordArt is a text path on a shape, so its text is
   * one string with one font, edited here as in Word's legacy Edit WordArt Text
   * dialog. The same dialog collects the text of new WordArt.
   */
  import { getWordArt } from '@office-kit/docx';
  import { commands } from '@office-kit/docx-editor';
  import Dialog from '../Dialog.svelte';
  import { getSession } from '../session.svelte';
  import { t } from '../i18n/index.svelte';

  const session = getSession();
  const DEFAULT_FONT = 'Calibri';
  const DEFAULT_SIZE = 36;

  let text = $state('');
  let font = $state(DEFAULT_FONT);
  let size = $state(DEFAULT_SIZE);
  let bold = $state(false);
  let italic = $state(false);

  // Load the selected WordArt's text each time the dialog opens.
  $effect(() => {
    if (session.dialog !== 'shape.wordArtText') return;
    const sel = session.selectedObject;
    const model = session.model;
    if (!sel || !model) return;
    try {
      const art = getWordArt(commands.shapeAt(model.doc, sel.at));
      if (!art) return;
      text = art.text;
      font = art.font;
      size = art.size;
      bold = art.bold;
      italic = art.italic;
    } catch {
      // The selection went stale; the dialog edits nothing.
    }
  });

  function ok(): boolean {
    const sel = session.selectedObject;
    if (!sel || !text.trim()) return false;
    session.apply(commands.wordArtTextCommand, { at: sel.at, text: { text, font, size, bold, italic } });
    return true;
  }
</script>

<Dialog id="shape.wordArtText" title={t('draw.editWordArtText')} onok={ok} okDisabled={!text.trim()}>
  <div class="field">
    <label for="wk-wa-font">{t('draw.font')}</label>
    <input id="wk-wa-font" bind:value={font} />
    <label for="wk-wa-size">{t('draw.fontSize')}</label>
    <input id="wk-wa-size" type="number" min="1" max="1638" bind:value={size} />
  </div>
  <div class="field">
    <label><input type="checkbox" bind:checked={bold} /> {t('draw.bold')}</label>
    <label><input type="checkbox" bind:checked={italic} /> {t('draw.italic')}</label>
  </div>
  <textarea class="wk-wordart-text" rows="4" bind:value={text} aria-label={t('draw.text')} style="font-family:{font};font-weight:{bold ? 700 : 400};font-style:{italic ? 'italic' : 'normal'}"></textarea>
</Dialog>
