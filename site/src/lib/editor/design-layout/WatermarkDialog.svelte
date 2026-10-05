<script lang="ts">
  /**
   * Design ▸ Watermark ▸ Custom Watermark…: no watermark, a picture
   * watermark (scale, washout) or a text watermark (language, text, font,
   * size, colour, semitransparent, diagonal / horizontal), as in Word.
   */
  import { commands } from '@office-kit/docx-editor';
  import { getSectionProperties, getWatermark, type Watermark } from '@office-kit/docx';
  import Dialog from '../Dialog.svelte';
  import { getSession } from '../session.svelte';
  import { LOCALES, locale, t, type LocaleId } from '../i18n/index.svelte';
  import designMessages from '../i18n/messages/design';

  const ID = 'design.watermark';
  const session = getSession();
  const PRESET_KEYS = ['dsn.wm.asap', 'dsn.wm.confidential', 'dsn.wm.copy', 'dsn.wm.doNotCopy', 'dsn.wm.draft', 'dsn.wm.original', 'dsn.wm.personal', 'dsn.wm.sample', 'dsn.wm.topSecret', 'dsn.wm.urgent'] as const;
  const SIZES = [36, 40, 44, 48, 54, 60, 66, 72, 80, 90, 96, 105, 120, 144];
  const SCALES = [500, 200, 150, 100, 50];
  const FONTS = ['Calibri', 'Calibri Light', 'Arial', 'Times New Roman', 'Aptos', 'Georgia', 'Verdana', 'Yu Gothic', 'Yu Mincho'];
  // CSS pixels are 1/96 in; points are 1/72 in.
  const POINTS_PER_PIXEL = 0.75;
  const TWIPS_PER_POINT = 20;

  let open = false;
  let kind = $state<'none' | 'picture' | 'text'>('none');
  let language = $state<LocaleId>('en');
  let text = $state('');
  let font = $state('Calibri');
  let size = $state<'auto' | number>('auto');
  let color = $state('C0C0C0');
  let semitransparent = $state(true);
  let layout = $state<'diagonal' | 'horizontal'>('diagonal');
  let picture = $state<{ bytes: Uint8Array; width: number; height: number; name: string } | null>(null);
  let scale = $state<'auto' | number>('auto');
  let washout = $state(true);
  let error = $state('');

  $effect(() => {
    const showing = session.dialog === ID;
    if (showing && !open && session.model) {
      const current = getWatermark(session.model.doc);
      language = locale();
      error = '';
      picture = null;
      if (current?.kind === 'text') {
        kind = 'text';
        ({ text, font, size, color, semitransparent, layout } = current);
      } else {
        kind = current ? 'picture' : 'none';
        text = t('dsn.wm.asap');
        size = 'auto';
        color = 'C0C0C0';
        semitransparent = true;
        layout = 'diagonal';
      }
    }
    open = showing;
  });

  async function choosePicture(e: Event): Promise<void> {
    const file = (e.currentTarget as HTMLInputElement).files?.[0];
    if (!file) return;
    try {
      const bitmap = await createImageBitmap(file);
      picture = { bytes: new Uint8Array(await file.arrayBuffer()), width: bitmap.width, height: bitmap.height, name: file.name };
      bitmap.close();
      kind = 'picture';
      error = '';
    } catch {
      error = t('dsn.wm.badPicture');
    }
  }

  function pictureSize(): { widthPt: number; heightPt: number } | undefined {
    if (!picture || !session.model) return undefined;
    const w = picture.width * POINTS_PER_PIXEL;
    const h = picture.height * POINTS_PER_PIXEL;
    if (scale !== 'auto') return { widthPt: (w * scale) / 100, heightPt: (h * scale) / 100 };
    const { pageSize, margins } = getSectionProperties(session.model.doc);
    const textWidth = (pageSize.widthTwips - margins.left - margins.right) / TWIPS_PER_POINT;
    return { widthPt: textWidth, heightPt: (textWidth * h) / w };
  }

  function ok(): boolean | void {
    let watermark: Watermark | undefined;
    if (kind === 'text') {
      if (!text.trim()) return false;
      watermark = { kind: 'text', text, font, size, color, semitransparent, layout };
    } else if (kind === 'picture') {
      const dims = pictureSize();
      // Keeping an existing picture watermark needs no new picture.
      if (!picture || !dims) return picture ? false : undefined;
      watermark = { kind: 'picture', bytes: picture.bytes, ...dims, washout };
    }
    session.apply(commands.watermarkCommand, { watermark });
  }
</script>

<Dialog id={ID} title={t('dsn.wm.title')} onok={ok}>
  <label class="radio"><input type="radio" bind:group={kind} value="none" />{t('dsn.wm.none')}</label>
  <label class="radio"><input type="radio" bind:group={kind} value="picture" />{t('dsn.wm.picture')}</label>
  <div class="indent form-grid" class:dim={kind !== 'picture'}>
    <span></span>
    <label class="push file-push">{t('dsn.wm.selectPicture')}<input type="file" accept="image/png,image/jpeg,image/gif,image/bmp" hidden onchange={choosePicture} /></label>
    {#if picture}<span></span><span class="muted">{picture.name}</span>{/if}
    <label for="wm-scale">{t('dsn.wm.scale')}</label>
    <select id="wm-scale" bind:value={scale} disabled={kind !== 'picture'}>
      <option value="auto">{t('dsn.auto')}</option>
      {#each SCALES as s (s)}<option value={s}>{s}%</option>{/each}
    </select>
    <span></span>
    <label class="check"><input type="checkbox" bind:checked={washout} disabled={kind !== 'picture'} />{t('dsn.wm.washout')}</label>
  </div>
  {#if error}<p class="error">{error}</p>{/if}
  <label class="radio"><input type="radio" bind:group={kind} value="text" />{t('dsn.wm.text')}</label>
  <div class="indent form-grid" class:dim={kind !== 'text'}>
    <label for="wm-lang">{t('dsn.wm.language')}</label>
    <select id="wm-lang" bind:value={language} disabled={kind !== 'text'}>
      {#each Object.entries(LOCALES) as [id, name] (id)}<option value={id}>{name}</option>{/each}
    </select>
    <label for="wm-text">{t('dsn.wm.textLabel')}</label>
    <input id="wm-text" type="text" list="wm-texts" bind:value={text} disabled={kind !== 'text'} />
    <label for="wm-font">{t('dsn.wm.font')}</label>
    <input id="wm-font" type="text" list="wm-fonts" bind:value={font} disabled={kind !== 'text'} />
    <label for="wm-size">{t('dsn.wm.size')}</label>
    <select id="wm-size" bind:value={size} disabled={kind !== 'text'}>
      <option value="auto">{t('dsn.auto')}</option>
      {#each SIZES as s (s)}<option value={s}>{s}</option>{/each}
    </select>
    <label for="wm-color">{t('dsn.wm.color')}</label>
    <span class="row">
      <input id="wm-color" type="color" value="#{color}" disabled={kind !== 'text'} oninput={(e) => (color = (e.currentTarget as HTMLInputElement).value.slice(1).toUpperCase())} />
      <label class="check"><input type="checkbox" bind:checked={semitransparent} disabled={kind !== 'text'} />{t('dsn.wm.semitransparent')}</label>
    </span>
    <span>{t('dsn.wm.layout')}</span>
    <span class="row">
      <label class="radio"><input type="radio" bind:group={layout} value="diagonal" disabled={kind !== 'text'} />{t('dsn.wm.diagonal')}</label>
      <label class="radio"><input type="radio" bind:group={layout} value="horizontal" disabled={kind !== 'text'} />{t('dsn.wm.horizontal')}</label>
    </span>
  </div>
  <datalist id="wm-texts">{#each PRESET_KEYS as key (key)}<option value={designMessages[language][key]}></option>{/each}</datalist>
  <datalist id="wm-fonts">{#each FONTS as f (f)}<option value={f}></option>{/each}</datalist>
</Dialog>
