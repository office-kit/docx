<script lang="ts">
  /**
   * Design ▸ Page Color ▸ Fill Effects… ▸ Gradient: a two-colour gradient
   * page background (VML `v:fill type="gradient"`) with Word's four shading
   * styles and four variants.
   */
  import { commands } from '@office-kit/docx-editor';
  import { getPageColor, type PageGradient } from '@office-kit/docx';
  import Dialog from '../Dialog.svelte';
  import { getSession } from '../session.svelte';
  import { t, type MessageKey } from '../i18n/index.svelte';

  const ID = 'design.fillEffects';
  const session = getSession();
  const STYLES: ReadonlyArray<{ value: PageGradient['style']; key: MessageKey; angle: number }> = [
    { value: 'horizontal', key: 'dsn.fill.horizontal', angle: 180 },
    { value: 'vertical', key: 'dsn.fill.vertical', angle: 90 },
    { value: 'diagonalUp', key: 'dsn.fill.diagonalUp', angle: 45 },
    { value: 'diagonalDown', key: 'dsn.fill.diagonalDown', angle: 135 },
  ];
  const VARIANTS = [1, 2, 3, 4] as const;

  let open = false;
  let color1 = $state('FFFFFF');
  let color2 = $state('4472C4');
  let style = $state<PageGradient['style']>('horizontal');
  let variant = $state<PageGradient['variant']>(1);

  $effect(() => {
    const showing = session.dialog === ID;
    if (showing && !open && session.model) {
      const current = getPageColor(session.model.doc);
      color1 = current?.color ?? 'FFFFFF';
      color2 = current?.gradient?.color2 ?? '4472C4';
      style = current?.gradient?.style ?? 'horizontal';
      variant = current?.gradient?.variant ?? 1;
    }
    open = showing;
  });

  /** A CSS preview of a variant: 1 = colour 1 → 2, 2 = reversed, 3 / 4 = mirrored from the middle. */
  function preview(v: PageGradient['variant']): string {
    const angle = STYLES.find((s) => s.value === style)?.angle ?? 180;
    const [a, b] = v === 2 || v === 4 ? [color2, color1] : [color1, color2];
    const stops = v === 3 || v === 4 ? `#${a}, #${b}, #${a}` : `#${a}, #${b}`;
    return `background:linear-gradient(${angle}deg, ${stops})`;
  }

  function ok(): void {
    session.apply(commands.pageColorCommand, { color: { color: color1, gradient: { color2, style, variant } } });
  }
</script>

<Dialog id={ID} title={t('dsn.fill.title')} onok={ok}>
  <div class="dialog-tabs" role="tablist"><span class="dialog-tab on" role="tab" aria-selected="true">{t('dsn.fill.gradient')}</span></div>
  <fieldset>
    <legend>{t('dsn.fill.colors')}</legend>
    <div class="form-grid">
      <label for="fe-c1">{t('dsn.fill.color1')}</label>
      <input id="fe-c1" type="color" value="#{color1}" oninput={(e) => (color1 = (e.currentTarget as HTMLInputElement).value.slice(1).toUpperCase())} />
      <label for="fe-c2">{t('dsn.fill.color2')}</label>
      <input id="fe-c2" type="color" value="#{color2}" oninput={(e) => (color2 = (e.currentTarget as HTMLInputElement).value.slice(1).toUpperCase())} />
    </div>
  </fieldset>
  <div class="row two-col">
    <fieldset>
      <legend>{t('dsn.fill.shadingStyles')}</legend>
      {#each STYLES as s (s.value)}<label class="radio"><input type="radio" bind:group={style} value={s.value} />{t(s.key)}</label>{/each}
    </fieldset>
    <fieldset>
      <legend>{t('dsn.fill.variants')}</legend>
      <div class="variant-grid">
        {#each VARIANTS as v (v)}
          <button type="button" class="variant" class:selected={variant === v} aria-pressed={variant === v} aria-label={`${t('dsn.fill.variants')} ${v}`} style={preview(v)} onclick={() => (variant = v)}></button>
        {/each}
      </div>
    </fieldset>
  </div>
</Dialog>
