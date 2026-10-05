<script lang="ts">
  /**
   * Word's Design tab: Themes, the Document Formatting gallery (style sets),
   * Colors, Fonts, Paragraph Spacing, Effects, and the Page Background group
   * (Watermark, Page Color, Page Borders). Set as Default is shown but
   * disabled: it writes to the user's Normal template, not to the document.
   */
  import {
    commands,
    resolveTheme,
  } from '@office-kit/docx-editor';
  import {
    currentStyleSet,
    getDefaultParagraphSpacing,
    getPageColor,
    getWatermark,
    PARAGRAPH_SPACING_PRESETS,
    STYLE_SETS,
    THEME_COLOR_SCHEMES,
    THEME_EFFECT_SCHEMES,
    THEME_FONT_SCHEMES,
    THEMES,
    themeColorValue,
    type StyleSetDefinition,
    type StyleSetStyle,
    type ColorValue,
    type ThemeColorScheme,
  } from '@office-kit/docx';
  import RibbonIcon from '../../RibbonIcon.svelte';
  import Button from '../Button.svelte';
  import SplitButton from '../SplitButton.svelte';
  import Group from '../Group.svelte';
  import { getSession } from '../../session.svelte';
  import { t, type MessageKey } from '../../i18n/index.svelte';
  import ColorMenu from '../../ColorMenu.svelte';
  import { WATERMARK_PRESETS, type WatermarkPreset } from '../../design-layout/presets';
  import CustomColorsDialog from '../../design-layout/CustomColorsDialog.svelte';
  import CustomFontsDialog from '../../design-layout/CustomFontsDialog.svelte';
  import ParagraphSpacingDialog from '../../design-layout/ParagraphSpacingDialog.svelte';
  import WatermarkDialog from '../../design-layout/WatermarkDialog.svelte';
  import FillEffectsDialog from '../../design-layout/FillEffectsDialog.svelte';
  import PageBordersDialog from '../../design-layout/PageBordersDialog.svelte';
  import '../../design-layout/design-layout.css';

  const session = getSession();
  // A fresh wrapper per edit: the document object never changes identity, so
  // deriving from it directly would never recompute after an edit.
  const live = $derived(session.tick >= 0 && session.model ? { doc: session.model.doc } : undefined);
  const theme = $derived(live ? resolveTheme(live.doc) : undefined);
  const styleSet = $derived(live ? currentStyleSet(live.doc) : undefined);
  const spacing = $derived(live ? getDefaultParagraphSpacing(live.doc) : undefined);
  const watermark = $derived(live ? getWatermark(live.doc) : undefined);
  const pageColor = $derived(live ? getPageColor(live.doc) : undefined);

  // Word's Paragraph Spacing names, in menu order.
  const SPACING_KEYS: Readonly<Record<string, MessageKey>> = {
    'No Paragraph Space': 'dsn.spacing.none',
    Compact: 'dsn.spacing.compact',
    Tight: 'dsn.spacing.tight',
    Open: 'dsn.spacing.open',
    Relaxed: 'dsn.spacing.relaxed',
    Double: 'dsn.spacing.double',
  };
  // Word's default watermark: Calibri, auto size, silver, semitransparent.
  const WATERMARK_FONT = 'Calibri';
  const WATERMARK_COLOR = 'C0C0C0';
  const TWIPS_PER_POINT = 20;

  function hex(c: string | undefined): string {
    return c && /^[0-9A-Fa-f]{6}$/.test(c) ? `#${c}` : 'transparent';
  }

  /** A style-set style as inline CSS for a gallery preview, in the current theme. */
  function previewCss(s: StyleSetStyle | undefined, colors: ThemeColorScheme, scale: number): string {
    if (!s) return '';
    const font = s.font === 'major' ? theme?.fonts.major : theme?.fonts.minor;
    const color = s.color ? (s.color.theme ? themeColorValue(colors, s.color.theme, { shade: s.color.shade, tint: s.color.tint }) : s.color.rgb) : undefined;
    const fill = s.shading ? (s.shading.theme ? themeColorValue(colors, s.shading.theme, { shade: s.shading.shade, tint: s.shading.tint }) : s.shading.rgb) : undefined;
    const rule = s.borderBottom ? (s.borderBottom.color.theme ? themeColorValue(colors, s.borderBottom.color.theme) : s.borderBottom.color.rgb) : undefined;
    return [
      font ? `font-family:'${font}',sans-serif` : '',
      s.sizeHalfPoints ? `font-size:${Math.max(4, (s.sizeHalfPoints / 2) * scale)}px` : '',
      s.bold ? 'font-weight:700' : '',
      s.italic ? 'font-style:italic' : '',
      s.caps ? 'text-transform:uppercase' : '',
      color ? `color:${hex(color)}` : '',
      fill ? `background:${hex(fill)}` : '',
      rule ? `border-bottom:1px solid ${hex(rule)}` : '',
      s.alignment ? `text-align:${s.alignment}` : '',
    ]
      .filter(Boolean)
      .join(';');
  }

  function applyWatermark(p: WatermarkPreset): void {
    session.apply(commands.watermarkCommand, {
      watermark: { kind: 'text', text: t(p.key), font: WATERMARK_FONT, size: 'auto', color: WATERMARK_COLOR, semitransparent: true, layout: p.layout },
    });
  }

  function isCurrentWatermark(p: WatermarkPreset): boolean {
    return watermark?.kind === 'text' && watermark.text === t(p.key) && watermark.layout === p.layout;
  }

  // `w:background` stores tint / shade as two hex digits.
  const hexByte = (n: number): string => n.toString(16).toUpperCase().padStart(2, '0');

  function applyPageColor(c: ColorValue): void {
    session.apply(commands.pageColorCommand, {
      color: {
        color: c.rgb,
        ...(c.themeColor ? { themeColor: c.themeColor } : {}),
        ...(c.themeShade === undefined ? {} : { themeShade: hexByte(c.themeShade) }),
        ...(c.themeTint === undefined ? {} : { themeTint: hexByte(c.themeTint) }),
      },
    });
  }

  function sameColors(a: ThemeColorScheme, b: ThemeColorScheme | undefined): boolean {
    return !!b && (['dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'] as const).every((k) => a[k] === b[k]);
  }

  const swatchSlots = ['dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'] as const;
</script>

{#snippet setTile(set: StyleSetDefinition, colors: ThemeColorScheme)}
  <button
    class="tile set-tile"
    class:selected={styleSet?.name === set.name}
    role="option"
    aria-selected={styleSet?.name === set.name}
    title={set.name}
    onclick={() => session.apply(commands.applyStyleSetCommand, { set })}
  >
    <span class="set-title" style={previewCss(set.styles.Title, colors, 0.32)}>{t('dsn.sample.title')}</span>
    <span class="set-heading" style={previewCss(set.styles.Heading1, colors, 0.42)}>{t('dsn.sample.heading1')}</span>
    <span class="set-body">{t('dsn.sample.body')}</span>
  </button>
{/snippet}

<Group label={t('dsn.group.themes')}>
  <SplitButton id="design.themes" size="large" tip={t('dsn.themes')} disabled={!theme}>
    {#snippet face()}<RibbonIcon name="themes" size={32} />{/snippet}
    {#snippet menu()}
      <div class="menu-head">{t('dsn.themes.office')}</div>
      <div class="theme-grid">
        {#each THEMES as th (th.name)}
          <button class="theme-tile" class:selected={theme?.name === th.name} role="menuitemradio" aria-checked={theme?.name === th.name} title={th.name} onclick={() => session.apply(commands.applyThemeCommand, { theme: th })}>
            <span class="theme-aa" style="font-family:'{th.fonts.major}',sans-serif">Aa</span>
            <span class="theme-bars">{#each ['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'] as const as k (k)}<i style="background:#{th.colors[k]}"></i>{/each}</span>
            <span class="theme-name">{th.name}</span>
          </button>
        {/each}
      </div>
      <!-- Word for Mac has no Effects button on the ribbon; the theme's effect schemes live with the themes. -->
      <div class="menu-head">{t('dsn.effects')}</div>
      {#each THEME_EFFECT_SCHEMES as scheme (scheme.name)}
        <button class="mi check" class:checked={theme?.effects === scheme.name} role="menuitemradio" aria-checked={theme?.effects === scheme.name} onclick={() => session.apply(commands.themeEffectsCommand, { effects: scheme })}>{scheme.name}</button>
      {/each}
      <hr />
      <button class="mi" role="menuitem" onclick={() => session.apply(commands.resetThemeCommand, undefined)}>{t('dsn.themes.reset')}</button>
    {/snippet}
  </SplitButton>
</Group>

<Group label={t('dsn.group.documentFormatting')} class="gallery-group">
  {#if theme}
    <div class="gallery set-gallery" role="listbox" aria-label={t('dsn.group.documentFormatting')}>
      {#each STYLE_SETS as set (set.name)}{@render setTile(set, theme.colors)}{/each}
    </div>
    <SplitButton id="design.styleSets" tip={t('dsn.styleSet.more')} alignRight>
      {#snippet face()}<RibbonIcon name="chevronDown" size={12} />{/snippet}
      {#snippet menu()}
        <div class="menu-head">{t('dsn.styleSet.builtIn')}</div>
        <div class="set-grid">{#each STYLE_SETS as set (set.name)}{@render setTile(set, theme.colors)}{/each}</div>
        <hr />
        <button class="mi" role="menuitem" onclick={() => session.apply(commands.resetStyleSetCommand, undefined)}>{t('dsn.styleSet.reset')}</button>
      {/snippet}
    </SplitButton>
  {/if}
  <SplitButton id="design.colors" size="large" tip={t('dsn.colors')} disabled={!theme}>
    {#snippet face()}<RibbonIcon name="themeColors" size={32} />{/snippet}
    {#snippet menu()}
      <div class="menu-head">{t('dsn.themes.office')}</div>
      <div class="scheme-list">
        {#each THEME_COLOR_SCHEMES as scheme (scheme.name)}
          <button class="mi check" class:checked={sameColors(scheme, theme?.colors)} role="menuitemradio" aria-checked={sameColors(scheme, theme?.colors)} onclick={() => session.apply(commands.themeColorsCommand, { colors: scheme })}>
            <span class="swatches">{#each swatchSlots as k (k)}<i style="background:#{scheme[k]}"></i>{/each}</span>
            {scheme.name}
          </button>
        {/each}
      </div>
      <hr />
      <button class="mi" role="menuitem" onclick={() => session.openDialog('design.customColors')}>{t('dsn.colors.customize')}</button>
    {/snippet}
  </SplitButton>
  <SplitButton id="design.fonts" size="large" tip={t('dsn.fonts')} disabled={!theme}>
    {#snippet face()}<RibbonIcon name="themeFonts" size={32} />{/snippet}
    {#snippet menu()}
      <div class="menu-head">{t('dsn.themes.office')}</div>
      <div class="scheme-list">
        {#each THEME_FONT_SCHEMES as scheme (scheme.name)}
          {@const on = theme?.fonts.major === scheme.major && theme?.fonts.minor === scheme.minor}
          <button class="mi check font-scheme" class:checked={on} role="menuitemradio" aria-checked={on} onclick={() => session.apply(commands.themeFontsCommand, { fonts: scheme })}>
            <span class="font-aa" style="font-family:'{scheme.major}',sans-serif">Aa</span>
            <span class="font-names">
              <b>{scheme.name}</b>
              <span style="font-family:'{scheme.major}',sans-serif">{scheme.major}</span>
              <span style="font-family:'{scheme.minor}',sans-serif">{scheme.minor}</span>
            </span>
          </button>
        {/each}
      </div>
      <hr />
      <button class="mi" role="menuitem" onclick={() => session.openDialog('design.customFonts')}>{t('dsn.fonts.customize')}</button>
    {/snippet}
  </SplitButton>
  <div class="col">
    <SplitButton id="design.spacing" size="mid" icon="paragraphSpacing" tip={t('dsn.spacing')}>
      {#snippet menu()}
        <div class="menu-head">{t('dsn.spacing.builtIn')}</div>
        <button class="mi check" class:checked={!!styleSet && spacing?.after === styleSet.after && spacing?.line === styleSet.line} role="menuitemradio" aria-checked="false" onclick={() => session.apply(commands.paragraphSpacingCommand, { spacing: 'default' })}>{t('dsn.spacing.default')}</button>
        {#each PARAGRAPH_SPACING_PRESETS as preset (preset.name)}
          {@const on = spacing?.after === preset.after && spacing?.line === preset.line && spacing?.before === preset.before}
          <button class="mi check" class:checked={on} role="menuitemradio" aria-checked={on} title={`${t('dsn.spacing.after')}: ${preset.after / TWIPS_PER_POINT} pt · ${t('dsn.spacing.line')}: ${preset.line / 240}`} onclick={() => session.apply(commands.paragraphSpacingCommand, { spacing: preset })}>
            {t(SPACING_KEYS[preset.name] ?? 'dsn.spacing.default')}
          </button>
        {/each}
        <hr />
        <button class="mi" role="menuitem" onclick={() => session.openDialog('design.spacing')}>{t('dsn.spacing.custom')}</button>
      {/snippet}
    </SplitButton>
    <Button size="mid" icon="setDefault" tip={t('dsn.setDefault.tip')} label={t('dsn.setDefault')} onclick={() => undefined} disabled />
  </div>
</Group>

<Group label={t('dsn.group.pageBackground')}>
  <SplitButton id="design.watermark" size="large" tip={t('dsn.watermark')} alignRight>
    {#snippet face()}<RibbonIcon name="watermark" size={32} />{/snippet}
    {#snippet menu()}
      {#each WATERMARK_PRESETS as group (group.group)}
        <div class="menu-head">{t(group.group)}</div>
        <div class="wm-grid">
          {#each group.items as p (`${p.key}-${p.layout}`)}
            <button class="wm-tile" class:selected={isCurrentWatermark(p)} role="menuitem" title={`${t(p.key)} ${p.layout === 'diagonal' ? 1 : 2}`} onclick={() => applyWatermark(p)}>
              <span class="wm-text" class:diag={p.layout === 'diagonal'}>{t(p.key)}</span>
            </button>
          {/each}
        </div>
      {/each}
      <hr />
      <button class="mi" role="menuitem" onclick={() => session.openDialog('design.watermark')}>{t('dsn.watermark.custom')}</button>
      <button class="mi" role="menuitem" disabled={!watermark} onclick={() => session.apply(commands.watermarkCommand, { watermark: undefined })}>{t('dsn.watermark.remove')}</button>
    {/snippet}
  </SplitButton>
  <SplitButton id="design.pageColor" size="large" tip={t('dsn.pageColor')} alignRight>
    {#snippet face()}<span class="page-color-face"><RibbonIcon name="pageColor" size={32} /><i style="background:{pageColor ? hex(pageColor.color) : 'transparent'}"></i></span>{/snippet}
    {#snippet menu()}
      <ColorMenu onpick={applyPageColor} none={{ label: t('dsn.pageColor.none'), onpick: () => session.apply(commands.pageColorCommand, { color: undefined }) }}>
        <button class="mi" role="menuitem" onclick={() => session.openDialog('design.fillEffects')}>{t('dsn.pageColor.fill')}</button>
      </ColorMenu>
    {/snippet}
  </SplitButton>
  <Button size="large" icon="pageBorders" tip={t('dsn.pageBorders')} onclick={() => session.openDialog('design.pageBorders')} />
</Group>

<CustomColorsDialog />
<CustomFontsDialog />
<ParagraphSpacingDialog />
<WatermarkDialog />
<FillEffectsDialog />
<PageBordersDialog />
