<script lang="ts">
  /**
   * Insert ▸ Header / Footer: the built-in gallery (Blank, Blank (Three
   * Columns) and the styled presets), Edit and Remove. Shared by
   * HeaderButton / FooterButton, which the Header & Footer tab also embeds.
   */
  import { commands } from '@office-kit/docx-editor';
  import SplitButton from '../ribbon/SplitButton.svelte';
  import { getSession } from '../session.svelte';
  import { t } from '../i18n/index.svelte';

  type Props = { kind: 'header' | 'footer' };
  const { kind }: Props = $props();
  const session = getSession();
  const presets = commands.HEADER_FOOTER_PRESETS;

  function apply(preset: (typeof presets)[number]): void {
    const cmd = kind === 'header' ? commands.insertHeaderPresetCommand : commands.insertFooterPresetCommand;
    session.apply(cmd, { preset });
  }

  function edit(): void {
    session.openMenu = null;
    session.editHeaderFooter(kind);
  }

  function remove(): void {
    session.apply(kind === 'header' ? commands.removeHeaderCommand : commands.removeFooterCommand, {});
  }
</script>

<SplitButton id="insert.{kind}" size="large" icon={kind} tip={t(kind === 'header' ? 'ins.header' : 'ins.footer')}>
  {#snippet menu()}
    <div class="menu-head">{t('ins.coverPage.builtIn')}</div>
    <div class="gallery-menu wide">
      {#each presets as preset (preset)}
        <button class="tile" role="menuitem" onclick={() => apply(preset)}>
          <span class="preview-page" aria-hidden="true">
            <span class="preview-mark preset-{preset}" style={kind === 'header' ? 'top: 8%' : 'bottom: 8%'}></span>
          </span>
          <span class="tile-name">{t(`ins.preset.${preset}`)}</span>
        </button>
      {/each}
    </div>
    <hr />
    <button class="mi" role="menuitem" onclick={edit}>{t(kind === 'header' ? 'ins.header.edit' : 'ins.footer.edit')}</button>
    <button class="mi" role="menuitem" onclick={remove}>{t(kind === 'header' ? 'ins.header.remove' : 'ins.footer.remove')}</button>
  {/snippet}
</SplitButton>
