<script lang="ts">
  /**
   * Insert ▸ Illustrations, in Word's order: Pictures, Shapes, Icons,
   * SmartArt, Chart, Screenshot. The pictures area owns Pictures, Icons,
   * Chart and Screenshot; Shapes and SmartArt live in the marked slots.
   */
  import Button from './Button.svelte';
  import SplitButton from './SplitButton.svelte';
  import ChartTypeDialog from '../picture/ChartTypeDialog.svelte';
  import IconsDialog from '../picture/IconsDialog.svelte';
  import { getSession } from '../session.svelte';
  import { t } from '../i18n/index.svelte';
  import { insertPictureFiles, insertScreenshot } from '../picture/insert';

  const session = getSession();
  let fileInput = $state<HTMLInputElement | null>(null);

  function onFiles(e: Event): void {
    const input = e.currentTarget as HTMLInputElement;
    const files = Array.from(input.files ?? []);
    input.value = '';
    insertPictureFiles(session, files).catch((err: unknown) => {
      session.status = `${t('ill.pictures')}: ${(err as Error).message}`;
    });
  }

  function screenshot(): void {
    session.openMenu = null;
    if (!navigator.mediaDevices?.getDisplayMedia) {
      session.status = t('ill.screenshotUnsupported');
      return;
    }
    insertScreenshot(session).catch((err: unknown) => {
      // The user cancelling the picker (or the browser refusing) is a NotAllowedError.
      session.status = (err as DOMException).name === 'NotAllowedError' ? t('ill.screenshotDenied') : `${t('ill.screenshot')}: ${(err as Error).message}`;
    });
  }
</script>

<!-- pictures area: Pictures -->
<SplitButton id="ill.pictures" size="large" icon="illPictures" tip={t('ill.pictures')}>
  {#snippet menu()}
    <button class="mi" role="menuitem" onclick={() => { session.openMenu = null; fileInput?.click(); }}>{t('ill.pictureFromFile')}</button>
  {/snippet}
</SplitButton>
<input bind:this={fileInput} type="file" accept="image/*" multiple onchange={onFiles} hidden />
<!-- end pictures area -->

<!-- shapes area: Shapes -->
<!-- end shapes area -->

<!-- pictures area: Icons -->
<Button size="large" icon="illIcons" tip={t('ill.icons')} onclick={() => session.openDialog('picture.icons')} />
<!-- end pictures area -->

<!-- shapes area: SmartArt -->
<!-- end shapes area -->

<!-- pictures area: Chart, Screenshot -->
<Button size="large" icon="illChart" tip={t('ill.chart')} onclick={() => session.openDialog('chart.insert')} />
<SplitButton id="ill.screenshot" size="large" icon="illScreenshot" tip={t('ill.screenshot')}>
  {#snippet menu()}
    <button class="mi" role="menuitem" onclick={screenshot}>{t('ill.screenClipping')}</button>
  {/snippet}
</SplitButton>
<IconsDialog />
<ChartTypeDialog mode="insert" />
<!-- end pictures area -->
