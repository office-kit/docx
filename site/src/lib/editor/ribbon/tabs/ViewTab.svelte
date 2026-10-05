<script lang="ts">
  /**
   * Word for Mac's View tab: Views, Focus, Show, Zoom and document
   * Properties. Immersive Reader, Window (New Window, Arrange All, Split,
   * Switch Windows) and Macros are left out: the page edits one document in
   * one browser tab, and VBA cannot run in a browser.
   */
  import RibbonIcon from '../../RibbonIcon.svelte';
  import Button from '../Button.svelte';
  import Group from '../Group.svelte';
  import { getSession, type ViewMode } from '../../session.svelte';
  import { t, locale, setLocale, LOCALES, type LocaleId, type MessageKey } from '../../i18n/index.svelte';
  import type { IconName } from '../../RibbonIcon.svelte';
  import { fitZoom, type ZoomFit } from '../../review-actions';

  const session = getSession();

  const VIEWS: readonly [ViewMode, IconName, MessageKey][] = [
    ['print', 'viewPrintLayout', 'view.printLayout'],
    ['web', 'viewWebLayout', 'view.webLayout'],
    ['outline', 'viewOutline', 'view.outline'],
    ['draft', 'viewDraft', 'view.draft'],
  ];

  function fit(kind: ZoomFit, pagesAcross: number): void {
    const zoom = fitZoom(session, kind);
    if (zoom === undefined) return;
    session.prefs.pagesAcross = pagesAcross;
    session.setZoom(zoom);
  }
</script>

<Group label={t('group.views')}>
  {#each VIEWS as [mode, icon, key] (mode)}
    <Button size="large" {icon} tip={t(key)} on={session.viewMode === mode} onclick={() => session.setView(mode)} />
  {/each}
</Group>
<Group label={t('group.immersive')}>
  <Button size="large" icon="viewFocus" tip={t('view.focus')} on={session.prefs.focus} onclick={() => (session.prefs.focus = true)} />
</Group>
<Group label={t('group.show')}>
  <label class="check-item"><input type="checkbox" bind:checked={session.showRuler} />{t('view.ruler')}</label>
  <label class="check-item"><input type="checkbox" bind:checked={session.showGridlines} disabled={session.viewMode !== 'print'} />{t('view.gridlines')}</label>
  <label class="check-item"><input type="checkbox" checked={session.pane.left === 'navigation'} onchange={() => session.togglePane('left', 'navigation')} />{t('view.navigation')}</label>
  <label class="check-item"><input type="checkbox" checked={session.pane.right === 'xml'} onchange={() => session.togglePane('right', 'xml')} />{t('view.xml')}</label>
</Group>
<Group label={t('group.zoom')}>
  <Button size="large" icon="viewZoom" tip={t('view.zoom')} onclick={() => session.openDialog('zoom')} />
  <Button size="large" icon="viewZoom100" tip={t('view.zoom100')} onclick={() => { session.prefs.pagesAcross = 1; session.setZoom(1); }} />
  <div class="rows">
    <Button size="mid" icon="viewOnePage" tip={t('view.onePage')} onclick={() => fit('wholePage', 1)} />
    <Button size="mid" icon="viewMultiplePages" tip={t('view.multiplePages')} onclick={() => fit('multiplePages', 2)} />
    <Button size="mid" icon="viewPageWidth" tip={t('view.pageWidth')} onclick={() => fit('pageWidth', 1)} />
  </div>
</Group>
<Group label={t('group.properties')}>
  <Button size="large" icon="viewProperties" tip={t('view.properties')} onclick={() => session.openDialog('properties')} />
</Group>
<Group label={t('group.language')}>
  <label class="big lang" title={t('language')}>
    <RibbonIcon name="language" size={32} />
    <select value={locale()} onchange={(e) => setLocale((e.currentTarget as HTMLSelectElement).value as LocaleId)} aria-label={t('language')}>
      {#each Object.entries(LOCALES) as [id, name] (id)}
        <option value={id}>{name}</option>
      {/each}
    </select>
  </label>
</Group>
