<script lang="ts">
  import RibbonIcon from '../../RibbonIcon.svelte';
  import Button from '../Button.svelte';
  import Group from '../Group.svelte';
  import { getSession } from '../../session.svelte';
  import { t, locale, setLocale, LOCALES, type LocaleId } from '../../i18n/index.svelte';

  const session = getSession();
</script>

<Group label={t('group.show')}>
  <label class="check-item"><input type="checkbox" checked={session.pane.left === 'navigation'} onchange={() => session.togglePane('left', 'navigation')} />{t('view.navigation')}</label>
  <label class="check-item"><input type="checkbox" checked={session.pane.right === 'xml'} onchange={() => session.togglePane('right', 'xml')} />{t('view.xml')}</label>
</Group>
<Group label={t('group.zoom')}>
  <Button size="large" icon="zoomIn" tip={t('view.zoomIn')} onclick={() => session.setZoom(session.zoom + 0.1)} />
  <Button size="large" icon="zoomOut" tip={t('view.zoomOut')} onclick={() => session.setZoom(session.zoom - 0.1)} />
  <Button size="large" icon="onePage" tip={t('view.zoom100')} onclick={() => session.setZoom(1)} />
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
