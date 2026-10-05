<script lang="ts">
  /**
   * Word's Styles pane, docked right: the current style, New Style, Clear
   * All, and every style (previewed or plain) to apply, modify or delete.
   */
  import { commands } from '@office-kit/docx-editor';
  import RibbonIcon from '../RibbonIcon.svelte';
  import { getSession } from '../session.svelte';
  import { t } from '../i18n/index.svelte';
  import { common, homeState, selectionFormats } from '../ribbon/tabs/home/state.svelte';
  import { stylePreviews } from '../ribbon/tabs/home/style-preview';

  const session = getSession();
  const home = homeState(session);

  let showPreview = $state(true);
  let list = $state<'recommended' | 'inUse'>('recommended');
  let menuFor = $state<string | null>(null);

  const model = $derived(session.tick >= 0 ? session.model : null);
  const styles = $derived(session.version >= 0 && session.model ? commands.listStyles(session.model) : []);
  const shown = $derived(list === 'inUse' ? styles.filter((s) => s.inDocument) : styles);
  const previews = $derived(session.version >= 0 && session.model ? stylePreviews(session.model) : undefined);
  const current = $derived.by(() => {
    if (!model) return undefined;
    const ids = selectionFormats(model).paragraphStyles;
    if (ids.length === 0) return undefined;
    return common(ids) ?? (ids.every((x) => x === undefined) ? commands.defaultParagraphStyleId(model) : undefined);
  });
  const currentName = $derived(styles.find((s) => s.styleId === current)?.name ?? '');

  function modify(styleId: string): void {
    menuFor = null;
    home.styleTarget = styleId;
    session.openDialog('home.style');
  }

  function remove(styleId: string): void {
    menuFor = null;
    session.apply(commands.deleteStyleCommand, { styleId });
  }
</script>

<div class="pane-head">
  <span>{t('home.styles.pane')}</span>
  <button class="pane-close" onclick={() => (session.pane.right = null)} aria-label={t('find.close')}><RibbonIcon name="close" size={14} /></button>
</div>
<div class="pane-note">{t('home.styles.current')}: <strong>{currentName}</strong></div>
<div class="pane-buttons">
  <button class="pane-button" onclick={() => { home.styleTarget = null; session.openDialog('home.style'); }}>{t('home.styles.new')}</button>
  <button class="pane-button" onclick={() => session.apply(commands.clearFormattingCommand, undefined)} disabled={!session.enabled(commands.clearFormattingCommand)}>{t('home.styles.clearAll')}</button>
</div>
<ul class="pane-list styles-list" role="listbox" aria-label={t('group.styles')}>
  {#each shown as s (s.styleId)}
    <li class:current={s.styleId === current}>
      <button class="style-apply" role="option" aria-selected={s.styleId === current} onclick={() => session.apply(commands.applyStyleCommand, { styleId: s.styleId })} disabled={!session.enabled(commands.applyStyleCommand)}>
        <span class="style-name" style={showPreview ? previews?.css(s) : ''}>{s.name}</span>
        <span class="style-kind" aria-hidden="true">{s.type === 'paragraph' ? '¶' : 'a'}</span>
      </button>
      <button class="pane-icon" onclick={() => (menuFor = menuFor === s.styleId ? null : s.styleId)} aria-label={`${s.name} ${t('home.find.options')}`} aria-expanded={menuFor === s.styleId}><RibbonIcon name="chevronDown" size={10} /></button>
      {#if menuFor === s.styleId}
        <div class="style-menu" role="menu">
          <button class="mi" role="menuitem" onclick={() => modify(s.styleId)}>{t('home.styles.modify')}</button>
          {#if s.inDocument && model && commands.isDeletableStyle(model, s.styleId)}
            <button class="mi" role="menuitem" onclick={() => remove(s.styleId)}>{t('home.styles.delete').replace('{name}', s.name)}</button>
          {/if}
        </div>
      {/if}
    </li>
  {/each}
</ul>
<label class="field"><input type="checkbox" bind:checked={showPreview} />{t('home.styles.showPreview')}</label>
<label class="field">{t('home.styles.list')}
  <select bind:value={list}>
    <option value="recommended">{t('home.styles.recommended')}</option>
    <option value="inUse">{t('home.styles.inUse')}</option>
  </select>
</label>
