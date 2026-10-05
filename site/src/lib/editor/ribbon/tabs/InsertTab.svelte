<script lang="ts">
  import { commands } from '@office-kit/docx-editor';
  import RibbonIcon from '../../RibbonIcon.svelte';
  import Button from '../Button.svelte';
  import Group from '../Group.svelte';
  import TableButton from '../TableButton.svelte';
  import { getSession } from '../../session.svelte';
  import { t } from '../../i18n/index.svelte';

  const session = getSession();
  const hasSelection = $derived(session.tick >= 0 && !!session.model?.selection);

  function insertLink(): void {
    const url = prompt('Link URL', 'https://');
    if (!url) return;
    const text = prompt('Link text', url) ?? url;
    session.apply(commands.insertHyperlinkCommand, { url, text });
  }

  function addComment(): void {
    const text = prompt('Comment');
    if (text) session.apply(commands.addCommentCommand, { author: 'You', initials: 'Y', text });
  }

  async function insertImage(e: Event): Promise<void> {
    const input = e.target as HTMLInputElement;
    const file = input.files?.[0];
    if (!file) return;
    const bytes = new Uint8Array(await file.arrayBuffer());
    session.apply(commands.insertImageCommand, { bytes, options: { widthEmu: 2743200, heightEmu: 2057400 } });
    input.value = '';
  }
</script>

<Group label={t('group.pages')}>
  <Button size="large" icon="pageBreak" tip={t('ins.pageBreak')} onclick={() => session.apply(commands.insertPageBreakCommand, undefined)} disabled={!hasSelection} />
</Group>
<Group label={t('group.tables')}>
  <TableButton />
</Group>
<Group label={t('group.illustrations')}>
  <label class="big" title={t('ins.picture')}><RibbonIcon name="picture" size={32} /><span>{t('ins.picture')}</span><input type="file" accept="image/*" onchange={insertImage} hidden /></label>
</Group>
<Group label={t('group.links')}>
  <Button size="large" icon="link" tip={t('ins.link')} onclick={insertLink} />
  <Button size="large" icon="bookmark" tip={t('ref.bookmark')} onclick={() => { const n = prompt('Bookmark name'); if (n) session.apply(commands.addBookmarkCommand, { name: n }); }} />
</Group>
<Group label={t('group.comments')}>
  <Button size="large" icon="comment" tip={t('ins.comment')} onclick={addComment} disabled={!hasSelection} />
</Group>
<Group label={t('group.headerFooter')}>
  <Button size="large" icon="pageNumber" tip={t('layout.pageNumbers')} onclick={() => session.apply(commands.addPageNumberFooterCommand, {})} />
</Group>
