<script lang="ts">
  import { commands } from '@office-kit/docx-editor';
  import Button from '../Button.svelte';
  import Group from '../Group.svelte';
  import IllustrationButtons from '../IllustrationButtons.svelte';
  import { getSession } from '../../session.svelte';
  import { t } from '../../i18n/index.svelte';

  const session = getSession();
  const hasSelection = $derived(session.tick >= 0 && !!session.model?.selection);

  function insertTable(): void {
    const spec = prompt('Table size (rows x cols)', '3x3');
    if (!spec) return;
    // Invalid sizes are rejected by the command itself and shown in the status bar.
    const [rows = Number.NaN, cols = Number.NaN] = spec.split(/[x×,]/).map((n) => Number(n.trim()));
    session.apply(commands.insertTableCommand, { rows, cols });
  }

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
</script>

<Group label={t('group.pages')}>
  <Button size="large" icon="pageBreak" tip={t('ins.pageBreak')} onclick={() => session.apply(commands.insertPageBreakCommand, undefined)} disabled={!hasSelection} />
</Group>
<Group label={t('group.tables')}>
  <Button size="large" icon="table" tip={t('ins.table')} onclick={insertTable} />
</Group>
<Group label={t('group.illustrations')}>
  <IllustrationButtons />
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
