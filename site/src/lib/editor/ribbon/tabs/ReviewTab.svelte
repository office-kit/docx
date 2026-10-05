<script lang="ts">
  import { commands } from '@office-kit/docx-editor';
  import Button from '../Button.svelte';
  import Group from '../Group.svelte';
  import { getSession } from '../../session.svelte';
  import { t } from '../../i18n/index.svelte';

  const session = getSession();
  const hasSelection = $derived(session.tick >= 0 && !!session.model?.selection);
</script>

<Group label={t('group.comments')}>
  <Button size="large" icon="comment" tip={t('review.newComment')} onclick={() => { const text = prompt('Comment'); if (text) session.apply(commands.addCommentCommand, { author: 'You', initials: 'Y', text }); }} disabled={!hasSelection} />
</Group>
<Group label={t('group.changes')}>
  <Button size="large" icon="accept" tip={t('review.acceptAll')} onclick={() => session.apply(commands.acceptAllRevisionsCommand, undefined)} />
  <Button size="large" icon="reject" tip={t('review.rejectAll')} onclick={() => session.apply(commands.rejectAllRevisionsCommand, undefined)} />
</Group>
