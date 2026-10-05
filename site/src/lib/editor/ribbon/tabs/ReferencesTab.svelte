<script lang="ts">
  import { commands } from '@office-kit/docx-editor';
  import Button from '../Button.svelte';
  import Group from '../Group.svelte';
  import { getSession } from '../../session.svelte';
  import { t } from '../../i18n/index.svelte';

  const session = getSession();
  const hasSelection = $derived(session.tick >= 0 && !!session.model?.selection);
</script>

<Group label={t('group.toc')}>
  <Button size="large" icon="toc" tip={t('ref.toc')} onclick={() => session.apply(commands.insertTocCommand, {})} />
</Group>
<Group label={t('group.footnotes')}>
  <Button size="large" icon="footnote" tip={t('ref.footnote')} onclick={() => { const txt = prompt('Footnote text'); if (txt) session.apply(commands.addFootnoteCommand, { text: txt }); }} disabled={!hasSelection} />
  <Button size="large" icon="endnote" tip={t('ref.endnote')} onclick={() => { const txt = prompt('Endnote text'); if (txt) session.apply(commands.addEndnoteCommand, { text: txt }); }} disabled={!hasSelection} />
</Group>
