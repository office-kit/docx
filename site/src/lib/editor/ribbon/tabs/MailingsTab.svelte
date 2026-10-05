<script lang="ts">
  import { commands } from '@office-kit/docx-editor';
  import Button from '../Button.svelte';
  import Group from '../Group.svelte';
  import { getSession } from '../../session.svelte';
  import { t } from '../../i18n/index.svelte';

  const session = getSession();
  const hasSelection = $derived(session.tick >= 0 && !!session.model?.selection);
</script>

<Group label={t('group.fields')}>
  <Button size="large" icon="newDoc" tip={t('mail.mergeField')} onclick={() => { const fieldName = prompt('Merge field name'); if (fieldName) session.apply(commands.insertMergeFieldCommand, { fieldName }); }} disabled={!hasSelection} />
</Group>
