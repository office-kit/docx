<script lang="ts">
  import { commands } from '@office-kit/docx-editor';
  import { PAGE_SIZE_A4, PAGE_SIZE_LETTER } from '@office-kit/docx';
  import Button from '../Button.svelte';
  import Group from '../Group.svelte';
  import { getSession } from '../../session.svelte';
  import { t } from '../../i18n/index.svelte';

  const session = getSession();
  const hasSelection = $derived(session.tick >= 0 && !!session.model?.selection);
</script>

<Group label={t('group.pageSetup')}>
  <Button size="large" icon="orientation" tip={t('layout.portrait')} onclick={() => session.apply(commands.setOrientationCommand, { orientation: 'portrait' })} />
  <Button size="large" icon="orientation" tip={t('layout.landscape')} onclick={() => session.apply(commands.setOrientationCommand, { orientation: 'landscape' })} />
  <Button size="large" icon="size" tip="A4" onclick={() => session.apply(commands.setPageSizeCommand, { size: PAGE_SIZE_A4 })} />
  <Button size="large" icon="size" tip="Letter" onclick={() => session.apply(commands.setPageSizeCommand, { size: PAGE_SIZE_LETTER })} />
  <Button size="large" icon="sectionBreak" tip={t('layout.sectionBreak')} onclick={() => session.apply(commands.insertSectionBreakCommand, { type: 'nextPage' })} disabled={!hasSelection} />
</Group>
<Group label={t('group.paragraph')}>
  <Button icon="indentLess" tip={t('tip.indentLess')} onclick={() => session.apply(commands.indentStepCommand, { direction: 'decrease' })} disabled={!session.enabled(commands.indentStepCommand)} />
  <Button icon="indentMore" tip={t('tip.indentMore')} onclick={() => session.apply(commands.indentStepCommand, { direction: 'increase' })} disabled={!session.enabled(commands.indentStepCommand)} />
</Group>
