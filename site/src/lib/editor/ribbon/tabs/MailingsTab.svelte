<script lang="ts">
  /**
   * Word's Mailings tab: Create (Envelopes, Labels), Start Mail Merge, Write &
   * Insert Fields, Preview Results and Finish. Word for Mac has no Address
   * Block / Greeting Line buttons, so they head the Insert Merge Field menu.
   * The dialogs live beside it under `../../mailings/`.
   */
  import { ADDRESS_FIELDS, mailMergeSettings, type MailMergeDocumentType, type MergeRule } from '@office-kit/docx';
  import { commands } from '@office-kit/docx-editor';
  import Button from '../Button.svelte';
  import Group from '../Group.svelte';
  import SplitButton from '../SplitButton.svelte';
  import { getSession } from '../../session.svelte';
  import { t } from '../../i18n/index.svelte';
  import { merge, parseRecipientFile } from '../../mailings/merge.svelte';
  import { ruleDialog } from '../../mailings/rule.svelte';
  import MailingsDialogs from '../../mailings/MailingsDialogs.svelte';

  const session = getSession();
  const hasCaret = $derived(session.tick >= 0 && !!session.model?.selection);
  const settings = $derived(session.tick >= 0 && session.model ? mailMergeSettings(session.model.doc) : undefined);
  const count = $derived(merge.list?.records.length ?? 0);
  const fieldNames = $derived(merge.list?.columns ?? ADDRESS_FIELDS);

  const START_TYPES: ReadonlyArray<{ type: MailMergeDocumentType; key: `mail.start.${MailMergeDocumentType}` }> = [
    { type: 'formLetters', key: 'mail.start.formLetters' },
    { type: 'email', key: 'mail.start.email' },
    { type: 'envelopes', key: 'mail.start.envelopes' },
    { type: 'mailingLabels', key: 'mail.start.mailingLabels' },
    { type: 'catalog', key: 'mail.start.catalog' },
  ];

  // Rules without options insert at once; the others open the rule dialog.
  const RULES: ReadonlyArray<MergeRule['kind']> = ['ask', 'fillIn', 'if', 'mergeRecord', 'mergeSequence', 'nextRecord', 'nextRecordIf', 'setBookmark', 'skipRecordIf'];
  const IMMEDIATE_RULES: Partial<Record<MergeRule['kind'], MergeRule>> = {
    mergeRecord: { kind: 'mergeRecord' },
    mergeSequence: { kind: 'mergeSequence' },
    nextRecord: { kind: 'nextRecord' },
  };

  function start(type: MailMergeDocumentType | 'normal'): void {
    session.apply(commands.startMailMergeCommand, { type });
    if (type === 'mailingLabels') session.openDialog('mailings.labels');
    else if (type === 'envelopes') session.openDialog('mailings.envelopes');
    else if (type === 'normal') {
      merge.preview = false;
      merge.list = null;
    }
  }

  async function useExistingList(e: Event): Promise<void> {
    session.openMenu = null;
    const input = e.currentTarget as HTMLInputElement;
    const file = input.files?.[0];
    input.value = '';
    if (!file) return;
    try {
      merge.attach(session, file.name, parseRecipientFile(file.name, await file.text()));
    } catch (err) {
      // An unreadable or malformed list is the user's file, not a bug: report it.
      session.status = `${t('mail.selectRecipients')}: ${(err as Error).message}`;
    }
  }

  function rule(kind: MergeRule['kind']): void {
    const immediate = IMMEDIATE_RULES[kind];
    if (immediate) {
      session.apply(commands.insertRuleCommand, { rule: immediate });
      return;
    }
    ruleDialog.kind = kind;
    session.openDialog('mailings.rule');
  }

  function togglePreview(): void {
    merge.preview = !merge.preview;
    merge.show(session);
  }

  function onRecordInput(e: Event): void {
    const n = Number((e.currentTarget as HTMLInputElement).value);
    if (Number.isInteger(n)) merge.go(session, n - 1);
  }
</script>

<Group label={t('group.create')}>
  <Button size="large" icon="envelope" tip={t('mail.envelopes')} onclick={() => session.openDialog('mailings.envelopes')} />
  <Button size="large" icon="labels" tip={t('mail.labels')} onclick={() => session.openDialog('mailings.labels')} />
</Group>

<Group label={t('group.startMerge')}>
  <SplitButton id="mail-start" size="large" icon="startMerge" tip={t('mail.startMerge')}>
    {#snippet menu()}
      {#each START_TYPES as item (item.type)}
        <button class="mi check" class:checked={settings?.type === item.type} onclick={() => start(item.type)}>{t(item.key)}</button>
      {/each}
      <hr />
      <button class="mi check" class:checked={!settings} onclick={() => start('normal')}>{t('mail.start.normal')}</button>
    {/snippet}
  </SplitButton>
  <SplitButton id="mail-recipients" size="large" icon="recipients" tip={t('mail.selectRecipients')}>
    {#snippet menu()}
      <button class="mi" onclick={() => session.openDialog('mailings.newList')}>{t('mail.typeNewList')}</button>
      <label class="mi">{t('mail.useExistingList')}<input type="file" accept=".csv,.txt,.tsv,text/csv,text/plain" onchange={useExistingList} hidden /></label>
    {/snippet}
  </SplitButton>
  <div class="col">
    <Button size="mid" icon="editRecipients" tip={t('mail.editRecipients')} onclick={() => session.openDialog('mailings.recipients')} disabled={!merge.list} />
    <Button size="mid" icon="filterRecipients" tip={t('mail.filterRecipients')} onclick={() => session.openDialog('mailings.recipients')} disabled={!merge.list} />
  </div>
</Group>

<Group label={t('group.fields')}>
  <SplitButton id="mail-field" size="large" icon="mergeField" tip={t('mail.mergeField')} disabled={!hasCaret}>
    {#snippet menu()}
      <button class="mi" onclick={() => session.openDialog('mailings.addressBlock')}>{t('mail.addressBlock')}</button>
      <button class="mi" onclick={() => session.openDialog('mailings.greetingLine')}>{t('mail.greetingLine')}</button>
      <hr />
      {#each fieldNames as name (name)}
        <button class="mi" onclick={() => session.apply(commands.insertMergeFieldCommand, { name })}>{name}</button>
      {/each}
    {/snippet}
  </SplitButton>
  <div class="col">
    <SplitButton id="mail-rules" size="mid" icon="rules" tip={t('mail.rules')} disabled={!hasCaret}>
      {#snippet menu()}
        {#each RULES as kind (kind)}
          <button class="mi" onclick={() => rule(kind)}>{t(`mail.rule.${kind}`)}</button>
        {/each}
      {/snippet}
    </SplitButton>
    <Button size="mid" icon="matchFields" tip={t('mail.matchFields')} onclick={() => session.openDialog('mailings.matchFields')} disabled={!merge.list} />
    <Button size="mid" icon="updateLabels" tip={t('mail.updateLabels')} onclick={() => session.apply(commands.updateLabelsCommand, undefined)} disabled={settings?.type !== 'mailingLabels'} />
  </div>
  <Button size="mid" icon="highlightFields" tip={t('mail.highlight')} on={merge.highlight} onclick={() => merge.setHighlight(!merge.highlight)} />
</Group>

<Group label={t('group.preview')}>
  <Button size="large" icon="previewResults" tip={t('mail.preview')} on={merge.preview} onclick={togglePreview} disabled={!merge.list} />
  <div class="rows">
    <div class="row">
      <Button icon="firstRecord" tip={t('mail.firstRecord')} onclick={() => merge.go(session, 0)} disabled={count === 0 || merge.index === 0} />
      <Button icon="previousRecord" tip={t('mail.previousRecord')} onclick={() => merge.go(session, merge.index - 1)} disabled={count === 0 || merge.index === 0} />
      <label class="field"><input type="number" min="1" max={count} value={merge.index + 1} onchange={onRecordInput} disabled={count === 0} aria-label={t('mail.record')} style="width: 44px" /></label>
      <Button icon="nextRecord" tip={t('mail.nextRecord')} onclick={() => merge.go(session, merge.index + 1)} disabled={count === 0 || merge.index >= count - 1} />
      <Button icon="lastRecord" tip={t('mail.lastRecord')} onclick={() => merge.go(session, count - 1)} disabled={count === 0 || merge.index >= count - 1} />
    </div>
    <div class="row">
      <Button size="mid" icon="findRecipient" tip={t('mail.findRecipient')} onclick={() => session.openDialog('mailings.find')} disabled={count === 0} />
      <Button size="mid" icon="checkErrors" tip={t('mail.checkErrors')} onclick={() => session.openDialog('mailings.errors')} disabled={!merge.list} />
    </div>
  </div>
</Group>

<Group label={t('group.finish')}>
  <SplitButton id="mail-finish" size="large" icon="finishMerge" tip={t('mail.finish')} disabled={!merge.list}>
    {#snippet menu()}
      <button class="mi" onclick={() => session.openDialog('mailings.merge')}>{t('mail.editIndividual')}</button>
    {/snippet}
  </SplitButton>
</Group>

<MailingsDialogs />

<style>
  /* Keep icons whole when the ribbon is narrower than the tab. */
  .col :global(svg),
  .rows :global(svg) {
    flex-shrink: 0;
  }
  /* Highlight Merge Fields: Word shades merge fields grey. The canvas tags each field result with its type. */
  :global(.wk-highlight-merge :is([data-wk-field="MERGEFIELD"], [data-wk-field="ADDRESSBLOCK"], [data-wk-field="GREETINGLINE"])) {
    background: #d9d9d9;
  }
</style>
