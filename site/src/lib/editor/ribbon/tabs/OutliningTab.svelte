<script lang="ts">
  /**
   * Word's Outlining tab, shown while the Outline view is on: outline levels,
   * promote / demote, moving paragraphs, and which levels are shown.
   * Expand / Collapse step Show Level by one, which is what Word does when the
   * whole outline is selected; collapsing a single heading's subtree needs
   * per-heading state that the document does not store.
   */
  import { BODY_TEXT_LEVEL, commands, outlineLevelOf, paragraphAt } from '@office-kit/docx-editor';
  import Button from '../Button.svelte';
  import Group from '../Group.svelte';
  import { getSession } from '../../session.svelte';
  import { t } from '../../i18n/index.svelte';

  const session = getSession();
  const prefs = session.prefs;
  const HEADING_LEVELS = [1, 2, 3, 4, 5, 6, 7, 8, 9] as const;

  const level = $derived.by(() => {
    if (session.tick < 0 || !session.model?.selection) return undefined;
    const para = paragraphAt(session.model.doc, session.model.selection.focus);
    return para ? outlineLevelOf(para) : undefined;
  });
  const canEdit = $derived(session.enabled(commands.promoteCommand));

  function levelName(n: number): string {
    return n === BODY_TEXT_LEVEL ? t('outline.bodyText') : t('outline.levelN').replace('{n}', String(n));
  }

  function setLevel(n: number): void {
    session.apply(commands.setOutlineLevelCommand, { level: n });
  }

  function stepShowLevel(delta: number): void {
    prefs.outlineShowLevel = Math.min(BODY_TEXT_LEVEL, Math.max(1, prefs.outlineShowLevel + delta));
  }
</script>

<Group label={t('group.outlineTools')}>
  <div class="rows">
    <div class="row">
      <Button icon="outlinePromoteFull" tip={t('outline.promoteToHeading1')} onclick={() => setLevel(1)} disabled={!canEdit} />
      <Button icon="outlinePromote" tip={t('outline.promote')} onclick={() => session.apply(commands.promoteCommand, undefined)} disabled={!canEdit} />
      <label class="field" title={t('outline.level')}>
        <select value={level ?? ''} onchange={(e) => setLevel(Number(e.currentTarget.value))} disabled={!canEdit} aria-label={t('outline.level')}>
          {#each [...HEADING_LEVELS, BODY_TEXT_LEVEL] as n (n)}<option value={n}>{levelName(n)}</option>{/each}
        </select>
      </label>
      <Button icon="outlineDemote" tip={t('outline.demote')} onclick={() => session.apply(commands.demoteCommand, undefined)} disabled={!canEdit} />
      <Button icon="outlineDemoteBody" tip={t('outline.demoteToBody')} onclick={() => setLevel(BODY_TEXT_LEVEL)} disabled={!canEdit} />
    </div>
    <div class="row">
      <Button icon="outlineMoveUp" tip={t('outline.moveUp')} onclick={() => session.apply(commands.moveBlocksCommand, { direction: -1 })} disabled={!session.enabled(commands.moveBlocksCommand)} />
      <Button icon="outlineMoveDown" tip={t('outline.moveDown')} onclick={() => session.apply(commands.moveBlocksCommand, { direction: 1 })} disabled={!session.enabled(commands.moveBlocksCommand)} />
      <Button icon="outlineExpand" tip={t('outline.expand')} onclick={() => stepShowLevel(1)} />
      <Button icon="outlineCollapse" tip={t('outline.collapse')} onclick={() => stepShowLevel(-1)} />
      <label class="field">
        {t('outline.showLevel')}
        <select bind:value={prefs.outlineShowLevel} aria-label={t('outline.showLevel')}>
          {#each HEADING_LEVELS as n (n)}<option value={n}>{levelName(n)}</option>{/each}
          <option value={BODY_TEXT_LEVEL}>{t('outline.allLevels')}</option>
        </select>
      </label>
    </div>
  </div>
  <div class="rows">
    <label class="check-item"><input type="checkbox" bind:checked={prefs.outlineShowFormatting} />{t('outline.showFormatting')}</label>
    <label class="check-item"><input type="checkbox" bind:checked={prefs.outlineFirstLineOnly} />{t('outline.firstLineOnly')}</label>
  </div>
</Group>
<Group label={t('group.close')}>
  <Button size="large" icon="outlineClose" tip={t('outline.close')} onclick={() => session.setView('print')} />
</Group>
