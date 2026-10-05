<script lang="ts">
  /** Review ▸ Check Accessibility: findings grouped as Errors, Warnings and Tips. */
  import { checkAccessibility, type AccessibilityIssueKind, type AccessibilitySeverity } from '@office-kit/docx';
  import RibbonIcon from '../RibbonIcon.svelte';
  import { getSession } from '../session.svelte';
  import { t, type MessageKey } from '../i18n/index.svelte';
  import { goToBlock } from '../review-actions';

  const session = getSession();
  const issues = $derived(session.tick >= 0 && session.model ? checkAccessibility(session.model.doc) : []);

  const GROUPS: readonly [AccessibilitySeverity, MessageKey][] = [
    ['error', 'review.a11yErrors'],
    ['warning', 'review.a11yWarnings'],
    ['tip', 'review.a11yTips'],
  ];
  const NAMES: Readonly<Record<AccessibilityIssueKind, MessageKey>> = {
    missingAltText: 'review.issue.missingAltText',
    missingTableHeader: 'review.issue.missingTableHeader',
    skippedHeadingLevel: 'review.issue.skippedHeadingLevel',
    lowContrast: 'review.issue.lowContrast',
    repeatedBlankParagraphs: 'review.issue.repeatedBlankParagraphs',
    missingTitle: 'review.issue.missingTitle',
  };
</script>

<div class="pane-head">
  <span>{t('review.a11yTitle')}</span>
  <button class="pane-close" onclick={() => (session.pane.right = null)} aria-label={t('group.close')}><RibbonIcon name="close" size={14} /></button>
</div>
{#if issues.length === 0}
  <p class="pane-note">{t('review.a11yNone')}</p>
{/if}
{#each GROUPS as [severity, label] (severity)}
  {@const found = issues.filter((i) => i.severity === severity)}
  {#if found.length}
    <div class="a11y-group">{t(label)} ({found.length})</div>
    {#each found as issue, i (i)}
      <button class="a11y-issue" disabled={issue.block === undefined} onclick={() => issue.block !== undefined && goToBlock(session, issue.block)}>
        <strong>{t(NAMES[issue.kind])}</strong>
        <span>{issue.detail}</span>
      </button>
    {/each}
  {/if}
{/each}
<label class="check-item keep"><input type="checkbox" bind:checked={session.prefs.keepAccessibilityRunning} />{t('review.a11yKeep')}</label>

<style>
  .a11y-group { margin-top: 6px; font-weight: 600; }
  .a11y-issue { display: flex; flex-direction: column; gap: 2px; padding: 6px 8px; border: none; border-radius: 6px; background: #fff; text-align: left; cursor: pointer; }
  .a11y-issue:hover:not(:disabled) { background: var(--hover); }
  .a11y-issue:disabled { cursor: default; }
  .a11y-issue span { color: var(--muted); overflow-wrap: anywhere; }
  .keep { margin-top: auto; }
</style>
