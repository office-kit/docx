<script lang="ts">
  import type { ValidationIssue } from '@office-kit/docx';

  type Props = {
    /** What `validate(doc)` returned. */
    issues: readonly ValidationIssue[];
  };

  const { issues }: Props = $props();
</script>

{#if issues.length === 0}
  <p class="clean"><code>validate(doc)</code> found nothing wrong with the package.</p>
{:else}
  <ul class="issues">
    {#each issues as issue, i (i)}
      <li class="issue-{issue.level}">
        <span class="level">{issue.level}</span>
        <span class="message">{issue.message}</span>
        {#if issue.partName}<code>{issue.partName}</code>{/if}
      </li>
    {/each}
  </ul>
{/if}

<style>
  .clean {
    margin: 0;
    color: var(--ink-2);
  }

  .issues {
    list-style: none;
    margin: 0;
    padding: 0;
  }

  li {
    display: flex;
    flex-wrap: wrap;
    align-items: baseline;
    gap: 0.4rem 0.75rem;
    margin: 0;
    padding: 0.7rem 0;
    border-bottom: 1px solid var(--line);
    font-size: 0.93rem;
  }

  .level {
    flex: none;
    padding: 0.05rem 0.5rem;
    border-radius: 999px;
    background: var(--wash);
    border: 1px solid var(--line-strong);
    font-size: 0.78rem;
    font-weight: 600;
  }

  .issue-error .level {
    background: var(--accent-wash);
    border-color: var(--accent);
    color: var(--accent-ink);
  }

  .message {
    flex: 1 1 16rem;
  }
</style>
