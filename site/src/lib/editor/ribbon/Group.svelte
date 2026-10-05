<script lang="ts">
  /** A ribbon group. Word for Mac draws no caption; the label is for assistive tech. */
  import type { Snippet } from 'svelte';

  type Props = { label: string; children: Snippet; class?: string };
  const { label, children, class: cls = '' }: Props = $props();
</script>

<div class="group {cls}" role="group" aria-label={label}>{@render children()}</div>
