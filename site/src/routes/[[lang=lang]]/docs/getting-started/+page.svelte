<script lang="ts">
  import CodeBlock from '$lib/components/CodeBlock.svelte';
  import RichText from '$lib/components/RichText.svelte';
  import { localized } from '$lib/i18n';
  import docs from '$lib/i18n/messages/docs';
  import type { PageProps } from './$types';

  const { data }: PageProps = $props();

  const m = $derived(localized(docs).gettingStarted);
</script>

<svelte:head>
  <title>{m.title} · @office-kit/docx</title>
</svelte:head>

<h1>{m.title}</h1>

<p class="lede"><RichText text={m.lede} /></p>

<h2>{m.install}</h2>

<pre><code>pnpm add @office-kit/docx @office-kit/docx-preview</code></pre>

<p><RichText text={m.installBody} /></p>

<h2>{m.fromScratch}</h2>

<p><RichText text={m.fromScratchBody} /></p>

<CodeBlock html={data.fromScratch.html} title={data.fromScratch.path} />

<h2>{m.template}</h2>

{#each m.templateBody as text, i (i)}
  <p><RichText {text} /></p>
{/each}

<CodeBlock html={data.templateFill.html} title={data.templateFill.path} />

<h2>{m.browser}</h2>

{#each m.browserBody as text, i (i)}
  <p><RichText {text} /></p>
{/each}

<h2>{m.next}</h2>

<ul>
  {#each m.nextLinks as text, i (i)}
    <li><RichText {text} /></li>
  {/each}
</ul>

<style>
  .lede {
    color: var(--ink-2);
    font-size: 1.08rem;
  }
</style>
