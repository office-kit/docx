<script lang="ts">
  import { base } from '$app/paths';
  import CodeBlock from '$lib/components/CodeBlock.svelte';
  import type { PageProps } from './$types';

  const { data }: PageProps = $props();
</script>

<svelte:head>
  <title>Recipes · @office-kit/docx</title>
</svelte:head>

<h1>Recipes</h1>

<p class="lede">
  Working code for common tasks. Every snippet is a real file under
  <code>site/src/lib/examples/</code> that is type-checked against the library on every build, so
  an API rename breaks this page before it ships. For a specific function, see the
  <a href="{base}/api">API reference</a>.
</p>

<nav class="jump" aria-label="Recipes on this page">
  {#each data.recipes as r (r.key)}
    <a href="#{r.key}">{r.title}</a>
  {/each}
</nav>

{#each data.recipes as r (r.key)}
  <section class="recipe" id={r.key}>
    <h2><a href="#{r.key}">{r.title}</a></h2>
    <p>{r.description}</p>
    <CodeBlock html={r.html} title={r.path} />
    {#if r.seeAlso}
      <p class="see-also">{r.seeAlso}</p>
    {/if}
  </section>
{/each}

<p class="more">
  To build a document end to end, walk through
  <a href="{base}/docs/getting-started">Getting started</a>, then open
  <a href="{base}/playground">the playground</a> to see the rendered output.
</p>

<style>
  .lede {
    color: var(--ink-2);
    font-size: 1.08rem;
  }

  .jump {
    display: flex;
    flex-wrap: wrap;
    gap: 0.4rem;
    margin: 1.75rem 0 0;
  }

  .jump a {
    padding: 0.35rem 0.75rem;
    border: 1px solid var(--line);
    border-radius: 999px;
    color: var(--ink-2);
    font-size: 0.9rem;
    font-weight: 500;
  }

  .jump a:hover {
    color: var(--ink);
    border-color: var(--ink-3);
    text-decoration: none;
  }

  .recipe {
    margin-top: 3rem;
    scroll-margin-top: calc(var(--header-h) + 1rem);
  }

  .recipe h2 {
    margin: 0 0 0.4rem;
  }

  .recipe h2 a {
    color: inherit;
  }

  .recipe > p {
    margin: 0;
    color: var(--ink-2);
  }

  .recipe .see-also {
    font-size: 0.9rem;
  }

  .more {
    margin-top: 3.5rem;
    padding-top: 1.5rem;
    border-top: 1px solid var(--line);
    color: var(--ink-2);
  }
</style>
