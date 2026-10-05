<script lang="ts">
  import CodeBlock from '$lib/components/CodeBlock.svelte';
  import RichText from '$lib/components/RichText.svelte';
  import { localized, translate } from '$lib/i18n';
  import docs from '$lib/i18n/messages/docs';
  import type { PageProps } from './$types';

  const { data }: PageProps = $props();

  const m = $derived(localized(docs).recipes);
  const recipes = $derived(
    data.recipes.map((r) => ({
      key: r.key,
      path: r.path,
      html: r.html,
      text: translate(r.text, r.translations),
    })),
  );
</script>

<svelte:head>
  <title>{m.title} · @office-kit/docx</title>
</svelte:head>

<h1>{m.title}</h1>

<p class="lede"><RichText text={m.lede} /></p>

<nav class="jump" aria-label={m.jump}>
  {#each recipes as r (r.key)}
    <a href="#{r.key}">{r.text.title}</a>
  {/each}
</nav>

{#each recipes as r (r.key)}
  <section class="recipe" id={r.key}>
    <h2><a href="#{r.key}">{r.text.title}</a></h2>
    <p>{r.text.description}</p>
    <CodeBlock html={r.html} title={r.path} />
    {#if r.text.seeAlso}
      <p class="see-also">{r.text.seeAlso}</p>
    {/if}
  </section>
{/each}

<p class="more"><RichText text={m.more} /></p>

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
