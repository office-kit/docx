<script lang="ts">
  import { afterNavigate } from '$app/navigation';
  import RichText from '$lib/components/RichText.svelte';
  import Sidebar from '$lib/components/Sidebar.svelte';
  import { allDocLinks } from '$lib/docs-nav';
  import { href, locale, localized, routePath } from '$lib/i18n';
  import docs from '$lib/i18n/messages/docs';

  type Props = {
    children?: import('svelte').Snippet;
  };

  const { children }: Props = $props();

  const m = $derived(localized(docs).layout);
  const links = $derived(allDocLinks(locale()));
  const index = $derived(links.findIndex((l) => l.href === routePath()));
  const current = $derived(links[index]);
  const prev = $derived(index > 0 ? links[index - 1] : undefined);
  const next = $derived(index >= 0 ? links[index + 1] : undefined);

  // The sidebar is always open on wide screens; this only governs the phone drawer.
  let navOpen = $state(false);
  afterNavigate(() => (navOpen = false));
</script>

<div class="docs frame">
  <aside class="sidebar" data-pagefind-ignore>
    <button
      type="button"
      class="sidebar-toggle"
      aria-expanded={navOpen}
      aria-controls="docs-nav"
      onclick={() => (navOpen = !navOpen)}
    >
      <span>{current?.title ?? m.documentation}</span>
      <svg width="12" height="12" viewBox="0 0 10 10" aria-hidden="true">
        <path d="M1.5 3.5 5 7l3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.6" />
      </svg>
    </button>
    <div id="docs-nav" class="sidebar-body" class:open={navOpen}>
      <Sidebar />
    </div>
  </aside>

  <div class="main">
    <article class="doc-content">
      {@render children?.()}
    </article>

    {#if current}
      <nav class="pager" aria-label={m.pager} data-pagefind-ignore>
        {#if prev}
          <a href={href(prev.href)} class="prev">
            <span>{m.previous}</span>
            {prev.title}
          </a>
        {/if}
        {#if next}
          <a href={href(next.href)} class="next">
            <span>{m.next}</span>
            {next.title}
          </a>
        {/if}
      </nav>
      <p class="llms-link" data-pagefind-ignore><RichText text={m.llms} /></p>
    {/if}
  </div>
</div>

<style>
  .docs {
    display: flex;
    align-items: stretch;
  }

  .sidebar {
    flex: 0 0 var(--sidebar-w);
    width: var(--sidebar-w);
    border-right: 1px solid var(--line);
  }

  .sidebar-body {
    position: sticky;
    top: var(--header-h);
    max-height: calc(100vh - var(--header-h));
    overflow-y: auto;
    padding: 2.25rem 1rem 3rem var(--gutter);
  }

  .sidebar-toggle {
    display: none;
  }

  .main {
    flex: 1;
    min-width: 0;
    padding: 2.75rem clamp(1.25rem, 4vw, 3.5rem) 4rem;
  }

  .doc-content,
  .pager,
  .llms-link {
    max-width: var(--measure);
  }

  /* Wide content scrolls in its own box so a long table or signature never
   * pushes the page sideways on a phone. */
  .doc-content :global(table) {
    display: block;
    overflow-x: auto;
  }

  .doc-content :global(h2) {
    margin-top: 3rem;
  }

  .pager {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 0.75rem;
    margin-top: 3.5rem;
  }

  .pager a {
    display: flex;
    flex-direction: column;
    gap: 0.15rem;
    padding: 0.9rem 1.1rem;
    border: 1px solid var(--line);
    border-radius: var(--radius);
    color: var(--ink);
    font-weight: 600;
  }

  .pager a:hover {
    border-color: var(--accent);
    text-decoration: none;
  }

  .pager span {
    color: var(--ink-3);
    font-size: 0.82rem;
    font-weight: 450;
  }

  .pager .next {
    grid-column: 2;
    text-align: right;
  }

  .llms-link {
    margin: 1.5rem 0 0;
    color: var(--ink-3);
    font-size: 0.88rem;
  }

  @media (max-width: 860px) {
    .docs {
      flex-direction: column;
    }

    .sidebar {
      position: sticky;
      top: var(--header-h);
      z-index: 20;
      flex: none;
      width: 100%;
      border-right: none;
      border-bottom: 1px solid var(--line);
      background: var(--paper);
    }

    .sidebar-toggle {
      display: flex;
      align-items: center;
      justify-content: space-between;
      width: 100%;
      min-height: 48px;
      padding: 0 var(--gutter);
      border: none;
      background: transparent;
      color: var(--ink);
      font-family: var(--sans);
      font-size: 0.95rem;
      font-weight: 600;
      cursor: pointer;
    }

    .sidebar-toggle[aria-expanded='true'] svg {
      transform: rotate(180deg);
    }

    .sidebar-body {
      display: none;
      position: static;
      max-height: calc(100dvh - var(--header-h) - 48px);
      padding: 0.5rem var(--gutter) 1.5rem;
      border-top: 1px solid var(--line);
    }

    .sidebar-body.open {
      display: block;
    }

    .main {
      padding-top: 2rem;
    }
  }

  @media (max-width: 480px) {
    .pager {
      grid-template-columns: 1fr;
    }

    .pager .next {
      grid-column: 1;
      text-align: left;
    }
  }
</style>
