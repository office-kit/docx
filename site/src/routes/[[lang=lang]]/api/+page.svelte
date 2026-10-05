<script lang="ts">
  import RichText from '$lib/components/RichText.svelte';
  import { apiGroups, apiTotalCount } from '$lib/api-groups';
  import { localized, translate } from '$lib/i18n';
  import api from '$lib/i18n/messages/api';

  // The listing lives in `$lib/api-groups` so `/llms-full.txt` and the
  // `check:api-page` CI gate read the same source as this page.

  // eslint-disable-next-line prefer-const -- reassigned by `bind:value` in template
  let filter = $state('');

  const m = $derived(localized(api));

  // Anchors come from the English title so a link to a group works in every locale.
  const slug = (title: string): string =>
    title
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, '-')
      .replace(/^-|-$/g, '');

  const groups = $derived(
    apiGroups.map((g) => ({
      id: slug(g.title),
      text: translate({ title: g.title, description: g.description }, g.translations),
      entries: g.entries,
    })),
  );

  const visible = $derived.by(() => {
    const q = filter.trim().toLowerCase();
    if (!q) return groups;
    return groups
      .map((g) => ({
        id: g.id,
        text: g.text,
        entries: g.entries.filter((e) => e.name.toLowerCase().includes(q)),
      }))
      .filter((g) => g.entries.length > 0);
  });

  const visibleCount = $derived(visible.reduce((n, g) => n + g.entries.length, 0));
</script>

<svelte:head>
  <title>{m.title} · @office-kit/docx</title>
</svelte:head>

<div class="api frame">
  <aside class="toc" data-pagefind-ignore>
    <nav aria-label={m.groupsLabel}>
      <h2>{m.groups}</h2>
      <ul>
        {#each groups as group (group.id)}
          <li><a href="#{group.id}">{group.text.title}</a></li>
        {/each}
      </ul>
    </nav>
  </aside>

  <div class="content">
    <h1>{m.title}</h1>
    <p class="lede"><RichText text={m.lede(apiTotalCount, apiGroups.length)} /></p>
    <p class="lede"><RichText text={m.sync} /></p>

    <div class="filter" data-pagefind-ignore>
      <label for="api-filter">{m.filter}</label>
      <input
        id="api-filter"
        type="search"
        bind:value={filter}
        placeholder={m.placeholder}
        autocomplete="off"
        spellcheck="false"
      />
      <p aria-live="polite">
        {#if filter.trim()}
          {visibleCount === 0 ? m.noMatch(filter.trim()) : m.count(visibleCount, apiTotalCount)}
        {/if}
      </p>
    </div>

    {#each visible as group (group.id)}
      <section class="group" id={group.id}>
        <h2>{group.text.title}</h2>
        <p class="g-desc">{group.text.description}</p>
        <ul class="exports">
          {#each group.entries as entry (entry.name)}
            <li>
              <code>{entry.name}</code>
              {#if entry.sig}<span class="sig">{entry.sig}</span>{/if}
            </li>
          {/each}
        </ul>
      </section>
    {/each}
  </div>
</div>

<style>
  .api {
    display: flex;
    align-items: stretch;
  }

  .toc {
    flex: 0 0 var(--sidebar-w);
    width: var(--sidebar-w);
    border-right: 1px solid var(--line);
  }

  .toc nav {
    position: sticky;
    top: var(--header-h);
    max-height: calc(100vh - var(--header-h));
    overflow-y: auto;
    padding: 2.25rem 1rem 3rem var(--gutter);
  }

  .toc h2 {
    margin: 0 0 0.5rem;
    font-family: var(--sans);
    font-size: 0.85rem;
    font-weight: 600;
    letter-spacing: 0;
    color: var(--ink-3);
  }

  .toc ul {
    list-style: none;
    margin: 0;
    padding: 0;
    border-left: 1px solid var(--line);
  }

  .toc li {
    margin: 0;
  }

  .toc a {
    display: block;
    padding: 0.28rem 0 0.28rem 0.95rem;
    color: var(--ink-2);
    font-size: 0.9rem;
  }

  .toc a:hover {
    color: var(--accent-ink);
    text-decoration: none;
  }

  .content {
    flex: 1;
    min-width: 0;
    padding: 2.75rem clamp(1.25rem, 4vw, 3.5rem) 5rem;
  }

  .lede {
    max-width: 68ch;
    color: var(--ink-2);
    font-size: 1.05rem;
  }

  .filter {
    position: sticky;
    top: var(--header-h);
    z-index: 5;
    display: grid;
    grid-template-columns: minmax(0, 24rem) 1fr;
    gap: 0.35rem 1rem;
    align-items: center;
    margin: 2rem 0 0;
    padding: 0.85rem 0;
    background: var(--paper);
    border-bottom: 1px solid var(--line);
  }

  .filter label {
    grid-column: 1 / -1;
    font-size: 0.85rem;
    font-weight: 600;
    color: var(--ink-2);
  }

  .filter input {
    width: 100%;
    height: 42px;
    padding: 0 0.85rem;
    border: 1px solid var(--line-strong);
    border-radius: var(--radius);
    background: var(--paper);
    color: var(--ink);
    font-family: var(--mono);
    /* 16px keeps iOS Safari from zooming the page when the field takes focus. */
    font-size: 1rem;
  }

  .filter input:focus-visible {
    outline-offset: 0;
    border-color: var(--accent);
  }

  .filter p {
    margin: 0;
    color: var(--ink-2);
    font-size: 0.9rem;
  }

  .group {
    margin-top: 3rem;
    scroll-margin-top: calc(var(--header-h) + 7rem);
  }

  .group h2 {
    margin: 0;
    font-size: 1.4rem;
  }

  .g-desc {
    max-width: 68ch;
    margin: 0.35rem 0 1rem;
    color: var(--ink-2);
    font-size: 0.97rem;
  }

  .exports {
    list-style: none;
    margin: 0;
    padding: 0;
    border-top: 1px solid var(--line);
    columns: 2;
    column-gap: 2.5rem;
  }

  .exports li {
    display: flex;
    flex-direction: column;
    gap: 0.15rem;
    margin: 0;
    padding: 0.5rem 0;
    border-bottom: 1px solid var(--line);
    break-inside: avoid;
  }

  .exports code {
    padding: 0;
    border: none;
    background: none;
    color: var(--accent-ink);
    font-size: 0.9rem;
    font-weight: 500;
  }

  .sig {
    color: var(--ink-2);
    font-family: var(--mono);
    font-size: 0.78rem;
    line-height: 1.5;
    overflow-wrap: anywhere;
  }

  @media (max-width: 1000px) {
    .toc {
      display: none;
    }
  }

  @media (max-width: 640px) {
    .exports {
      columns: 1;
    }

    .filter {
      grid-template-columns: 1fr;
    }
  }
</style>
