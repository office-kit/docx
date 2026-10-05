<script lang="ts">
  import { docSections } from '$lib/docs-nav';
  import { href, locale, localized, routePath } from '$lib/i18n';
  import docs from '$lib/i18n/messages/docs';

  const current = $derived(routePath());
</script>

<nav class="sidebar-nav" aria-label={localized(docs).layout.documentation}>
  {#each docSections(locale()) as section (section.title)}
    <section>
      <h2>{section.title}</h2>
      <ul>
        {#each section.links as link (link.href)}
          <li>
            <a href={href(link.href)} aria-current={link.href === current ? 'page' : undefined}
              >{link.title}</a
            >
          </li>
        {/each}
      </ul>
    </section>
  {/each}
</nav>

<style>
  section + section {
    margin-top: 1.75rem;
  }

  h2 {
    margin: 0 0 0.5rem;
    font-family: var(--sans);
    font-size: 0.85rem;
    font-weight: 600;
    letter-spacing: 0;
    color: var(--ink-3);
  }

  ul {
    list-style: none;
    padding: 0;
    margin: 0;
    border-left: 1px solid var(--line);
  }

  li {
    margin: 0;
  }

  a {
    display: block;
    margin-left: -1px;
    padding: 0.4rem 0 0.4rem 0.95rem;
    border-left: 1px solid transparent;
    color: var(--ink-2);
    font-size: 0.95rem;
    line-height: 1.35;
  }

  a:hover {
    color: var(--ink);
    text-decoration: none;
  }

  a[aria-current='page'] {
    border-left-color: var(--accent);
    color: var(--accent-ink);
    font-weight: 600;
  }
</style>
