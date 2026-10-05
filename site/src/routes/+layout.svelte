<script lang="ts">
  import '@office-kit/site-kit/site.css';
  import { KitFooter, KitHeader, KitSeo, type NavLink } from '@office-kit/site-kit';
  import { onMount } from 'svelte';
  import { goto } from '$app/navigation';
  import { base } from '$app/paths';
  import { page } from '$app/state';
  import CodeCopyEnhancer from '$lib/components/CodeCopyEnhancer.svelte';
  import LanguageSwitcher from '$lib/components/LanguageSwitcher.svelte';
  import Search from '$lib/components/Search.svelte';
  import {
    DEFAULT_LOCALE,
    locale,
    localized,
    localizePath,
    routePath,
    storedLocale,
  } from '$lib/i18n';
  import core from '$lib/i18n/messages/core';

  type Props = {
    children?: import('svelte').Snippet;
  };

  const { children }: Props = $props();

  const m = $derived(localized(core));
  // The editor is a full-window app with its own title bar, like Word; the
  // site's header and footer would only take room from the page.
  const chrome = $derived(page.route.id !== '/editor');
  const at = (path: string): string => localizePath(path, locale());

  const links: NavLink[] = $derived([
    { path: at('/docs/getting-started'), label: m.nav.docs, section: at('/docs') },
    { path: at('/docs/recipes'), label: m.nav.recipes },
    { path: at('/api'), label: m.nav.api },
    { path: at('/playground'), label: m.nav.playground },
    { path: at('/repl'), label: m.nav.repl },
    { path: '/editor', label: m.nav.editor },
  ]);

  // Prerendering sets `<html lang>`; a client-side navigation that changes
  // locale (the site-kit header links home in English) must update it.
  $effect(() => {
    document.documentElement.lang = locale();
  });

  // An unprefixed URL means "no explicit language", so a visitor who picked
  // one in the switcher before is sent to that translation.
  onMount(() => {
    const path = routePath();
    const stored = storedLocale();
    if (path === undefined || locale() !== DEFAULT_LOCALE || !stored || stored === DEFAULT_LOCALE) {
      return;
    }
    void goto(`${base}${localizePath(path, stored)}${page.url.hash}`, { replaceState: true });
  });
</script>

<KitSeo product="docx" title={m.seo.title} description={m.seo.description} />

{#if chrome}
  <a class="skip" href="#main">{m.skip}</a>

  <KitHeader product="docx" {links}>
    <!-- The shell's only slot for site controls; the language switch sits
         beside search there. -->
    {#snippet search()}
      <Search />
      <LanguageSwitcher />
    {/snippet}
  </KitHeader>
{/if}

<main id="main" class:app={!chrome}>
  {@render children?.()}
</main>

<CodeCopyEnhancer />

{#if chrome}
  <KitFooter product="docx" {links} />
{/if}

<style>
  main {
    min-height: calc(100vh - var(--header-h));
  }

  main.app {
    min-height: 100vh;
  }

  .skip {
    position: absolute;
    left: 1rem;
    top: -4rem;
    z-index: 40;
    padding: 0.6rem 1rem;
    border-radius: var(--radius);
    background: var(--ink);
    color: var(--paper);
  }

  .skip:focus {
    top: 0.75rem;
  }

  /* Japanese has no spaces, so a heading would otherwise wrap mid-word. */
  :global(:root:lang(ja) :is(h1, h2, h3)) {
    word-break: auto-phrase;
  }
</style>
