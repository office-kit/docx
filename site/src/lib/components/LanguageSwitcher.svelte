<script lang="ts">
  import { base } from '$app/paths';
  import {
    isSiteLocale,
    LOCALES,
    locale,
    localized,
    localizePath,
    routePath,
    storeLocale,
    type SiteLocale,
  } from '$lib/i18n';
  import core from '$lib/i18n/messages/core';

  const m = $derived(localized(core));
  const path = $derived(routePath());
  const others = $derived(
    Object.keys(LOCALES).filter((id): id is SiteLocale => isSiteLocale(id) && id !== locale()),
  );
</script>

<!-- A page without translations (the editor) offers no switch. Switching is a
     full page load so that `<html lang>`, and the search index that follows
     it, change along with the page. -->
{#if path !== undefined}
  {#each others as id (id)}
    <a
      class="switch"
      href="{base}{localizePath(path, id)}"
      hreflang={id}
      lang={id}
      aria-label="{m.language.switchTo} {LOCALES[id]}"
      data-sveltekit-reload
      onclick={() => storeLocale(id)}
    >
      <svg width="16" height="16" viewBox="0 0 16 16" aria-hidden="true">
        <circle cx="8" cy="8" r="6.2" fill="none" stroke="currentColor" stroke-width="1.4" />
        <path
          d="M1.8 8h12.4M8 1.8c1.8 1.8 2.6 3.8 2.6 6.2S9.8 12.4 8 14.2C6.2 12.4 5.4 10.4 5.4 8S6.2 3.6 8 1.8Z"
          fill="none"
          stroke="currentColor"
          stroke-width="1.4"
        />
      </svg>
      <span>{LOCALES[id]}</span>
    </a>
  {/each}
{/if}

<style>
  .switch {
    display: inline-flex;
    align-items: center;
    gap: 0.35rem;
    height: 40px;
    padding: 0 0.6rem;
    border-radius: var(--radius);
    color: var(--ink-2);
    font-size: 0.88rem;
    font-weight: 500;
    white-space: nowrap;
  }

  .switch:hover {
    background: var(--wash);
    color: var(--ink);
    text-decoration: none;
  }
</style>
