<script lang="ts">
  import { onDestroy } from 'svelte';
  import type { Handle } from '@office-kit/docx-preview';

  type Props = {
    /** The saved .docx to draw; `null` clears the pane. */
    bytes: Uint8Array | null;
    /** Dims the pane: what it shows no longer matches what the visitor is editing. */
    stale?: boolean;
    /** CSS max-height of the scroll region; unset, it fills its parent. */
    maxHeight?: string;
    onerror?: (err: unknown) => void;
  };

  const { bytes, stale = false, maxHeight, onerror }: Props = $props();

  type Rendered = { element: HTMLDivElement; handle: Handle; naturalWidth: number };

  let host = $state<HTMLDivElement | undefined>();
  let staging = $state<HTMLDivElement | undefined>();
  let current: Rendered | undefined;
  // Renders overlap while someone types; only the newest one may reach the screen.
  let generation = 0;

  function discard(rendered: Rendered | undefined): void {
    rendered?.handle.dispose();
    rendered?.element.remove();
  }

  // docx-preview lays pages out at their real size (A4 is 794 CSS pixels
  // wide), so a narrow pane would scroll sideways. Scale the pages down to
  // the pane instead, and never up.
  function fit(): void {
    if (!host || !current) return;
    const scale = Math.min(1, host.clientWidth / current.naturalWidth);
    current.element.style.zoom = String(scale);
  }

  async function render(
    next: Uint8Array,
    into: HTMLDivElement,
    offstage: HTMLDivElement,
  ): Promise<void> {
    const mine = ++generation;
    const { previewToDOM } = await import('@office-kit/docx-preview');

    // The new pages are laid out invisibly and swapped in when they are
    // complete, so the pane never flashes empty between runs.
    const element = document.createElement('div');
    element.className = 'preview-pages';
    offstage.append(element);
    let handle: Handle;
    try {
      handle = await previewToDOM(next, element);
    } catch (err) {
      element.remove();
      if (mine === generation) onerror?.(err);
      return;
    }
    const rendered = { element, handle, naturalWidth: element.offsetWidth };
    if (mine !== generation) {
      discard(rendered);
      return;
    }
    discard(current);
    current = rendered;
    into.append(element);
    fit();
  }

  $effect(() => {
    if (!host || !staging) return;
    if (bytes) {
      void render(bytes, host, staging);
    } else {
      generation++;
      discard(current);
      current = undefined;
    }
  });

  $effect(() => {
    if (!host) return;
    const observer = new ResizeObserver(fit);
    observer.observe(host);
    return () => observer.disconnect();
  });

  onDestroy(() => {
    generation++;
    discard(current);
  });
</script>

<!-- A scrollable region must be focusable, or keyboard users cannot scroll it. -->
<!-- svelte-ignore a11y_no_noninteractive_tabindex -->
<div
  class="preview"
  class:stale
  style:max-height={maxHeight}
  tabindex="0"
  role="region"
  aria-label="Rendered document"
  bind:this={host}
>
  <div class="staging" bind:this={staging}></div>
</div>

<style>
  /* docx-preview paints its own grey desk and white sheets. Those are the
   * document's colours, not this site's, so they do not follow the dark theme;
   * the pane matches the desk so no seam shows around it. */
  .preview {
    position: relative;
    height: 100%;
    min-height: 12rem;
    overflow: auto;
    background: #808080;
  }

  .preview :global(.preview-pages) {
    width: max-content;
    transition: opacity 120ms ease;
  }

  /* On stage the pages fill the pane, so the desk has no ragged right edge. */
  .preview > :global(.preview-pages) {
    min-width: 100%;
  }

  .preview.stale :global(.preview-pages) {
    opacity: 0.45;
  }

  /* Pages are laid out here at their own width, which is what `fit` scales
   * against. The box clips them so a render in progress never gives the pane
   * a scrollbar. */
  .staging {
    position: absolute;
    inset: 0;
    overflow: hidden;
    visibility: hidden;
    pointer-events: none;
  }
</style>
