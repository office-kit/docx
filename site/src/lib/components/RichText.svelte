<script lang="ts">
  import { href } from '$lib/i18n';

  type Props = {
    /**
     * A translated sentence. `` `code` `` becomes inline code and `[label](/path)`
     * a link; a site path is localized, an `https:` URL is used as written.
     * Messages carry this much markup so translators can move code and links
     * to wherever their grammar puts them.
     */
    text: string;
  };

  const { text }: Props = $props();

  type Piece =
    | { kind: 'text'; value: string }
    | { kind: 'code'; value: string }
    | { kind: 'link'; value: string; to: string };

  const MARKUP = /`([^`]+)`|\[([^\]]+)\]\(([^)\s]+)\)/g;

  const pieces = $derived.by(() => {
    const out: Piece[] = [];
    let last = 0;
    for (const m of text.matchAll(MARKUP)) {
      if (m.index > last) out.push({ kind: 'text', value: text.slice(last, m.index) });
      const [whole, code, label, to] = m;
      if (code !== undefined) out.push({ kind: 'code', value: code });
      else if (label !== undefined && to !== undefined) out.push({ kind: 'link', value: label, to });
      last = m.index + whole.length;
    }
    if (last < text.length) out.push({ kind: 'text', value: text.slice(last) });
    return out;
  });
</script>

{#each pieces as piece, i (i)}{#if piece.kind === 'code'}<code>{piece.value}</code>{:else if piece.kind === 'link'}<a
      href={piece.to.startsWith('/') ? href(piece.to) : piece.to}>{piece.value}</a
    >{:else}{piece.value}{/if}{/each}
