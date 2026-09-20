<script lang="ts">
  import { base } from '$app/paths';
  import CodeBlock from '$lib/components/CodeBlock.svelte';
  import type { PageProps } from './$types';

  const { data }: PageProps = $props();
</script>

<svelte:head>
  <title>Getting started · @office-kit/docx</title>
</svelte:head>

<h1>Getting started</h1>

<p class="lede">
  <code>@office-kit/docx</code> is a library of plain functions. Every operation is a standalone
  export that takes a <code>Docx</code> value as its first argument, and nothing is hidden behind a
  class instance. That keeps the package tree-shakeable and makes it work the same way in Node and
  in the browser.
</p>

<h2>Install</h2>

<pre><code>pnpm add @office-kit/docx @office-kit/docx-preview</code></pre>

<p>
  <code>@office-kit/docx</code> is the authoring API. <code>@office-kit/docx-preview</code> is an
  optional companion that mounts a read-only preview of any <code>Docx</code> value into a DOM
  container. Both ship as ESM with bundled <code>.d.ts</code> types, and neither depends on a Node
  built-in.
</p>

<h2>Build a document from scratch</h2>

<p>
  <code>createDocx</code> hands back a plain <code>Docx</code> object that the rest of the API
  treats as a value. Append to it, then call <code>toUint8Array</code> for the bytes of the file,
  or <code>toBlob</code> in the browser.
</p>

<CodeBlock html={data.fromScratch.html} title={data.fromScratch.path} />

<h2>Open a template and fill placeholders</h2>

<p>
  An existing <code>.docx</code> can be opened with <code>openDocx</code> and edited in place.
  Every XML element the library does not model is kept as a pass-through node at its original
  position, so saving a template you did not change does not rewrite it into something else.
</p>

<p>
  <code>replaceTextEverywhere</code> walks the body, headers, footers, footnotes, endnotes, and
  comments, not just the main document. A match such as <code>{'{{name}}'}</code> that Word split
  across several runs of one paragraph is still found; a match that spans two paragraphs is not.
</p>

<CodeBlock html={data.templateFill.html} title={data.templateFill.path} />

<h2>Render in the browser</h2>

<p>
  <code>@office-kit/docx-preview</code> mounts a read-only preview of a <code>Docx</code>, or of
  raw bytes, into a DOM container. It wraps the open-source <code>docx-preview</code> renderer
  behind one function, <code>previewToDOM</code>.
</p>

<p>
  Try it in the <a href="{base}/playground">playground</a>, or read the
  <a href="{base}/docs/recipes">recipes</a> for the embedding pattern.
</p>

<h2>Where to go next</h2>

<ul>
  <li><a href="{base}/docs/recipes">Recipes</a>: common scenarios you can copy.</li>
  <li><a href="{base}/api">API reference</a>: every public export, grouped by area.</li>
  <li><a href="{base}/playground">Playground</a>: drop in a .docx and see the preview.</li>
  <li><a href="https://github.com/office-kit/docx">GitHub</a>: source and issues.</li>
</ul>

<style>
  .lede {
    color: var(--ink-2);
    font-size: 1.08rem;
  }
</style>
