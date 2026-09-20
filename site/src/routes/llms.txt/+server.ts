// /llms.txt — a short, machine-readable index of this site's documentation
// for LLM consumers. Format follows the llmstxt.org proposal: H1 title,
// blockquote summary, then sections with linked bullets.
//
// URLs are relative so the file works under any base path (custom domain,
// user page, or the /docx/ project page).

import type { RequestHandler } from "./$types";

export const prerender = true;

const BODY = `# @office-kit/docx

> OOXML-compliant (ECMA-376) Word \`.docx\` library. Reads, edits, and writes WordprocessingML in Node (tested on 22 and 24) and modern browsers. Elements and parts the library doesn't yet model are passed through unchanged on save, so hand-designed templates survive an edit. Function-first API: every operation is a standalone, tree-shakeable export. Optional companion \`@office-kit/docx-preview\` mounts a read-only render of any \`Docx\` into a DOM container.

This file is a short index. For every page concatenated into a single Markdown document — usable as a one-fetch LLM payload — see [\`/llms-full.txt\`](./llms-full.txt).

## Packages

- [\`@office-kit/docx\`](./api) — the public authoring API. \`Docx\` is a plain interface; every operation is a standalone function. Minimal slice (\`createDocx + appendParagraph + toUint8Array\`) bundles to ~42 KB minified; full surface ~133 KB.
- [\`@office-kit/docx-preview\`](./api) — browser-side read-only preview. Single function entry \`previewToDOM(source, container, options?)\`. Wraps the OSS \`docx-preview\` renderer.

## Docs

- [Getting started](./docs/getting-started) — install, build a doc from scratch, open and fill a template, embed the preview.
- [Recipes](./docs/recipes) — type-checked snippets for mail-merge across every part, starting from a designed base, tracked changes, and the browser embed.
- [API reference](./api) — every public export grouped by area: lifecycle, paragraphs & blocks, inline & text, styles & numbering, tables, images, headers/footers/sections, comments/notes/hyperlinks/bookmarks, fields & tracked changes, document properties, diagnostics, browser preview.
- [Playground](./playground) — drop a \`.docx\` (or load the built-in sample) to open it, validate it, re-save it, and see the re-saved copy rendered by \`@office-kit/docx-preview\`. The bytes never leave the page.

## Scope

In scope: WordprocessingML read/edit/write, browser preview via the companion package, OPC packaging and DrawingML as underlying layers.

Out of scope: \`.pptx\` (PresentationML) and \`.xlsx\` (SpreadsheetML) — see the sibling libraries [@office-kit/pptx](https://office-kit.github.io/pptx/) and [@office-kit/xlsx](https://office-kit.github.io/xlsx/).

Out of scope (permanent): rendering to PDF, headless Word automation, binary \`.doc\` (pre-2007 Word).

## Status

Pre-1.0. Core (\`openDocx\` / \`createDocx\` / \`toUint8Array\`) is stable. The test suite opens, re-saves, and re-reads the mammoth.js and python-docx fixture corpora and compares paragraphs and text; it runs on Node 22 and 24 on every change, and a tree-shake budget gate prevents accidental bundle bloat.

## Source

- [GitHub repository](https://github.com/office-kit/docx)
- [Office Kit family](https://office-kit.github.io/)
`;

export const GET: RequestHandler = () => {
  return new Response(BODY, {
    headers: {
      "Content-Type": "text/plain; charset=utf-8",
      "Cache-Control": "public, max-age=300",
    },
  });
};
