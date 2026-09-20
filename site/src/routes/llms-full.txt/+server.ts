// /llms-full.txt — every doc page concatenated into a single Markdown
// document. Companion of /llms.txt (which is just an index): this is the
// payload, intended for LLM ingestion in one fetch.
//
// The docs pages are hand-written `.svelte` rather than Markdown, so there
// is no page source to concatenate. The body is assembled from the two
// sources of truth the pages themselves render from:
//
//   * `$lib/examples` — the type-checked example files used by the home
//     page and the recipes page.
//   * `$lib/api-groups` — the curated grouped-export list used by the
//     `/api` page and the `check:api-page` CI gate.
//
// That keeps the LLM payload aligned with what the site actually
// renders, with no second copy of the content to drift.

import { examples } from "$lib/examples";
import { apiGroups, apiTotalCount } from "$lib/api-groups";
import type { RequestHandler } from "./$types";

export const prerender = true;

const PREAMBLE = `# @office-kit/docx — full documentation

This file is the concatenation of every page on the @office-kit/docx docs site, intended for LLM ingestion in a single fetch. Page boundaries are marked with H1 headings prefixed by the source path. The companion index at \`/llms.txt\` lists the same pages with one-line descriptions.

Source repo: https://github.com/office-kit/docx
`;

const PROJECT_OVERVIEW = `OOXML-compliant (ECMA-376 Part 1 — WordprocessingML) Word \`.docx\` library for Node (tested on 22 and 24) and modern browsers. Function-first API: every operation is a standalone, tree-shakeable export, never a method on a class. Elements and parts the library does not yet model are passed through unchanged on save, so hand-designed templates survive an edit.

Two packages on npm:

- \`@office-kit/docx\` — the public authoring API.
- \`@office-kit/docx-preview\` — optional companion that mounts a read-only preview of any \`Docx\` (or raw bytes) into a DOM container, by wrapping the OSS \`docx-preview\` renderer.

Bundle sizes: minimal \`createDocx + appendParagraph + toUint8Array\` slice ~42 KB minified (CI fails above 50 KB); full surface ~133 KB.

Test gate: the vitest suite runs on Node 22 and 24 on every change, and opens, re-saves, and re-reads the mammoth.js and python-docx fixture corpora, comparing paragraphs and text.
`;

const GETTING_STARTED = `Install both packages:

\`\`\`sh
pnpm add @office-kit/docx @office-kit/docx-preview
\`\`\`

Both ship as ESM with bundled \`.d.ts\` types. Neither has Node-only dependencies.

## Build a document from scratch

The hello-world. \`createDocx\` returns a plain \`Docx\` object that the rest of the API treats as a value.

\`\`\`ts title="${examples.fromScratch.path}"
${examples.fromScratch.source.trimEnd()}
\`\`\`

## Open a template, fill placeholders

Existing \`.docx\` files are opened with \`openDocx\` and edited in place. Every XML element the library does not yet model is kept as a pass-through node at its original position, so saving a template you did not change does not rewrite it into something else.

\`replaceTextEverywhere\` walks the body, headers, footers, footnotes, endnotes, and comments — not just the main document. A match like \`{{name}}\` that is split across several runs of one paragraph is still found; a match that spans two paragraphs is not.

\`\`\`ts title="${examples.templateFill.path}"
${examples.templateFill.source.trimEnd()}
\`\`\`

## Render in the browser

The companion package \`@office-kit/docx-preview\` mounts a read-only preview of any \`Docx\` into a DOM container. It wraps the OSS \`docx-preview\` renderer behind a stable function-API entry point.

\`\`\`ts title="${examples.previewEmbed.path}"
${examples.previewEmbed.source.trimEnd()}
\`\`\`
`;

function buildRecipesMarkdown(): string {
  const intro = `Type-checked snippets for common scenarios. Every snippet below lives under \`site/src/lib/examples/\` and is type-checked by \`svelte-check\` against the live \`@office-kit/docx\` / \`@office-kit/docx-preview\` surface — an API rename breaks the docs build before anything ships.

A note under each snippet says where the recipe can be seen end to end.
`;

  const recipeKeys = [
    "recipeMailMerge",
    "recipeStyledBase",
    "recipeTrackedChanges",
    "previewEmbed",
  ] as const;

  const sections = recipeKeys.map((key) => {
    const ex = examples[key];
    return [
      `### ${ex.title}`,
      "",
      ex.description,
      "",
      ex.seeAlso,
      "",
      `\`\`\`ts title="${ex.path}"`,
      ex.source.trimEnd(),
      "```",
    ].join("\n");
  });

  return `${intro}\n${sections.join("\n\n")}\n`;
}

function buildApiMarkdown(): string {
  const intro = `\`@office-kit/docx\` exposes ${apiTotalCount} standalone functions and constants, grouped by area below. \`previewToDOM\` lives in the companion \`@office-kit/docx-preview\` package; everything else lives in \`@office-kit/docx\`. For full signatures, read \`src/index.ts\` in the repo.
`;

  const sections = apiGroups.map((g) => {
    const items = g.entries.map((e) => {
      const sig = e.sig ? `: \`${e.sig}\`` : "";
      return `- \`${e.name}\`${sig}`;
    });
    return `### ${g.title}\n\n${g.description}\n\n${items.join("\n")}`;
  });

  return `${intro}\n${sections.join("\n\n")}\n`;
}

const PLAYGROUND = `An interactive page at \`/playground\`. Drop a \`.docx\` onto the page (or load the built-in sample, which is built in the browser via \`@office-kit/docx\`). The page opens it with \`openDocx\`, shows \`statistics\` and \`validate\` results, saves it again with \`toUint8Array\`, and renders the saved copy with \`@office-kit/docx-preview\`. The bytes never leave the browser.

The "Download the re-saved file" button hands back the round-tripped document.
`;

const SCOPE = `## In scope

WordprocessingML (\`.docx\`) — read, edit, write. OPC packaging (ECMA-376 Part 2) and DrawingML are part of the OOXML stack and are bundled into the package, but exist to serve docx. Browser preview lives in the companion \`@office-kit/docx-preview\`.

## Out of scope

\`.pptx\` (PresentationML) and \`.xlsx\` (SpreadsheetML). See the sibling libraries [@office-kit/pptx](https://office-kit.github.io/pptx/) and [@office-kit/xlsx](https://office-kit.github.io/xlsx/).

## Out of scope (permanent)

Rendering to PDF, headless Word automation, binary \`.doc\` (pre-2007 Word).
`;

function buildBody(): string {
  const parts = [
    PREAMBLE,
    "\n---\n\n<!-- Page: / -->\n# @office-kit/docx — overview\n\n" + PROJECT_OVERVIEW,
    "\n---\n\n<!-- Page: /docs/getting-started -->\n# Getting started\n\n" + GETTING_STARTED,
    "\n---\n\n<!-- Page: /docs/recipes -->\n# Recipes\n\n" + buildRecipesMarkdown(),
    "\n---\n\n<!-- Page: /api -->\n# API reference\n\n" + buildApiMarkdown(),
    "\n---\n\n<!-- Page: /playground -->\n# Playground\n\n" + PLAYGROUND,
    "\n---\n\n<!-- Page: scope -->\n# Scope\n\n" + SCOPE,
  ];
  return parts.join("\n");
}

export const GET: RequestHandler = () => {
  return new Response(buildBody(), {
    headers: {
      "Content-Type": "text/plain; charset=utf-8",
      "Cache-Control": "public, max-age=300",
    },
  });
};
