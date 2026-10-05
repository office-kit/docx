// Registry of every example file. The ?raw imports give us the on-disk
// source verbatim. The .ts files themselves are in the project's tsconfig
// include path, so svelte-check type-checks them against the live
// @office-kit/docx / @office-kit/docx-preview surface — an API rename breaks the
// docs build before anything ships, even though the modules are never
// evaluated at runtime.

import fromScratch from "./from-scratch.ts?raw";
import templateFill from "./template-fill.ts?raw";
import previewEmbed from "./preview-embed.ts?raw";
import recipeMailMerge from "./recipe-mail-merge.ts?raw";
import recipeStyledBase from "./recipe-styled-base.ts?raw";
import recipeTrackedChanges from "./recipe-tracked-changes.ts?raw";
import type { Translations } from "$lib/i18n/define";

/** The prose around a snippet; the code itself is the same in every locale. */
export type ExampleText = {
  /** Human title for the snippet (shown above the code block). */
  title: string;
  /** Short description used in docs. */
  description: string;
  /** Where the recipe can be seen end to end: a `pnpm sample` output or a live page. */
  seeAlso?: string;
};

// The English text sits on the example itself because `/llms-full.txt`, which
// is English only, reads it from there too.
export type Example = ExampleText & {
  /** Repo-relative path, also used for the file-tab caption. */
  path: string;
  /** Verbatim source text. */
  source: string;
  translations: Translations<ExampleText>;
};

export const examples = {
  fromScratch: {
    title: "Build a document from scratch",
    path: "site/src/lib/examples/from-scratch.ts",
    description:
      "Heading, body paragraph, bullet list, table, A4 page size — the canonical authoring slice.",
    source: fromScratch,
    translations: {
      ja: {
        title: "ドキュメントをゼロから作成する",
        description:
          "見出し、本文の段落、箇条書き、表、A4 の用紙サイズ。ドキュメント作成の基本をひととおり押さえた例です。",
      },
    },
  },
  templateFill: {
    title: "Open a template, fill placeholders",
    path: "site/src/lib/examples/template-fill.ts",
    description:
      "replaceTextEverywhere walks the body, headers, footers, footnotes, endnotes, and comments — not just the main document.",
    source: templateFill,
    translations: {
      ja: {
        title: "テンプレートを開いてプレースホルダーを埋める",
        description:
          "replaceTextEverywhere は本文だけでなく、ヘッダー、フッター、脚注、文末脚注、コメントも走査します。",
      },
    },
  },
  previewEmbed: {
    title: "Render in the browser with @office-kit/docx-preview",
    path: "site/src/lib/examples/preview-embed.ts",
    description:
      "previewToDOM mounts a read-only render of any Docx into a DOM container. Returns an idempotent dispose handle.",
    source: previewEmbed,
    seeAlso: "The playground on this site runs this code.",
    translations: {
      ja: {
        title: "@office-kit/docx-preview でブラウザーに表示する",
        description:
          "previewToDOM は、任意の Docx の読み取り専用レンダリングを DOM コンテナにマウントします。戻り値の破棄用ハンドルは、何度呼び出しても安全です。",
        seeAlso: "このサイトのプレイグラウンドでは、このコードが実際に動いています。",
      },
    },
  },
  recipeMailMerge: {
    title: "Mail-merge across every part of the package",
    path: "site/src/lib/examples/recipe-mail-merge.ts",
    description:
      "replaceTextEverywhere walks the body, headers, footers, footnotes, endnotes, and comments, so a placeholder in a header is filled the same way one in the body is.",
    source: recipeMailMerge,
    seeAlso: "pnpm sample writes samples/21-mailmerge-cross-part-template.docx and -filled.docx.",
    translations: {
      ja: {
        title: "パッケージ内のすべてのパーツを対象に差し込み印刷を行う",
        description:
          "replaceTextEverywhere は本文、ヘッダー、フッター、脚注、文末脚注、コメントを走査するため、ヘッダー内のプレースホルダーも本文と同じように埋められます。",
        seeAlso:
          "pnpm sample を実行すると、samples/21-mailmerge-cross-part-template.docx と -filled.docx が出力されます。",
      },
    },
  },
  recipeStyledBase: {
    title: "Start from a designed base",
    path: "site/src/lib/examples/recipe-styled-base.ts",
    description:
      "Open a hand-designed .docx, lift its style table into your authoring graph with mergeStylesFromTemplate, then keep appending. Custom styles applied by name via setParagraphStyle.",
    source: recipeStyledBase,
    seeAlso: "pnpm sample writes samples/30-styled-base-template.docx and -filled.docx.",
    translations: {
      ja: {
        title: "デザイン済みのベースから始める",
        description:
          "手作業でデザインした .docx を開き、mergeStylesFromTemplate でそのスタイル定義を作成中のドキュメントに取り込んでから、内容を追加していきます。カスタムスタイルは setParagraphStyle で名前を指定して適用します。",
        seeAlso:
          "pnpm sample を実行すると、samples/30-styled-base-template.docx と -filled.docx が出力されます。",
      },
    },
  },
  recipeTrackedChanges: {
    title: "Accept or reject tracked changes in bulk",
    path: "site/src/lib/examples/recipe-tracked-changes.ts",
    description:
      "acceptAllRevisions inlines suggested inserts and bakes in deletions. rejectAllRevisions does the inverse — useful for producing both 'final' and 'original' branches from one reviewed doc.",
    source: recipeTrackedChanges,
    seeAlso: "pnpm sample writes samples/09-tracked-changes.docx.",
    translations: {
      ja: {
        title: "変更履歴を一括で承諾・却下する",
        description:
          "acceptAllRevisions は提案された挿入を本文に取り込み、削除を確定します。rejectAllRevisions はその逆を行います。校閲済みの 1 つの文書から「最終版」と「元の版」の両方を作りたいときに便利です。",
        seeAlso: "pnpm sample を実行すると、samples/09-tracked-changes.docx が出力されます。",
      },
    },
  },
} as const satisfies Record<string, Example>;

export type ExampleKey = keyof typeof examples;
