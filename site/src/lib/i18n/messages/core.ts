// The shell every page shares: header, footer, search, and the small controls
// several pages use.

import { defineMessages } from "../define.js";

export default defineMessages({
  en: {
    seo: {
      title: "@office-kit/docx: read, edit, and write Word documents in TypeScript",
      description:
        "Build a .docx from nothing or open one that already exists and change it. Paragraphs, lists, tables, images, headers, comments, and tracked changes are plain typed functions. Runs in Node and the browser.",
    },
    skip: "Skip to content",
    nav: {
      docs: "Docs",
      recipes: "Recipes",
      api: "API",
      playground: "Playground",
      repl: "REPL",
      editor: "Editor",
    },
    language: {
      switchTo: "Read this page in",
    },
    search: {
      trigger: "Search docs (press / or Cmd-K)",
      label: "Search",
      placeholder: "Search docs and API…",
      unavailable:
        "Search index not available. Run `pnpm --filter word-kit-site build` to generate it (the dev server skips indexing).",
      loading: "Loading search index…",
      noResults: (query: string) => `No results for "${query}".`,
      hint: "Type to search across docs, recipes, and the full API reference.",
    },
    copy: {
      label: "Copy code to clipboard",
      idle: "Copy",
      copied: "Copied",
      failed: "Failed",
    },
    validation: {
      clean: "`validate(doc)` found nothing wrong with the package.",
      level: { error: "error", warning: "warning" },
    },
    preview: {
      label: "Rendered document",
    },
  },
  ja: {
    seo: {
      title: "@office-kit/docx: TypeScript で Word 文書を読み書き・編集する",
      description:
        ".docx をゼロから作成することも、既存のファイルを開いて書き換えることもできます。段落、リスト、表、画像、ヘッダー、コメント、変更履歴を型付きの関数で扱えます。Node とブラウザーで動作します。",
    },
    skip: "本文へスキップ",
    nav: {
      docs: "ドキュメント",
      recipes: "レシピ",
      api: "API",
      playground: "プレイグラウンド",
      repl: "REPL",
      editor: "エディター",
    },
    language: {
      switchTo: "このページを次の言語で表示:",
    },
    search: {
      trigger: "ドキュメントを検索（/ または Cmd-K）",
      label: "検索",
      placeholder: "ドキュメントと API を検索…",
      unavailable:
        "検索インデックスがありません。`pnpm --filter word-kit-site build` を実行して生成してください（開発サーバーではインデックスを作成しません）。",
      loading: "検索インデックスを読み込んでいます…",
      noResults: (query: string) => `「${query}」に一致する結果はありません。`,
      hint: "キーワードを入力すると、ドキュメント、レシピ、API リファレンス全体を検索できます。",
    },
    copy: {
      label: "コードをクリップボードにコピー",
      idle: "コピー",
      copied: "コピーしました",
      failed: "失敗しました",
    },
    validation: {
      clean: "`validate(doc)` はパッケージに問題を検出しませんでした。",
      level: { error: "エラー", warning: "警告" },
    },
    preview: {
      label: "描画されたドキュメント",
    },
  },
});
