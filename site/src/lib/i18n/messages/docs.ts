// The /docs section: its layout, Getting started, and Recipes. Recipe titles
// and descriptions live with the examples in `$lib/examples`.

import { defineMessages } from "../define.js";

export default defineMessages({
  en: {
    layout: {
      documentation: "Documentation",
      pager: "Previous and next page",
      previous: "Previous",
      next: "Next",
      llms: "Models and tools can fetch [/llms-full.txt](/llms-full.txt) for these docs as one Markdown file, or [/llms.txt](/llms.txt) for the index.",
    },
    gettingStarted: {
      title: "Getting started",
      lede: "`@office-kit/docx` is a library of plain functions. Every operation is a standalone export that takes a `Docx` value as its first argument, and nothing is hidden behind a class instance. That keeps the package tree-shakeable and makes it work the same way in Node and in the browser.",
      install: "Install",
      installBody:
        "`@office-kit/docx` is the authoring API. `@office-kit/docx-preview` is an optional companion that mounts a read-only preview of any `Docx` value into a DOM container. Both ship as ESM with bundled `.d.ts` types, and neither depends on a Node built-in.",
      fromScratch: "Build a document from scratch",
      fromScratchBody:
        "`createDocx` hands back a plain `Docx` object that the rest of the API treats as a value. Append to it, then call `toUint8Array` for the bytes of the file, or `toBlob` in the browser.",
      template: "Open a template and fill placeholders",
      templateBody: [
        "An existing `.docx` can be opened with `openDocx` and edited in place. Every XML element the library does not model is kept as a pass-through node at its original position, so saving a template you did not change does not rewrite it into something else.",
        "`replaceTextEverywhere` walks the body, headers, footers, footnotes, endnotes, and comments, not just the main document. A match such as `{{name}}` that Word split across several runs of one paragraph is still found; a match that spans two paragraphs is not.",
      ],
      browser: "Render in the browser",
      browserBody: [
        "`@office-kit/docx-preview` mounts a read-only preview of a `Docx`, or of raw bytes, into a DOM container. It wraps the open-source `docx-preview` renderer behind one function, `previewToDOM`.",
        "Try it in the [playground](/playground), or read the [recipes](/docs/recipes) for the embedding pattern.",
      ],
      next: "Where to go next",
      nextLinks: [
        "[Recipes](/docs/recipes): common scenarios you can copy.",
        "[API reference](/api): every public export, grouped by area.",
        "[Playground](/playground): drop in a .docx and see the preview.",
        "[REPL](/repl): edit code and watch the document it builds.",
        "[GitHub](https://github.com/office-kit/docx): source and issues.",
      ],
    },
    recipes: {
      title: "Recipes",
      lede: "Working code for common tasks. Every snippet is a real file under `site/src/lib/examples/` that is type-checked against the library on every build, so an API rename breaks this page before it ships. For a specific function, see the [API reference](/api).",
      jump: "Recipes on this page",
      more: "To build a document end to end, walk through [Getting started](/docs/getting-started), then open [the playground](/playground) to see the rendered output.",
    },
  },
  ja: {
    layout: {
      documentation: "ドキュメント",
      pager: "前後のページ",
      previous: "前へ",
      next: "次へ",
      llms: "AI モデルやツール向けに、このドキュメントを 1 つの Markdown ファイルにまとめた [/llms-full.txt](/llms-full.txt) と、目次にあたる [/llms.txt](/llms.txt) を用意しています（いずれも英語）。",
    },
    gettingStarted: {
      title: "はじめに",
      lede: "`@office-kit/docx` はシンプルな関数で構成されたライブラリです。すべての操作は、第 1 引数に `Docx` 値を受け取る独立したエクスポートとして提供されており、クラスのインスタンスに隠れた状態はありません。そのためパッケージはツリーシェイキングが効き、Node でもブラウザーでも同じように動作します。",
      install: "インストール",
      installBody:
        "`@office-kit/docx` がドキュメントを作成・編集するための API です。`@office-kit/docx-preview` は必要に応じて導入する姉妹パッケージで、任意の `Docx` 値の読み取り専用プレビューを DOM コンテナにマウントします。どちらも `.d.ts` 型定義を同梱した ESM として配布されており、Node の組み込みモジュールには依存していません。",
      fromScratch: "ドキュメントをゼロから作成する",
      fromScratchBody:
        "`createDocx` はプレーンな `Docx` オブジェクトを返し、ほかの API はこれを値として扱います。内容を追加したら `toUint8Array` を呼び出してファイルのバイト列を取得します。ブラウザーでは `toBlob` も使えます。",
      template: "テンプレートを開いてプレースホルダーを埋める",
      templateBody: [
        "既存の `.docx` は `openDocx` で開いて、そのまま編集できます。ライブラリがモデル化していない XML 要素は、すべて元の位置にパススルーノードとして保持されます。そのため、手を加えていないテンプレートを保存しても、別の内容に書き換わることはありません。",
        "`replaceTextEverywhere` は本文だけでなく、ヘッダー、フッター、脚注、文末脚注、コメントも走査します。Word が 1 つの段落内で複数のランに分割した `{{name}}` のような文字列も検出できますが、2 つの段落にまたがる文字列は検出されません。",
      ],
      browser: "ブラウザーで表示する",
      browserBody: [
        "`@office-kit/docx-preview` は、`Docx` またはバイト列の読み取り専用プレビューを DOM コンテナにマウントします。オープンソースの `docx-preview` レンダラーを `previewToDOM` という 1 つの関数でラップしたものです。",
        "[プレイグラウンド](/playground)で試せるほか、埋め込み方は[レシピ](/docs/recipes)で紹介しています。",
      ],
      next: "次のステップ",
      nextLinks: [
        "[レシピ](/docs/recipes)：そのまま使える、よくあるシナリオのコード集。",
        "[API リファレンス](/api)：すべての公開エクスポートを分野別に掲載。",
        "[プレイグラウンド](/playground)：.docx をドロップしてプレビューを確認。",
        "[REPL](/repl)：コードを編集しながら、生成されるドキュメントを確認。",
        "[GitHub](https://github.com/office-kit/docx)：ソースコードと Issue。",
      ],
    },
    recipes: {
      title: "レシピ",
      lede: "よくある作業のための、動作するコード集です。各スニペットは `site/src/lib/examples/` 以下にある実際のファイルで、ビルドのたびにライブラリに対して型チェックされます。そのため API の名前が変わると、公開される前にこのページのビルドが失敗します。個々の関数については [API リファレンス](/api)を参照してください。",
      jump: "このページのレシピ",
      more: "ドキュメントを作成する一連の流れは[はじめに](/docs/getting-started)で確認し、描画結果は[プレイグラウンド](/playground)で確認してください。",
    },
  },
});
