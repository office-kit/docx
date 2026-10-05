// The two interactive pages: the playground and the REPL.

import { defineMessages } from "../define.js";

export default defineMessages({
  en: {
    playground: {
      title: "Playground",
      heading: "Open a .docx in your browser",
      lede: "Drop a file and this page opens it with `@office-kit/docx`, counts what is inside, validates the package, saves it again, and renders the saved copy with `@office-kit/docx-preview`. Nothing is uploaded: the whole pipeline runs in this tab.",
      loading: "Loading the library…",
      opening: (name: string) => `Opening ${name}…`,
      opened: (name: string, read: string, saved: string) =>
        `Opened ${name} (${read} bytes) and saved it again (${saved} bytes). The preview shows the saved copy.`,
      failed: (message: string) => `This file could not be opened: ${message}`,
      chooseLabel: "Choose a .docx file",
      dropHere: "Drop a .docx file here",
      choose: "Choose a file",
      sample: "Load the sample document",
      downloadSaved: "Download the re-saved file",
      caveat:
        "The preview is drawn by the open-source docx-preview renderer, which approximates Word’s layout: pagination, fonts, and floating objects can differ. Word and LibreOffice remain the exact renderers.",
      stats: {
        paragraphs: "Paragraphs",
        headings: "Headings",
        tables: "Tables",
        images: "Images",
        words: "Words",
        comments: "Comments",
        notes: "Footnotes and endnotes",
        title: "Title",
        author: "Author",
        notSet: "Not set",
      },
      validation: "Validation",
      preview: "Preview",
    },
    repl: {
      title: "REPL",
      lede: "Write code and the document redraws as you type. Every public function is already in scope, and `doc` is a new, empty document from `createDocx`. The preview is drawn from the bytes `toUint8Array` wrote, and the download is those same bytes. Nothing is uploaded: the code, the library, and the preview all run in this tab.",
      code: "Code",
      reset: "Reset",
      copy: "Copy",
      preview: "Preview",
      building: "Building…",
      download: "Download .docx",
      atLine: (line: number, message: string) => `Line ${line}: ${message}`,
      previewFailed: (message: string) => `The preview could not draw this document: ${message}`,
    },
  },
  ja: {
    playground: {
      title: "プレイグラウンド",
      heading: "ブラウザーで .docx を開く",
      lede: "ファイルをドロップすると、このページが `@office-kit/docx` でファイルを開いて中身を集計し、パッケージを検証してから保存し直し、保存したファイルを `@office-kit/docx-preview` で描画します。処理はすべてこのタブ内で完結するため、ファイルがアップロードされることはありません。",
      loading: "ライブラリを読み込んでいます…",
      opening: (name: string) => `${name} を開いています…`,
      opened: (name: string, read: string, saved: string) =>
        `${name}（${read} バイト）を開き、保存し直しました（${saved} バイト）。プレビューには保存し直したファイルを表示しています。`,
      failed: (message: string) => `このファイルを開けませんでした: ${message}`,
      chooseLabel: ".docx ファイルを選択",
      dropHere: "ここに .docx ファイルをドロップ",
      choose: "ファイルを選択",
      sample: "サンプル文書を読み込む",
      downloadSaved: "保存し直したファイルをダウンロード",
      caveat:
        "プレビューはオープンソースの docx-preview レンダラーで描画しており、Word のレイアウトを近似したものです。改ページの位置、フォント、フローティング オブジェクトの表示が異なる場合があります。正確な表示は Word や LibreOffice で確認してください。",
      stats: {
        paragraphs: "段落",
        headings: "見出し",
        tables: "表",
        images: "画像",
        words: "単語数",
        comments: "コメント",
        notes: "脚注と文末脚注",
        title: "タイトル",
        author: "作成者",
        notSet: "未設定",
      },
      validation: "検証",
      preview: "プレビュー",
    },
    repl: {
      title: "REPL",
      lede: "コードを書くと、入力に合わせてドキュメントが再描画されます。すべての公開関数はあらかじめスコープに入っており、`doc` は `createDocx` で作成した新しい空のドキュメントです。プレビューは `toUint8Array` が書き出したバイト列から描画しており、ダウンロードされるのも同じバイト列です。コード、ライブラリ、プレビューはすべてこのタブ内で動作するため、何もアップロードされません。",
      code: "コード",
      reset: "リセット",
      copy: "コピー",
      preview: "プレビュー",
      building: "生成しています…",
      download: ".docx をダウンロード",
      atLine: (line: number, message: string) => `${line} 行目: ${message}`,
      previewFailed: (message: string) =>
        `このドキュメントをプレビューで描画できませんでした: ${message}`,
    },
  },
});
