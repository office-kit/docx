// The landing page.

import { defineMessages } from "../define.js";

type Cell = { text: string; tone?: "yes" | "no" | "part" };

type Home = {
  title: string;
  hero: { heading: string; lede: string; start: string; playground: string };
  stage: {
    heading: string;
    codeLabel: string;
    sheetLabel: string;
    note: string;
    building: string;
    download: string;
    saved: (name: string, kilobytes: string) => string;
    failed: string;
  };
  paths: Array<{ title: string; body: string; code: string; href: string; link: string }>;
  template: { heading: string; body: string[] };
  proof: { heading: string; lede: string; items: Array<{ claim: string; how: string }> };
  compare: {
    heading: string;
    lede: string;
    capability: string;
    rows: Array<{ topic: string; ours: Cell; theirs: Cell }>;
    advice: string;
  };
  caps: {
    heading: string;
    lede: string;
    items: Array<{ area: string; items: string }>;
    notYet: string;
    notYetItems: string[];
  };
  agents: { heading: string; body: string; index: string; full: string };
  family: { heading: string; body: string };
};

export default defineMessages<Home>({
  en: {
    title: "@office-kit/docx: read, edit, and write Word files in TypeScript",
    hero: {
      heading: "Read, edit, and write Word documents in TypeScript",
      lede: "Build a .docx from nothing, or open one that already exists and change it. Paragraphs, lists, tables, images, headers, comments, and tracked changes are plain typed functions, and whatever the library does not model is written back untouched. It runs in Node and in the browser.",
      start: "Get started",
      playground: "Open the playground",
    },
    stage: {
      heading: "A document and the code that built it",
      codeLabel: "Source of hero-document.ts",
      sheetLabel: "The document that code produces",
      note: "That page is real output. The code shown with it built the document; this site saved it as .docx bytes, opened those bytes again with `openDocx`, and drew what it read as HTML. Download the file and open it in Word: the headings are styles, the list is a list, and the table is a table.",
      building: "Building the document…",
      download: "Download the .docx",
      saved: (name, kilobytes) => `Saved ${name} (${kilobytes} KB), built in this tab.`,
      failed: "The document could not be built in this browser. The console has the error.",
    },
    paths: [
      {
        title: "Start from an empty document",
        body: "createDocx() returns a document with its styles part and relationships already in place, so there is no template file to ship. Append headings, paragraphs, bullet and numbered lists, tables, inline images, headers, footers, footnotes, and a table of contents.",
        code: "createDocx()",
        href: "/repl",
        link: "Write one in the REPL",
      },
      {
        title: "Edit a file you already have",
        body: "Open a .docx, replace text across the body, headers, footers, footnotes, endnotes, and comments, swap an image by its alt text, accept or reject tracked insertions and deletions, and save. Whatever the library does not model goes back into the file as it came.",
        code: "openDocx(bytes)",
        href: "/docs/recipes",
        link: "Browse the recipes",
      },
      {
        title: "Show the document in the page",
        body: "The companion @office-kit/docx-preview package renders a Docx, or raw bytes, into a DOM element. It wraps the open-source docx-preview renderer behind one function. It runs in the browser, so the file is never uploaded.",
        code: "previewToDOM(doc, element)",
        href: "/playground",
        link: "Try it in the playground",
      },
    ],
    template: {
      heading: "Design the template in Word. Fill it in from code.",
      body: [
        "The letterhead, the contract, the report with the right fonts: most documents that matter already exist as a .docx somebody cares about. Open it, replace what changes, and everything else stays as it was designed.",
        "Word often splits a placeholder such as `{{name}}` across several runs. Matching works on the paragraph’s text rather than its XML, so it is still found, and `replaceTextEverywhere` reaches the headers, footers, notes, and comments as well as the body.",
      ],
    },
    proof: {
      heading: "Round trips, checked by machines",
      lede: "A library that edits other people’s documents has one job above all: do not break them. This is what backs that up.",
      items: [
        {
          claim: "Real Word files go in and come out the same.",
          how: "The test suite opens the .docx corpora that mammoth.js and python-docx test against, saves every file, reopens it, and compares the paragraphs and the text. CI runs it on Node 22 and Node 24 for every pull request.",
        },
        {
          claim: "What the library does not understand, it does not touch.",
          how: "Elements it does not model are kept as pass-through nodes at their original position, and parts it does not read stay in the package. A content control or a custom XML part survives an edit.",
        },
        {
          claim: "validate() looks for what makes Word offer a repair.",
          how: "Relationships that point at a missing part, comment, footnote and endnote references with no target, unpaired bookmarks, missing media. Twenty feature combinations are asserted clean in the test suite.",
        },
        {
          claim: "You ship only the functions you import.",
          how: "120 side-effect-free exports and no classes. createDocx, appendParagraph, and toUint8Array bundle to 42 KB minified; the whole API is 133 KB. CI fails if the small bundle passes 50 KB.",
        },
        {
          claim: "One runtime dependency.",
          how: "fflate, for ZIP. The XML parser and serializer are written here, and nothing in the library imports a Node built-in.",
        },
      ],
    },
    compare: {
      heading: "How it differs from the docx package",
      lede: "The `docx` package is the established way to generate a Word file in JavaScript, and it can author things this library cannot yet. The difference is direction: that package builds new documents, while this library also reads and edits the ones you already have. Compared against docx 9.7.1.",
      capability: "Capability",
      rows: [
        {
          topic: "Open and edit an existing .docx",
          ours: {
            text: "Yes. The file becomes a model: change any paragraph, run, or table",
            tone: "yes",
          },
          theirs: {
            text: "Placeholders only. patchDocument swaps {{tokens}} for new content",
            tone: "part",
          },
        },
        {
          topic: "Read back what is in a file",
          ours: {
            text: "Text, outline, statistics, images, fields, hyperlinks, bookmarks",
            tone: "yes",
          },
          theirs: { text: "No. patchDetector lists the placeholders and nothing else", tone: "no" },
        },
        {
          topic: "What a placeholder can become",
          ours: { text: "Text", tone: "part" },
          theirs: { text: "Text, paragraphs, tables, or images", tone: "yes" },
        },
        {
          topic: "Tracked changes",
          ours: {
            text: "Accepts or rejects the ones in a file; cannot author new ones",
            tone: "part",
          },
          theirs: {
            text: "Authors insertions and deletions; cannot resolve existing ones",
            tone: "part",
          },
        },
        {
          topic: "Merged cells, floating images, text boxes",
          ours: { text: "Not yet", tone: "no" },
          theirs: { text: "Yes", tone: "yes" },
        },
        {
          topic: "API shape",
          ours: { text: "Tree-shakeable functions over plain data, ESM only" },
          theirs: { text: "A tree of class instances; ESM, CommonJS, and script-tag bundles" },
        },
        {
          topic: "Runtime dependencies",
          ours: { text: "One: fflate" },
          theirs: { text: "Six, including jszip and xml-js" },
        },
      ],
      advice:
        "Pick the docx package if you only ever generate new documents and need merged cells, floating images, text boxes, or a CommonJS build. Pick this library if the document already exists and you need to read it, change it, or check it.",
    },
    caps: {
      heading: "What you can build today",
      lede: "The library is pre-1.0 and says so. This is what works now, and what does not yet. [The API reference](/api) lists every function.",
      items: [
        {
          area: "Text",
          items:
            "Paragraphs, headings, runs with bold, italic, underline, colour, highlight, size, and fonts. Alignment, indents, spacing, borders, shading.",
        },
        {
          area: "Lists and styles",
          items:
            "Bullet and numbered lists. Add, remove, and look up styles, or lift the whole style table from a designed template.",
        },
        {
          area: "Tables",
          items:
            "Rows, cell text, borders, cell shading, vertical alignment, row heights, repeating header rows. Unwrap a table back into paragraphs.",
        },
        {
          area: "Images",
          items:
            "Inline PNG, JPEG, GIF, BMP, TIFF, and SVG. List the images in a file and replace one by part name or alt text.",
        },
        {
          area: "Page setup",
          items:
            "Page size, margins, orientation, section breaks. Headers and footers for default, first, and even pages, with page numbers.",
        },
        {
          area: "Review",
          items:
            "Comments, footnotes, endnotes, bookmarks, external and internal hyperlinks. Accept or reject every tracked insertion and deletion.",
        },
        {
          area: "Fields",
          items: "Table of contents, merge fields, and PAGE, NUMPAGES, and DATE fields.",
        },
        {
          area: "Document",
          items:
            "Core and app properties, find and replace across every part, outline, statistics, plain-text extraction, validate().",
        },
      ],
      notYet: "Not yet",
      notYetItems: [
        "Merged table cells and nested tables",
        "Floating images, text boxes, and charts (kept when already in a file)",
        "A typed API for content controls (kept when already in a file)",
        "ISO/IEC 29500 Strict files",
        "Rendering to PDF, which is out of scope for good",
      ],
    },
    agents: {
      heading: "Readable by AI agents too",
      body: "Documents are increasingly generated by agents, so the docs are published in a form a model can fetch directly. Every snippet in them is a file that is type-checked against the library when the site builds, so an agent is not handed an API that no longer exists.",
      index: "An index of the docs, for a model to pick from.",
      full: "The guides, the recipes, and the API listing in one Markdown file.",
    },
    family: {
      heading: "One kit, three file formats",
      body: "Office Kit is a family of libraries built on the same rules: the ECMA-376 spec is the source of truth, files that go in must come out intact, and one ESM build has to run everywhere.",
    },
  },
  ja: {
    title: "@office-kit/docx: TypeScript で Word ファイルを読み書き・編集する",
    hero: {
      heading: "TypeScript で Word 文書を読み書き・編集する",
      lede: ".docx をゼロから作成することも、既存のファイルを開いて書き換えることもできます。段落、リスト、表、画像、ヘッダー、コメント、変更履歴は、どれも型付きのシンプルな関数で扱えます。ライブラリがモデル化していない部分は、手を加えずにそのまま書き戻します。Node でもブラウザーでも動作します。",
      start: "はじめる",
      playground: "プレイグラウンドを開く",
    },
    stage: {
      heading: "ドキュメントと、それを生成したコード",
      codeLabel: "hero-document.ts のソースコード",
      sheetLabel: "このコードが生成するドキュメント",
      note: "このページは実際の出力です。隣のコードでドキュメントを生成し、このサイトがそれを .docx のバイト列として保存したうえで `openDocx` で開き直し、読み取った内容を HTML として描画しています。ファイルをダウンロードして Word で開くと、見出しはスタイル、リストはリスト、表は表として作られていることがわかります。",
      building: "ドキュメントを生成しています…",
      download: ".docx をダウンロード",
      saved: (name, kilobytes) => `${name}（${kilobytes} KB）をこのタブ内で生成して保存しました。`,
      failed:
        "このブラウザーではドキュメントを生成できませんでした。エラーの内容はコンソールで確認できます。",
    },
    paths: [
      {
        title: "空のドキュメントから始める",
        body: "createDocx() はスタイル パーツとリレーションシップを設定済みのドキュメントを返すため、テンプレート ファイルを同梱する必要はありません。見出し、段落、箇条書きと番号付きリスト、表、インライン画像、ヘッダー、フッター、脚注、目次を追加できます。",
        code: "createDocx()",
        href: "/repl",
        link: "REPL で書いてみる",
      },
      {
        title: "手元のファイルを編集する",
        body: ".docx を開き、本文、ヘッダー、フッター、脚注、文末脚注、コメントにまたがってテキストを置換したり、代替テキストを手がかりに画像を差し替えたり、変更履歴の挿入と削除を承諾・却下したりしてから保存できます。ライブラリがモデル化していない部分は、元の状態のままファイルに戻されます。",
        code: "openDocx(bytes)",
        href: "/docs/recipes",
        link: "レシピを見る",
      },
      {
        title: "ページ上にドキュメントを表示する",
        body: "姉妹パッケージの @office-kit/docx-preview は、Docx またはバイト列を DOM 要素に描画します。オープンソースの docx-preview レンダラーを 1 つの関数でラップしたものです。ブラウザー内で動作するため、ファイルがアップロードされることはありません。",
        code: "previewToDOM(doc, element)",
        href: "/playground",
        link: "プレイグラウンドで試す",
      },
    ],
    template: {
      heading: "テンプレートは Word でデザインし、中身はコードで埋める",
      body: [
        "レターヘッド、契約書、フォントまで整えたレポート。大切な文書の多くは、誰かが手をかけた .docx としてすでに存在しています。それを開いて変わる部分だけを置き換えれば、それ以外はデザインしたとおりに残ります。",
        "Word は `{{name}}` のようなプレースホルダーを複数のランに分割して保存することがよくあります。照合は XML ではなく段落のテキストに対して行うため、分割されていても見つかります。また `replaceTextEverywhere` は本文だけでなく、ヘッダー、フッター、脚注・文末脚注、コメントも置換の対象にします。",
      ],
    },
    proof: {
      heading: "ラウンドトリップを機械で検証",
      lede: "他人の文書を編集するライブラリにとって、何より大切なのは文書を壊さないことです。それを裏付ける仕組みを紹介します。",
      items: [
        {
          claim: "実際の Word ファイルを読み込み、同じ内容のまま書き出します。",
          how: "テストスイートは、mammoth.js と python-docx がテストに使っている .docx のコーパスを開き、すべてのファイルを保存してから開き直して、段落とテキストを比較します。CI はプルリクエストごとに、これを Node 22 と Node 24 で実行します。",
        },
        {
          claim: "理解できないものには手を触れません。",
          how: "モデル化していない要素は元の位置にパススルー ノードとして保持され、読み取らないパーツもパッケージ内にそのまま残ります。コンテンツ コントロールやカスタム XML パーツも、編集によって失われることはありません。",
        },
        {
          claim: "validate() は、Word が修復を求める原因になる問題を検出します。",
          how: "存在しないパーツを指すリレーションシップ、参照先のないコメント・脚注・文末脚注の参照、対になっていないブックマーク、欠落したメディアなどです。テストスイートでは、20 通りの機能の組み合わせで問題が出ないことを確認しています。",
        },
        {
          claim: "バンドルに含まれるのは、インポートした関数だけです。",
          how: "副作用のないエクスポートが 120 個あり、クラスはありません。createDocx、appendParagraph、toUint8Array だけなら minify 後 42 KB、API 全体でも 133 KB です。この最小構成のバンドルが 50 KB を超えると CI が失敗します。",
        },
        {
          claim: "ランタイム依存は 1 つだけです。",
          how: "ZIP 処理のための fflate だけです。XML パーサーとシリアライザーは自前で実装しており、ライブラリ内に Node の組み込みモジュールをインポートしている箇所はありません。",
        },
      ],
    },
    compare: {
      heading: "docx パッケージとの違い",
      lede: "`docx` パッケージは JavaScript で Word ファイルを生成する定番の手段で、このライブラリではまだ作成できないものも作れます。両者の違いは方向性です。docx パッケージは新しい文書の生成に特化しているのに対し、このライブラリは手元にある文書の読み取りと編集にも対応しています。比較対象は docx 9.7.1 です。",
      capability: "機能",
      rows: [
        {
          topic: "既存の .docx を開いて編集",
          ours: {
            text: "対応。ファイルをモデルとして読み込み、任意の段落、ラン、表を変更可能",
            tone: "yes",
          },
          theirs: {
            text: "プレースホルダーのみ。patchDocument で {{tokens}} を新しい内容に差し替え",
            tone: "part",
          },
        },
        {
          topic: "ファイルの内容の読み取り",
          ours: {
            text: "テキスト、アウトライン、統計情報、画像、フィールド、ハイパーリンク、ブックマーク",
            tone: "yes",
          },
          theirs: {
            text: "非対応。patchDetector で取得できるのはプレースホルダーの一覧のみ",
            tone: "no",
          },
        },
        {
          topic: "プレースホルダーの置換先",
          ours: { text: "テキスト", tone: "part" },
          theirs: { text: "テキスト、段落、表、画像", tone: "yes" },
        },
        {
          topic: "変更履歴",
          ours: { text: "ファイル内の変更履歴を承諾・却下できるが、新規作成は不可", tone: "part" },
          theirs: { text: "挿入と削除を作成できるが、既存の変更履歴の解決は不可", tone: "part" },
        },
        {
          topic: "セルの結合、フローティング画像、テキストボックス",
          ours: { text: "未対応", tone: "no" },
          theirs: { text: "対応", tone: "yes" },
        },
        {
          topic: "API の形",
          ours: { text: "プレーンなデータを扱う、ツリーシェイキング可能な関数群。ESM のみ" },
          theirs: {
            text: "クラス インスタンスのツリー。ESM、CommonJS、script タグ用のバンドルを提供",
          },
        },
        {
          topic: "ランタイム依存",
          ours: { text: "1 つ（fflate）" },
          theirs: { text: "6 つ（jszip、xml-js など）" },
        },
      ],
      advice:
        "新しい文書を生成するだけで、セルの結合、フローティング画像、テキストボックス、CommonJS ビルドが必要なら、docx パッケージを選んでください。すでにある文書を読み取ったり、変更したり、検証したりする必要があるなら、このライブラリが適しています。",
    },
    caps: {
      heading: "現在できること",
      lede: "このライブラリはまだ 1.0 未満です。現時点で動作するものと、まだ対応していないものを以下にまとめます。すべての関数は [API リファレンス](/api)に掲載しています。",
      items: [
        {
          area: "テキスト",
          items:
            "段落、見出し、ラン（太字、斜体、下線、文字の色、蛍光ペン、サイズ、フォント）。配置、インデント、間隔、罫線、網かけ。",
        },
        {
          area: "リストとスタイル",
          items:
            "箇条書きと番号付きリスト。スタイルの追加、削除、検索と、デザイン済みテンプレートからのスタイル定義一式の取り込み。",
        },
        {
          area: "表",
          items:
            "行、セルのテキスト、罫線、セルの網かけ、垂直方向の配置、行の高さ、タイトル行の繰り返し。表を解除して段落に戻す操作。",
        },
        {
          area: "画像",
          items:
            "PNG、JPEG、GIF、BMP、TIFF、SVG のインライン画像。ファイル内の画像の一覧取得と、パーツ名または代替テキストによる差し替え。",
        },
        {
          area: "ページ設定",
          items:
            "用紙サイズ、余白、印刷の向き、セクション区切り。標準ページ・先頭ページ・偶数ページ用のヘッダーとフッター、ページ番号。",
        },
        {
          area: "校閲",
          items:
            "コメント、脚注、文末脚注、ブックマーク、外部・内部へのハイパーリンク。変更履歴の挿入と削除のすべての承諾・却下。",
        },
        {
          area: "フィールド",
          items: "目次、差し込みフィールド、PAGE・NUMPAGES・DATE フィールド。",
        },
        {
          area: "ドキュメント",
          items:
            "コア プロパティとアプリケーション プロパティ、すべてのパーツを対象にした検索と置換、アウトライン、統計情報、プレーンテキストの抽出、validate()。",
        },
      ],
      notYet: "未対応",
      notYetItems: [
        "表のセルの結合と入れ子の表",
        "フローティング画像、テキストボックス、グラフ（既存のファイルに含まれるものは保持されます）",
        "コンテンツ コントロール用の型付き API（既存のファイルに含まれるものは保持されます）",
        "ISO/IEC 29500 Strict 形式のファイル",
        "PDF へのレンダリング（今後も対応する予定はありません）",
      ],
    },
    agents: {
      heading: "AI エージェントにも読みやすく",
      body: "AI エージェントが文書を生成する場面が増えているため、ドキュメントはモデルが直接取得できる形式でも公開しています。掲載しているコードスニペットはすべて、サイトのビルド時にライブラリに対して型チェックされるファイルです。そのため、エージェントにもう存在しない API を教えてしまうことはありません。",
      index: "モデルが参照先を選ぶための、ドキュメントの目次（英語）。",
      full: "ガイド、レシピ、API 一覧を 1 つにまとめた Markdown ファイル（英語）。",
    },
    family: {
      heading: "1 つのキットで 3 つのファイル形式に対応",
      body: "Office Kit は、共通のルールに沿って作られたライブラリ群です。ECMA-376 仕様を唯一の拠り所とすること、読み込んだファイルを損なわずに書き出すこと、そして 1 つの ESM ビルドがどこでも動作することです。",
    },
  },
});
