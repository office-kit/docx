// Single source of truth for the public API listing.
//
// Consumed by:
//   - `/api` — renders the groups as sections on the docs site.
//   - `/llms-full.txt` — flattens the same groups into the API section
//     of the LLM-readable concatenation.
//   - `scripts/check-api-page.mjs` — diffs the entries here against the
//     live `@office-kit/docx` exports (plus a `previewToDOM` whitelist for
//     `@office-kit/docx-preview`). CI fails if a new export isn't added here.

import type { Translations } from "$lib/i18n/define";

export type ApiEntry = { name: string; sig?: string };
export type ApiGroupText = { title: string; description: string };
// Names and signatures are code and read the same in every locale; only the
// group headings are translated. English stays on the group itself because
// `/llms-full.txt` is English only.
export type ApiGroup = ApiGroupText & {
  translations: Translations<ApiGroupText>;
  entries: ApiEntry[];
};

export const apiGroups: ApiGroup[] = [
  {
    title: "Lifecycle",
    description: "Create, open, clone, and serialise a document.",
    translations: {
      ja: {
        title: "ライフサイクル",
        description: "ドキュメントの作成、読み込み、複製、シリアライズ。",
      },
    },
    entries: [
      { name: "createDocx", sig: "({ paragraphs? }?) => Docx" },
      { name: "openDocx", sig: "(bytes: Uint8Array) => Docx" },
      { name: "fromBlob", sig: "(blob: Blob) => Promise<Docx>" },
      { name: "toUint8Array", sig: "(doc: Docx) => Uint8Array" },
      { name: "toBlob", sig: "(doc: Docx) => Blob" },
      { name: "clone", sig: "(doc: Docx) => Docx" },
    ],
  },
  {
    title: "Paragraphs & blocks",
    description: "Append, insert, and remove body paragraphs, headings, and breaks.",
    translations: {
      ja: {
        title: "段落とブロック",
        description: "本文の段落、見出し、区切りの追加、挿入、削除。",
      },
    },
    entries: [
      { name: "appendParagraph" },
      { name: "insertParagraphAt" },
      { name: "removeParagraph" },
      { name: "appendHeading" },
      { name: "appendPageBreak" },
      { name: "appendLineBreak" },
      { name: "appendSectionBreak" },
      { name: "clearBody" },
      { name: "paragraphs" },
    ],
  },
  {
    title: "Inline & text",
    description:
      "Find and replace, runs and their formatting, paragraph alignment, indents, spacing, borders, and shading.",
    translations: {
      ja: {
        title: "インラインとテキスト",
        description: "検索と置換、ランとその書式、段落の配置、インデント、間隔、罫線、網かけ。",
      },
    },
    entries: [
      { name: "replaceText" },
      { name: "replaceTextEverywhere" },
      { name: "findText" },
      { name: "findTextEverywhere" },
      { name: "appendTextRun" },
      { name: "setParagraphText" },
      { name: "paragraphText" },
      { name: "setRunFormat" },
      { name: "clearRunFormat" },
      { name: "getRunFormat" },
      {
        name: "HIGHLIGHT_COLORS",
        sig: "readonly HighlightColor[] — the only w:highlight values run-format writers accept",
      },
      { name: "setParagraphAlignment" },
      { name: "getParagraphAlignment" },
      { name: "setParagraphIndent" },
      { name: "setParagraphSpacing" },
      { name: "setParagraphBorders" },
      { name: "setParagraphShading" },
      { name: "getParagraphStyle" },
      { name: "getParagraphNumbering" },
      { name: "mergeAdjacentRuns" },
      { name: "mergeAdjacentRunsInBody" },
      { name: "setRunOnOff" },
      { name: "setRunValProp" },
      { name: "getRunProp" },
      { name: "setParagraphOnOff" },
      { name: "setParagraphValProp" },
      { name: "getParagraphProp" },
      { name: "splitParagraphAt" },
      { name: "mergeParagraphIntoPrevious" },
      { name: "runTextLength" },
      { name: "isolateParagraphRunRange" },
    ],
  },
  {
    title: "Styles & numbering",
    description:
      "The style table, bullet and numbered lists, and lifting styles from a designed template.",
    translations: {
      ja: {
        title: "スタイルと段落番号",
        description:
          "スタイル定義、箇条書きと番号付きリスト、デザイン済みテンプレートからのスタイルの取り込み。",
      },
    },
    entries: [
      { name: "addStyle" },
      { name: "removeStyle" },
      { name: "listStyles" },
      { name: "ensureHeadingStyles" },
      { name: "findStyleIdByName" },
      { name: "setParagraphStyle" },
      { name: "addBulletList" },
      { name: "addNumberedList" },
      { name: "applyListToParagraph" },
      { name: "mergeStylesFromTemplate" },
      { name: "setStyleOnOff" },
      { name: "setStyleValProp" },
      { name: "getStyleProp" },
      { name: "setNumberingLevelOnOff" },
      { name: "setNumberingLevelVal" },
      { name: "getNumberingLevelProp" },
    ],
  },
  {
    title: "Tables",
    description: "Build tables and edit rows, cell text, borders, shading, and alignment.",
    translations: {
      ja: {
        title: "表",
        description: "表の作成と、行、セルのテキスト、罫線、網かけ、配置の編集。",
      },
    },
    entries: [
      { name: "addTable" },
      { name: "tables" },
      { name: "removeTable" },
      { name: "removeAllTables" },
      { name: "unwrapTable" },
      { name: "appendTableRow" },
      { name: "removeTableRow" },
      { name: "setTableRowAsHeader" },
      { name: "setTableRowHeight" },
      { name: "setTableBorders" },
      { name: "setTableCellText" },
      { name: "getTableCellText" },
      { name: "setTableCellShading" },
      { name: "setTableCellVerticalAlign" },
    ],
  },
  {
    title: "Table layout",
    description:
      "Word's Table Layout operations: rows, columns and cells at any position, merge and split, sizes, margins, sorting, text conversion and formulas.",
    translations: {
      ja: {
        title: "表のレイアウト",
        description:
          "Word の［レイアウト］（表ツール）タブの操作。任意の位置への行・列・セルの挿入、結合と分割、サイズ、余白、並べ替え、文字列への変換、計算式。",
      },
    },
    entries: [
      { name: "tableCellPlacements", sig: "(table) => TableCellPlacement[][]" },
      { name: "tableColumnCount" },
      { name: "tableColumnWidths" },
      { name: "tableCellBlocks" },
      { name: "insertTableRow", sig: "(table, index, { formatFrom? }?) => WmlTableRow" },
      { name: "deleteTableRows" },
      { name: "insertTableColumn", sig: "(table, index, { widthTwips?, formatFrom? }?) => void" },
      { name: "deleteTableColumns" },
      { name: "deleteTableCells", sig: "(table, range, 'left' | 'up') => void" },
      { name: "mergeTableCells", sig: "(table, range: TableCellRange) => void" },
      { name: "splitTableCell", sig: "(table, row, cell, { columns, rows }) => void" },
      { name: "splitTable", sig: "(doc, table, row) => WmlTable" },
      { name: "setTableWidth" },
      { name: "setTableAlignment" },
      { name: "setTableIndent" },
      { name: "setTableLayout" },
      { name: "setTableCellSpacing" },
      { name: "setTableDefaultCellMargins" },
      { name: "setTableAltText" },
      { name: "setTablePosition" },
      { name: "setTableCellWidth" },
      { name: "setTableCellMargins" },
      { name: "setTableCellTextDirection" },
      { name: "setTableColumnWidth" },
      { name: "distributeTableColumns" },
      { name: "autoFitTable", sig: "(table, 'contents' | 'window' | 'fixed') => void" },
      { name: "setTableRangeBorders", sig: "(table, range, edges, border?) => void" },
      { name: "sortTableRows", sig: "(table, { keys, headerRow? }) => void" },
      {
        name: "convertTextToTable",
        sig: "(doc, first, last, { separator, columns? }) => WmlTable",
      },
      { name: "evaluateTableFormula", sig: "(table, row, cell, formula) => number" },
      { name: "formatFieldNumber" },
    ],
  },
  {
    title: "Table styles",
    description:
      "Word's built-in table styles (Plain, Grid and List Tables), Table Style Options, and table style formatting.",
    translations: {
      ja: {
        title: "表のスタイル",
        description:
          "Word の組み込みの表スタイル（標準の表、グリッド テーブル、リスト テーブル）、表スタイルのオプション、表スタイルの書式。",
      },
    },
    entries: [
      { name: "BUILT_IN_TABLE_STYLES" },
      { name: "addBuiltInTableStyle", sig: "(doc, styleId) => void" },
      { name: "builtInTableStyle" },
      { name: "setTableStyle" },
      { name: "getTableStyle" },
      { name: "setTableLook", sig: "(table, look: TableLook) => void" },
      { name: "getTableLook" },
      { name: "setTableStyleFormatting", sig: "(doc, styleId, region, formatting) => void" },
    ],
  },
  {
    title: "Images",
    description:
      "Add inline images, list the ones in a file, and replace them by part name or alt text.",
    translations: {
      ja: {
        title: "画像",
        description:
          "インライン画像の追加、ファイル内の画像の一覧取得、パーツ名または代替テキストによる差し替え。",
      },
    },
    entries: [
      { name: "addImage" },
      { name: "addImageRun" },
      { name: "insertImageInto" },
      { name: "images" },
      { name: "imageReferences" },
      { name: "replaceImage" },
      { name: "replaceImageByAltText" },
      { name: "removeAllImages" },
      { name: "imageDrawings" },
      { name: "getImageInfo" },
      { name: "setImageAltText" },
      { name: "setImageSizeEmu" },
    ],
  },
  {
    title: "Pictures, charts & arrangement",
    description:
      "Picture Format and Arrange: wrapping, position, z-order, crop, shape, border, effects, color adjustments; Insert ▸ Chart with an embedded workbook. Drawings are addressed by their imageDrawings index.",
    translations: {
      ja: {
        title: "図、グラフ、配置",
        description:
          "［図の形式］と［配置］の操作。文字列の折り返し、位置、重なり順、トリミング、図形、枠線、効果、色の調整に加え、ブックを埋め込んだ［挿入］▸［グラフ］。描画オブジェクトは imageDrawings のインデックスで指定します。",
      },
    },
    entries: [
      { name: "readDrawing", sig: "(doc, drawing: XmlElement) => DrawingInfo" },
      { name: "imagePixelSize", sig: "(bytes) => ImagePixelSize | undefined" },
      { name: "imageNaturalSizeEmu", sig: "(bytes) => { widthEmu, heightEmu } | undefined" },
      { name: "setDrawingWrap", sig: "(doc, index, wrap: WrapStyle, position?) => void" },
      { name: "setDrawingPosition", sig: "(doc, index, { horizontal?, vertical? }) => void" },
      { name: "setDrawingAnchorOptions", sig: "(doc, index, AnchorOptions) => void" },
      { name: "arrangeDrawing", sig: "(doc, index, op: DrawingOrder) => void" },
      { name: "setDrawingTransform", sig: "(doc, index, { rotation?, flipH?, flipV? }) => void" },
      { name: "setDrawingName", sig: "(doc, index, name) => void" },
      { name: "setDrawingHidden", sig: "(doc, index, hidden) => void" },
      { name: "setDrawingAspectLock", sig: "(doc, index, locked) => void" },
      { name: "setDrawingHyperlink", sig: "(doc, index, url | undefined) => void" },
      { name: "removeDrawing", sig: "(doc, index) => void" },
      { name: "setPictureCrop", sig: "(doc, index, PictureCrop) => void" },
      { name: "setPictureGeometry", sig: "(doc, index, preset: string) => void" },
      { name: "setPictureOutline", sig: "(doc, index, PictureOutline | undefined) => void" },
      { name: "setPictureEffects", sig: "(doc, index, PictureEffects) => void" },
      { name: "setPictureColorAdjustments", sig: "(doc, index, PictureColorAdjustments) => void" },
      { name: "changePicture", sig: "(doc, index, bytes, contentType?) => void" },
      { name: "resetPicture", sig: "(doc, index, { size? }?) => void" },
      {
        name: "addChartRun",
        sig: "(doc, spec: ChartSpec, { widthEmu, heightEmu, name?, altText? }) => WmlRun",
      },
      { name: "readChart", sig: "(doc, drawing: XmlElement) => ChartSpec | undefined" },
      { name: "setChart", sig: "(doc, index, spec: ChartSpec) => void" },
    ],
  },
  {
    title: "Headers, footers, sections",
    description:
      "Headers and footers for default, first, and even pages, plus page size, margins, and orientation.",
    translations: {
      ja: {
        title: "ヘッダー、フッター、セクション",
        description:
          "標準ページ・先頭ページ・偶数ページ用のヘッダーとフッター、用紙サイズ、余白、印刷の向き。",
      },
    },
    entries: [
      { name: "addHeader" },
      { name: "addFooter" },
      { name: "addPageNumberFooter" },
      { name: "setPageSize" },
      { name: "setPageMargins" },
      { name: "setPageOrientation" },
      { name: "headers" },
      { name: "footers" },
      { name: "removeAllHeaders" },
      { name: "removeAllFooters" },
    ],
  },
  {
    title: "Comments, notes, hyperlinks, bookmarks",
    description: "Review and navigation content: add it, list it, rewrite links, or strip it out.",
    translations: {
      ja: {
        title: "コメント、脚注、ハイパーリンク、ブックマーク",
        description:
          "校閲やナビゲーションのためのコンテンツの追加、一覧取得、リンクの書き換え、削除。",
      },
    },
    entries: [
      { name: "addComment" },
      { name: "addFootnote" },
      { name: "addEndnote" },
      { name: "removeAllComments" },
      { name: "removeAllFootnotes" },
      { name: "removeAllEndnotes" },
      { name: "addHyperlink" },
      { name: "addInternalHyperlink" },
      { name: "externalHyperlinks" },
      { name: "setHyperlinkUrl" },
      { name: "removeAllHyperlinks" },
      { name: "addBookmark" },
      { name: "removeBookmark" },
      { name: "removeAllBookmarks" },
      { name: "bookmarks" },
    ],
  },
  {
    title: "Fields & tracked changes",
    description:
      "Complex fields such as a table of contents or a merge field, and bulk accept or reject of tracked insertions and deletions.",
    translations: {
      ja: {
        title: "フィールドと変更履歴",
        description:
          "目次や差し込みフィールドなどの複合フィールドと、変更履歴の挿入・削除の一括承諾・却下。",
      },
    },
    entries: [
      { name: "appendField" },
      { name: "addTableOfContents" },
      { name: "appendMergeField" },
      { name: "acceptAllRevisions" },
      { name: "rejectAllRevisions" },
    ],
  },
  {
    title: "Insert at a position",
    description:
      "What Word's Insert tab places at a character offset of a paragraph: complex fields with computed results, links, bookmarks, symbols, equations, drop caps and signature lines.",
    translations: {
      ja: {
        title: "指定位置への挿入",
        description:
          "Word の［挿入］タブで段落内の文字位置に配置できるもの。計算結果付きの複合フィールド、リンク、ブックマーク、記号と特殊文字、数式、ドロップ キャップ、署名欄。",
      },
    },
    entries: [
      {
        name: "insertField",
        sig: "(doc, paragraph, offset, instruction, { result?, rPr?, lang?, context? }?) => WmlRun[]",
      },
      { name: "buildComplexField", sig: "(instruction, result, rPr?, beginChildren?) => WmlRun[]" },
      { name: "updateFields", sig: "(doc, { now?, page?, pageCount?, types? … }?) => number" },
      { name: "complexFields", sig: "(doc) => FieldInfo[]" },
      { name: "insertFormField", sig: "(doc, paragraph, offset, FormFieldOptions) => WmlRun[]" },
      { name: "formatDatePicture", sig: "(date, picture, locale?) => string" },
      { name: "quoteFieldArgument", sig: "(value) => string" },
      {
        name: "insertHyperlink",
        sig: "(doc, paragraph, offset, text, { url? | bookmark?, tooltip?, targetFrame? }) => XmlElement",
      },
      { name: "paragraphHyperlinks", sig: "(doc, paragraph) => HyperlinkInfo[]" },
      { name: "editHyperlink" },
      { name: "removeHyperlink" },
      { name: "insertBookmark", sig: "(doc, name, start: TextPoint, end?: TextPoint) => number" },
      { name: "isValidBookmarkName", sig: "(name) => boolean" },
      { name: "ensureReferenceBookmark", sig: "(doc, paragraph, run?) => string" },
      { name: "insertSymbol", sig: "(doc, paragraph, offset, SymbolSpec, rPr?) => WmlRun" },
      {
        name: "insertEquation",
        sig: "(doc, paragraph, offset, linear, { display? }?) => XmlElement",
      },
      { name: "buildEquation", sig: "(linear, { display? }?) => XmlElement" },
      { name: "equationLinear", sig: "(element) => string" },
      { name: "paragraphEquations", sig: "(paragraph) => XmlElement[]" },
      { name: "setEquation", sig: "(doc, paragraph, index, linear) => XmlElement" },
      { name: "getDropCap", sig: "(doc, paragraph) => DropCapOptions" },
      {
        name: "setDropCap",
        sig: "(doc, paragraph, { position, lines?, distanceTwips?, font? }) => void",
      },
      {
        name: "insertSignatureLine",
        sig: "(doc, paragraph, offset, SignatureLineOptions) => WmlRun",
      },
    ],
  },
  {
    title: "Pages, headers & page numbers",
    description:
      "Cover pages, header and footer content, page numbers and their format, and another document's content merged in (Insert ▸ Object ▸ Text from File).",
    translations: {
      ja: {
        title: "ページ、ヘッダー、ページ番号",
        description:
          "表紙、ヘッダーとフッターの内容、ページ番号とその書式、別のドキュメントの内容の取り込み（［挿入］▸［オブジェクト］▸［ファイルからテキスト］）。",
      },
    },
    entries: [
      { name: "insertCoverPage", sig: "(doc, paragraphs) => void" },
      { name: "removeCoverPage", sig: "(doc) => boolean" },
      { name: "hasCoverPage", sig: "(doc) => boolean" },
      { name: "setDifferentFirstPage", sig: "(doc, on) => void" },
      { name: "headerFooterParagraphs", sig: "(doc, kind, type?) => WmlParagraph[]" },
      { name: "setHeaderFooterParagraphs", sig: "(doc, kind, type, paragraphs) => void" },
      { name: "removeHeaderFooter", sig: "(doc, kind, type?) => boolean" },
      { name: "ensureHeaderFooterStyle", sig: "(doc, kind) => string" },
      { name: "insertPageNumbers", sig: "(doc, { position, align, style? }) => void" },
      { name: "removePageNumbers", sig: "(doc) => number" },
      {
        name: "setPageNumberFormat",
        sig: "(doc, { format?, start?, chapterStyle?, chapterSeparator? }) => void",
      },
      { name: "getPageNumberFormat", sig: "(doc) => PageNumberFormatOptions" },
      { name: "PAGE_NUMBER_FORMATS" },
      { name: "CHAPTER_SEPARATORS" },
      { name: "insertDocumentContent", sig: "(doc, source: Docx, blockIndex) => number" },
    ],
  },
  {
    title: "Document properties",
    description: "Core and app properties, with shortcuts for the title and the author.",
    translations: {
      ja: {
        title: "ドキュメントのプロパティ",
        description:
          "コアプロパティとアプリケーションプロパティ。タイトルと作成者にはショートカットがあります。",
      },
    },
    entries: [
      { name: "coreProperties" },
      { name: "setCoreProperties" },
      { name: "appProperties" },
      { name: "setAppProperties" },
      { name: "title" },
      { name: "author" },
      { name: "setTitle" },
      { name: "setAuthor" },
      { name: "getDocumentSetting" },
      { name: "setDocumentSettingOnOff" },
      { name: "setDocumentSettingVal" },
    ],
  },
  {
    title: "Shapes, text boxes, WordArt & ink (VML)",
    description:
      "Draw shapes from Word's gallery, text boxes, WordArt, ink strokes and drawing canvases as ECMA-376 VML (`w:pict`), and edit their fill, outline, shadow, layout, wrap, z-order, text and grouping. Shape handles are the VML elements.",
    translations: {
      ja: {
        title: "図形、テキストボックス、ワードアート、インク（VML）",
        description:
          "Word のギャラリーにある図形、テキストボックス、ワードアート、インクのストローク、描画キャンバスを ECMA-376 の VML（`w:pict`）として描画し、塗りつぶし、枠線、影、レイアウト、文字列の折り返し、重なり順、テキスト、グループ化を編集します。図形のハンドルは VML 要素そのものです。",
      },
    },
    entries: [
      {
        name: "addShape",
        sig: "(doc, paragraph, { preset, width, height, left?, top?, wrap?, fill?, stroke?, text?, points?, ink? }, index?) => XmlElement",
      },
      {
        name: "addWordArt",
        sig: "(doc, paragraph, { text, width, height, font?, size?, bold?, italic? }, index?) => XmlElement",
      },
      {
        name: "addDrawingCanvas",
        sig: "(doc, paragraph, { width, height }, index?) => XmlElement",
      },
      { name: "addShapeToGroup", sig: "(doc, group, options) => XmlElement" },
      { name: "vmlShapes", sig: "(doc) => XmlElement[]" },
      { name: "runShape", sig: "(run) => XmlElement | undefined" },
      { name: "shapeKind" },
      { name: "shapePreset" },
      { name: "SHAPE_PRESETS" },
      { name: "SHAPE_CATEGORIES" },
      { name: "changeShapePreset" },
      { name: "getShapePoints" },
      { name: "setShapePoints" },
      { name: "getShapeLayout" },
      { name: "setShapeLayout" },
      { name: "getShapeWrap" },
      { name: "setShapeWrap" },
      { name: "setShapeOrder" },
      { name: "setShapeMoveWithText" },
      { name: "getShapeFill" },
      { name: "setShapeFill" },
      { name: "vmlImageData" },
      { name: "FILL_PATTERNS" },
      { name: "getShapeStroke" },
      { name: "setShapeStroke" },
      { name: "getShapeShadow" },
      { name: "setShapeShadow" },
      { name: "shapeText" },
      { name: "setShapeText" },
      { name: "getTextBoxLayout" },
      { name: "setTextBoxLayout" },
      { name: "linkTextBoxes" },
      { name: "linkedTextBox" },
      { name: "getWordArt" },
      { name: "setWordArt" },
      { name: "getShapeAltText" },
      { name: "setShapeAltText" },
      { name: "getShapeName" },
      { name: "setShapeName" },
      { name: "removeShape" },
      { name: "groupShapes" },
      { name: "ungroupShapes" },
      { name: "groupMembers" },
    ],
  },
  {
    title: "SmartArt",
    description:
      "Insert SmartArt diagrams (DrawingML diagrams: data model, layout, style and colour parts) and edit their bullets, layout, colours, style and size.",
    translations: {
      ja: {
        title: "SmartArt",
        description:
          "SmartArt グラフィック（DrawingML のダイアグラム。データモデル、レイアウト、スタイル、色の各パーツ）の挿入と、箇条書き、レイアウト、色、スタイル、サイズの編集。",
      },
    },
    entries: [
      {
        name: "addSmartArt",
        sig: "(doc, paragraph, { layout, nodes, width?, height?, colors?, style? }, index?) => WmlRun",
      },
      { name: "smartArts", sig: "(doc) => SmartArtRef[]" },
      { name: "runSmartArt", sig: "(doc, run) => SmartArtRef | undefined" },
      { name: "getSmartArt" },
      { name: "setSmartArtNodes" },
      { name: "setSmartArtLayout" },
      { name: "setSmartArtColors" },
      { name: "setSmartArtStyle" },
      { name: "setSmartArtSize" },
      { name: "removeSmartArt" },
      { name: "SMARTART_LAYOUTS" },
      { name: "SMARTART_COLORS" },
      { name: "SMARTART_STYLES" },
    ],
  },
  {
    title: "Diagnostics",
    description:
      "Validate the package and read a document back: outline, fields, statistics, plain text.",
    translations: {
      ja: {
        title: "診断",
        description:
          "パッケージの検証と、ドキュメント内容の読み取り（アウトライン、フィールド、統計情報、プレーンテキスト）。",
      },
    },
    entries: [
      { name: "validate" },
      { name: "validatePackage" },
      { name: "statistics" },
      { name: "outline" },
      { name: "fields" },
      { name: "text" },
    ],
  },
  {
    title: "Review: tracked changes",
    description:
      "List, accept, and reject revisions; record insertions, deletions, and formatting changes as tracked; compare two documents.",
    translations: {
      ja: {
        title: "校閲: 変更履歴",
        description:
          "変更履歴の一覧取得、承諾、却下。挿入・削除・書式変更の変更履歴としての記録と、2 つのドキュメントの比較。",
      },
    },
    entries: [
      { name: "revisions", sig: "(doc: Docx) => RevisionInfo[]" },
      { name: "acceptRevisions", sig: "(doc: Docx, ids: readonly string[]) => number" },
      { name: "rejectRevisions", sig: "(doc: Docx, ids: readonly string[]) => number" },
      { name: "insertTrackedText" },
      { name: "deleteTrackedText" },
      { name: "paragraphMarkRevision" },
      { name: "trackParagraphMark" },
      { name: "trackRunFormatChange" },
      { name: "trackParagraphFormatChange" },
      {
        name: "compareDocuments",
        sig: "(original: Docx, revised: Docx, options: CompareOptions) => Docx",
      },
    ],
  },
  {
    title: "Review: comments, proofing, accessibility",
    description:
      "Read and edit comments, set the proofing language, count words as Word does, and check accessibility.",
    translations: {
      ja: {
        title: "校閲: コメント、文章校正、アクセシビリティ",
        description:
          "コメントの読み取りと編集、校正言語の設定、Word と同じ方法での文字数カウント、アクセシビリティ チェック。",
      },
    },
    entries: [
      { name: "comments", sig: "(doc: Docx) => CommentInfo[]" },
      { name: "setCommentText" },
      { name: "removeComment" },
      { name: "getRunLanguage" },
      { name: "setRunLanguage" },
      { name: "wordCount", sig: "(doc: Docx, options?: WordCountOptions) => WordCount" },
      { name: "checkAccessibility", sig: "(doc: Docx) => AccessibilityIssue[]" },
      { name: "contrastRatio" },
    ],
  },
  {
    title: "Document protection",
    description:
      "Restrict editing (with an optional password), editable exceptions, and Always Open Read-Only.",
    translations: {
      ja: {
        title: "ドキュメントの保護",
        description:
          "編集の制限（パスワードは任意）、編集を許可する例外、常に読み取り専用で開く設定。",
      },
    },
    entries: [
      { name: "documentProtection" },
      { name: "protectDocument", sig: "(doc: Docx, options: ProtectOptions) => void" },
      { name: "unprotectDocument", sig: "(doc: Docx, password: string) => boolean" },
      { name: "verifyProtectionPassword" },
      { name: "writeProtection" },
      { name: "setWriteProtection" },
      { name: "editableRanges" },
      { name: "addEditableRange" },
      { name: "removeEditableRange" },
    ],
  },
  {
    title: "View & custom properties",
    description: "The view and zoom a document opens with, and custom document properties.",
    translations: {
      ja: {
        title: "表示とユーザー設定のプロパティ",
        description:
          "ドキュメントを開いたときの表示モードとズーム、ユーザー設定のドキュメント プロパティ。",
      },
    },
    entries: [
      { name: "documentView" },
      { name: "setDocumentView" },
      { name: "documentZoom" },
      { name: "setDocumentZoom" },
      { name: "customProperties" },
      { name: "setCustomProperty" },
      { name: "removeCustomProperty" },
    ],
  },
  {
    title: "Low-level part access",
    description: "The parsed side parts, for when the functions above do not reach far enough.",
    translations: {
      ja: {
        title: "低レベルのパーツ アクセス",
        description: "解析済みの補助パーツ。上記の関数では手が届かない場合に使います。",
      },
    },
    entries: [
      { name: "stylesPart" },
      { name: "numberingPart" },
      { name: "commentsPart" },
      { name: "footnotesPart" },
      { name: "endnotesPart" },
      { name: "xmlPartNames" },
      { name: "getRawPartRoot" },
      { name: "markRawPartDirty" },
      { name: "childElementsOf" },
      { name: "appendChildElement" },
      { name: "makePropsElement" },
      { name: "getElementAttr" },
      { name: "setElementAttr" },
      { name: "getElementProp" },
      { name: "setElementOnOff" },
      { name: "setElementValProp" },
    ],
  },
  {
    title: "Stories & section headers",
    description:
      "Headers, footers, footnotes, endnotes and comments as typed block content, and which header/footer each section shows.",
    translations: {
      ja: {
        title: "ストーリーとセクションのヘッダー",
        description:
          "ヘッダー、フッター、脚注、文末脚注、コメントを型付きのブロック コンテンツとして扱う関数と、各セクションに表示するヘッダー・フッターの指定。",
      },
    },
    entries: [
      { name: "storyBody", sig: "(doc, ref: StoryRef) => WmlBody | undefined" },
      {
        name: "contentControlBlocks",
        sig: "(block: WmlBlock) => readonly WmlBlock[] | undefined",
      },
      { name: "wrappedRuns", sig: "(inline: WmlInline) => readonly WmlRun[] | undefined" },
      { name: "storyView", sig: "(doc, ref: StoryRef) => Docx | undefined" },
      { name: "storyKey", sig: "(ref: StoryRef) => string" },
      { name: "parseStoryKey", sig: "(key: string) => StoryRef | undefined" },
      { name: "sectionProperties", sig: "(doc) => Array<XmlElement | undefined>" },
      {
        name: "resolveHeaderFooter",
        sig: "(doc, section, kind, type) => { partName, section } | undefined",
      },
      { name: "ensureHeaderFooter", sig: "(doc, section, kind, type) => string" },
      { name: "isHeaderFooterLinked", sig: "(doc, section, kind, type) => boolean" },
      { name: "setHeaderFooterLinked", sig: "(doc, section, kind, type, linked) => void" },
    ],
  },
  {
    title: "Character & paragraph formatting",
    description:
      "Theme colours, every underline style, theme fonts, run shading and borders, per-side paragraph borders, tab stops, phonetic guides, enclosed characters, list definitions, and Word's built-in styles.",
    translations: {
      ja: {
        title: "文字書式と段落書式",
        description:
          "テーマの色、すべての下線スタイル、テーマのフォント、ランの網かけと罫線、段落の辺ごとの罫線、タブ位置、ルビ、囲い文字、リストの定義、Word の組み込みスタイル。",
      },
    },
    entries: [
      { name: "setRunColor", sig: "(run, color: ColorValue | undefined) => void" },
      { name: "setRunUnderline", sig: "(run, underline: UnderlineValue | undefined) => void" },
      { name: "setRunFont", sig: "(run, script, font: FontChoice | undefined) => void" },
      { name: "setRunShading", sig: "(run, shading: ShadingOptions | undefined) => void" },
      { name: "setRunBorder", sig: "(run, line: BorderLine | undefined) => void" },
      { name: "setParagraphBorder", sig: "(p, side, line: BorderLine | undefined) => void" },
      { name: "setParagraphTabs", sig: "(p, tabs: TabStop[]) => void" },
      { name: "getParagraphTabs", sig: "(p) => TabStop[]" },
      { name: "buildRubyRun", sig: "(base, ruby, options, rPr?) => WmlRun" },
      { name: "readRuby", sig: "(run) => RubyInfo | undefined" },
      { name: "buildEnclosedCharacterRuns", sig: "(text, options, rPr?) => WmlRun[]" },
      { name: "addListDefinition", sig: "(doc, levels: ListLevel[]) => number" },
      { name: "restartList", sig: "(doc, numId, ilvl, start) => number | undefined" },
      { name: "ensureBuiltinStyle", sig: "(doc, styleId) => boolean" },
      { name: "builtinStyles", sig: "() => BuiltinStyleInfo[]" },
      { name: "updateStyleFormatting", sig: "(doc, styleId, { run?, paragraph? }) => boolean" },
      { name: "THEME_COLORS" },
      { name: "UNDERLINE_STYLES" },
      { name: "SHADING_PATTERNS" },
      { name: "BORDER_LINE_STYLES" },
      { name: "RUBY_ALIGNMENTS" },
      { name: "ENCLOSURES" },
      { name: "NUMBER_FORMATS" },
    ],
  },
  {
    title: "Page-size & margin constants",
    description: "Ready-made values for setPageSize and setPageMargins, plus the library version.",
    translations: {
      ja: {
        title: "用紙サイズと余白の定数",
        description: "setPageSize と setPageMargins にそのまま渡せる値と、ライブラリのバージョン。",
      },
    },
    entries: [
      { name: "PAGE_SIZE_A4" },
      { name: "PAGE_SIZE_LETTER" },
      { name: "MARGINS_NORMAL" },
      { name: "VERSION" },
    ],
  },
  {
    title: "Section layout",
    description:
      "Columns, line numbers, page borders, vertical alignment, document grid and section breaks, per section.",
    translations: {
      ja: {
        title: "セクションのレイアウト",
        description:
          "セクション単位の段組み、行番号、ページ罫線、垂直方向の配置、文字グリッド、セクション区切り。",
      },
    },
    entries: [
      { name: "sectionCount", sig: "(doc) => number" },
      { name: "sectionIndexAt", sig: "(doc, blockIndex) => number" },
      { name: "getSectionProperties", sig: "(doc, section?) => SectionProperties" },
      { name: "setSectionProperties", sig: "(doc, patch, scope?) => void" },
      { name: "insertSectionBreak", sig: "(doc, blockIndex, start?) => number" },
      { name: "PAGE_BORDER_LINE_STYLES", sig: "readonly string[] — line styles for page borders" },
      { name: "PAGE_BORDER_ART", sig: "readonly string[] — art border names (ST_Border)" },
    ],
  },
  {
    title: "Design: themes, style sets, page background",
    description:
      "The theme part (colours, fonts, effects), style sets, default paragraph spacing, page colour and watermarks.",
    translations: {
      ja: {
        title: "デザイン: テーマ、スタイル セット、ページの背景",
        description:
          "テーマ パーツ（配色、フォント、効果）、スタイル セット、段落の既定の間隔、ページの色、透かし。",
      },
    },
    entries: [
      { name: "setTheme", sig: "(doc, theme: ThemeDefinition) => void" },
      { name: "getTheme", sig: "(doc) => ThemeInfo | undefined" },
      { name: "setThemeColors", sig: "(doc, colors: ThemeColorScheme) => void" },
      { name: "setThemeFonts", sig: "(doc, fonts: ThemeFontScheme) => void" },
      { name: "setThemeEffects", sig: "(doc, effects: ThemeEffectScheme) => void" },
      {
        name: "themeColorValue",
        sig: "(colors, themeColor, { shade?, tint? }?) => string | undefined",
      },
      { name: "applyStyleSet", sig: "(doc, set: StyleSetDefinition) => void" },
      { name: "currentStyleSet", sig: "(doc) => StyleSetDefinition | undefined" },
      { name: "setDefaultParagraphSpacing", sig: "(doc, { before, after, line }) => void" },
      { name: "getDefaultParagraphSpacing", sig: "(doc) => DefaultParagraphSpacing | undefined" },
      { name: "setPageColor", sig: "(doc, color: PageColor | undefined) => void" },
      { name: "getPageColor", sig: "(doc) => PageColor | undefined" },
      { name: "setWatermark", sig: "(doc, watermark: Watermark | undefined) => void" },
      { name: "getWatermark", sig: "(doc) => WatermarkInfo | undefined" },
      { name: "THEMES" },
      { name: "THEME_COLOR_SCHEMES" },
      { name: "THEME_COLOR_SLOTS" },
      { name: "THEME_FONT_SCHEMES" },
      { name: "THEME_EFFECT_SCHEMES" },
      { name: "STYLE_SETS" },
      { name: "PARAGRAPH_SPACING_PRESETS" },
    ],
  },
  {
    title: "References",
    description:
      "Word's References tab: tables of contents and figures, footnotes and endnotes, citations and bibliography, captions, index, and tables of authorities — with computed field results.",
    translations: {
      ja: {
        title: "参考資料",
        description:
          "Word の［参考資料］タブ。目次と図表目次、脚注と文末脚注、引用文献と文献目録、図表番号、索引、引用文献一覧を、計算済みのフィールド結果とともに扱います。",
      },
    },
    entries: [
      { name: "insertTableOfContents", sig: "(doc, at, options?) => void" },
      { name: "tableOfContentsInstruction" },
      { name: "removeTableOfContents" },
      { name: "setTocLevel", sig: "(doc, paragraph, level | 'none') => void" },
      { name: "tocLevel" },
      { name: "updateTables", sig: "(doc, { types?, pageOf? }?) => number" },
      {
        name: "insertNote",
        sig: "(doc, paragraph, offset, 'footnote' | 'endnote', options?) => number",
      },
      { name: "noteMarks" },
      { name: "noteText" },
      { name: "setNoteText" },
      { name: "noteProperties" },
      { name: "setNoteProperties" },
      { name: "convertNotes" },
      { name: "bibliographySources" },
      { name: "setBibliographySources" },
      { name: "bibliographyStyle" },
      { name: "setBibliographyStyle" },
      { name: "suggestSourceTag" },
      { name: "citationInstruction" },
      { name: "insertCitation" },
      { name: "insertBibliography" },
      {
        name: "insertCaption",
        sig: "(doc, blockIndex, 'above' | 'below', options) => WmlParagraph",
      },
      { name: "captionLabels" },
      { name: "addCaptionLabel" },
      { name: "indexEntryInstruction" },
      { name: "markIndexEntry" },
      { name: "markAllIndexEntries" },
      { name: "indexInstruction" },
      { name: "insertIndex" },
      { name: "markAuthorityCitation" },
      { name: "insertTableOfAuthorities" },
      { name: "CITATION_STYLES" },
      { name: "SOURCE_TYPES" },
      { name: "SOURCE_FIELDS" },
      { name: "BIBLIOGRAPHY_TITLES" },
      { name: "CAPTION_LABELS" },
      { name: "AUTHORITY_CATEGORIES" },
    ],
  },
  {
    title: "Mailings",
    description:
      "Envelopes, labels, and mail merge: data source settings, merge fields and rules, previewing a recipient, and merging to a new document.",
    translations: {
      ja: {
        title: "差し込み文書",
        description:
          "封筒、ラベル、差し込み印刷。データ ソースの設定、差し込みフィールドとルール、宛先のプレビュー、新規文書への差し込み。",
      },
    },
    entries: [
      { name: "addEnvelope" },
      { name: "removeEnvelope" },
      { name: "createLabelDocument", sig: "(options) => Docx" },
      { name: "updateLabels" },
      { name: "mailMergeSettings" },
      { name: "setMailMergeDocumentType" },
      { name: "parseRecipientCsv" },
      { name: "recipientListToCsv" },
      { name: "attachRecipientList" },
      { name: "autoFieldMap" },
      { name: "mailMergeFieldMap" },
      { name: "setMailMergeFieldMap" },
      { name: "setRecipientInclusion" },
      { name: "recipientInclusion" },
      { name: "mergeFieldInstruction" },
      { name: "insertMergeField" },
      { name: "addressBlockInstruction" },
      { name: "insertAddressBlock" },
      { name: "greetingLineInstruction" },
      { name: "insertGreetingLine" },
      { name: "insertMergeRule" },
      { name: "mergeFieldNames" },
      { name: "mergeFieldErrors" },
      { name: "previewMailMerge", sig: "(doc, list, index | undefined, options?) => void" },
      { name: "mergeToNewDocument", sig: "(doc, list, options?) => Docx" },
      { name: "ENVELOPE_SIZES" },
      { name: "LABEL_PRODUCTS" },
      { name: "ADDRESS_FIELDS" },
      { name: "ADDRESS_NAME_FORMATS" },
      { name: "GREETING_NAME_FORMATS" },
    ],
  },
  {
    title: "Browser preview (@office-kit/docx-preview)",
    description: "The companion package's single entry point.",
    translations: {
      ja: {
        title: "ブラウザー プレビュー（@office-kit/docx-preview）",
        description: "姉妹パッケージの唯一のエントリーポイント。",
      },
    },
    entries: [{ name: "previewToDOM", sig: "(source, container, options?) => Promise<Handle>" }],
  },
];

export const apiTotalCount: number = apiGroups.reduce((n, g) => n + g.entries.length, 0);
