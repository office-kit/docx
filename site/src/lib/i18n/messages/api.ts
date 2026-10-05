// The /api page. Group titles and descriptions live with the groups in
// `$lib/api-groups`.

import { defineMessages } from "../define.js";

const SOURCE = "https://github.com/office-kit/docx/blob/main/src/api/index.ts";

export default defineMessages({
  en: {
    title: "API reference",
    groups: "Groups",
    groupsLabel: "API groups",
    lede: (total: number, groups: number) =>
      `Every public export of \`@office-kit/docx\`, plus the one function of \`@office-kit/docx-preview\`: ${total} names in ${groups} groups. The library is functions and constants only, with no classes. Most functions take a \`Docx\` as their first argument, and the module has no side effects, so a bundler drops whatever you do not import.`,
    sync: `A CI check compares this page with the built package, so it cannot fall behind. For parameter shapes, read the type declarations that ship with the package or [the source](${SOURCE}). For code you can paste, see [Getting started](/docs/getting-started) and the [recipes](/docs/recipes).`,
    filter: "Filter by name",
    placeholder: "Table, replace, Footnote…",
    noMatch: (query: string) => `No export name contains “${query}”.`,
    count: (visible: number, total: number) => `${visible} of ${total} exports`,
  },
  ja: {
    title: "API リファレンス",
    groups: "グループ",
    groupsLabel: "API のグループ",
    lede: (total: number, groups: number) =>
      `\`@office-kit/docx\` のすべての公開エクスポートと、\`@office-kit/docx-preview\` の唯一の関数を、${groups} のグループに分けて掲載しています（計 ${total} 個）。ライブラリは関数と定数だけで構成されており、クラスはありません。ほとんどの関数は第 1 引数に \`Docx\` を受け取ります。モジュールには副作用がないため、インポートしなかったものはバンドラーが取り除きます。`,
    sync: `このページは CI で実際にビルドしたパッケージと照合しているため、内容が古くなることはありません。引数の型は、パッケージに同梱されている型定義か[ソースコード](${SOURCE})を参照してください。そのまま貼り付けて使えるコードは[はじめに](/docs/getting-started)と[レシピ](/docs/recipes)にあります。`,
    filter: "名前で絞り込み",
    placeholder: "Table、replace、Footnote…",
    noMatch: (query: string) => `「${query}」を含むエクスポート名はありません。`,
    count: (visible: number, total: number) => `${total} 個中 ${visible} 個のエクスポート`,
  },
});
