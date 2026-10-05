/**
 * Whether a font is an East Asian (CJK) font. Word's font box applies such a
 * font to East Asian text as well as Latin text; a Latin font leaves the East
 * Asian font alone, so Japanese text keeps 游明朝 when the Latin font changes.
 */

// Fonts with CJK glyphs that are known by Latin names.
const LATIN_NAMED: ReadonlySet<string> = new Set([
  "Yu Mincho",
  "Yu Gothic",
  "Yu Gothic UI",
  "MS Mincho",
  "MS PMincho",
  "MS Gothic",
  "MS PGothic",
  "MS UI Gothic",
  "Meiryo",
  "Meiryo UI",
  "Osaka",
  "SimSun",
  "SimHei",
  "NSimSun",
  "DengXian",
  "Microsoft YaHei",
  "Microsoft JhengHei",
  "PMingLiU",
  "MingLiU",
  "Malgun Gothic",
  "Batang",
  "Gulim",
  "Dotum",
]);
const LATIN_NAMED_FAMILIES = /^(Yu |Hiragino |Noto (Sans|Serif) (CJK|JP|SC|TC|KR)|Source Han )/;
// Kana, CJK ideographs and full-width forms: a font named in these is one.
const CJK_NAME = /[぀-ヿ㐀-鿿가-힯＀-￯]/;

export function isEastAsianFont(name: string): boolean {
  return CJK_NAME.test(name) || LATIN_NAMED.has(name) || LATIN_NAMED_FAMILIES.test(name);
}

// Kana and CJK ideographs: text Word draws with the East Asian font.
const EAST_ASIAN_TEXT = /[぀-ヿ㐀-鿿가-힯]/;

/** Whether text contains characters drawn with the East Asian font. */
export function hasEastAsianText(text: string): boolean {
  return EAST_ASIAN_TEXT.test(text);
}
