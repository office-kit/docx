// Word for Mac sets a large button's label on two lines once it is more than a
// word or two ("Table of / Contents", "Insert / Footnote"); short labels and
// single words ("Comment", "Cross-reference") stay on one.
const ONE_LINE_MAX_CHARS = 7;
// Japanese has no spaces to break at, and Word's breaks follow the words
// ("改/ページ", "テキスト/の方向"), so messages mark them with a zero-width
// space. It always breaks: Word splits even a four-character label there.
const LABEL_BREAK = "\u200b";

/** The label's lines: split at the break that balances the two halves best. */
export function largeLabelLines(label: string): readonly string[] {
  const separator = label.includes(LABEL_BREAK) ? LABEL_BREAK : " ";
  if (separator === " " && (label.length <= ONE_LINE_MAX_CHARS || !label.includes(" "))) {
    return [label];
  }
  let best = -1;
  for (let i = label.indexOf(separator); i !== -1; i = label.indexOf(separator, i + 1)) {
    if (best === -1 || Math.abs(label.length - 2 * i) < Math.abs(label.length - 2 * best)) best = i;
  }
  return [label.slice(0, best), label.slice(best + 1)].map((line) =>
    line.replaceAll(LABEL_BREAK, ""),
  );
}
