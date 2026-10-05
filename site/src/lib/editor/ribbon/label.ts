// Word for Mac sets a large button's label on two lines once it is more than a
// word or two ("Table of / Contents", "Insert / Footnote"); short labels and
// single words ("Comment", "Cross-reference") stay on one.
const ONE_LINE_MAX_CHARS = 7;

/** The label's lines: split at the space that balances the two halves best. */
export function largeLabelLines(label: string): readonly string[] {
  if (label.length <= ONE_LINE_MAX_CHARS || !label.includes(" ")) return [label];
  let best = -1;
  for (let i = label.indexOf(" "); i !== -1; i = label.indexOf(" ", i + 1)) {
    if (best === -1 || Math.abs(label.length - 2 * i) < Math.abs(label.length - 2 * best)) best = i;
  }
  return [label.slice(0, best), label.slice(best + 1)];
}
