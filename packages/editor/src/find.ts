/**
 * Find and Replace as Word does it: plain text with Match case / Find whole
 * words only, Word's `^` special characters (`^t`, `^p`, `^?` …), its
 * wildcard syntax (`?`, `*`, `[a-z]`, `{n,m}`, `@`, `<`, `>`), and a format
 * filter (find only bold text, text in a style …). Matches never cross a
 * paragraph mark; `^p` matches at the end of a paragraph.
 */

import {
  type Docx,
  getElementAttr,
  getParagraphStyle,
  isolateParagraphRunRange,
  type RunFormatting,
  runTextLength,
  setRunFormat,
  stylesPart,
  type WmlParagraph,
  type WmlRun,
} from "@office-kit/docx";
import { createStyleResolver, type ResolvedRunFormat, type StyleResolver } from "./resolve.js";
import type { DocPosition, Selection } from "./selection.js";

/** What the Find and Replace dialog's Format button can require of found text. */
export interface FindFormat {
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly underline?: boolean;
  readonly strike?: boolean;
  readonly font?: string;
  readonly sizeHalfPoints?: number;
  /** Hex RGB. */
  readonly color?: string;
  readonly highlight?: boolean;
  /** Paragraph style id, or the character style id of the found text. */
  readonly style?: string;
  /** Paragraph alignment (`w:jc` value). */
  readonly alignment?: string;
}

export interface FindOptions {
  readonly query: string;
  readonly matchCase?: boolean;
  readonly wholeWords?: boolean;
  readonly wildcards?: boolean;
  readonly format?: FindFormat;
}

/** A found range: one paragraph and a character range in it. */
export interface FindMatch {
  /** The paragraph's address (block, and cell / paragraph-in-cell inside a table). */
  readonly at: DocPosition;
  readonly start: number;
  readonly end: number;
  readonly text: string;
  /** The pattern's groups, for `\1` … `\9` in a wildcard replacement. */
  readonly groups: readonly string[];
}

/** Characters of a run as the canvas counts them (same unit as `runTextLength`). */
function runChars(run: WmlRun): string {
  let out = "";
  for (const piece of run.pieces) {
    if (piece.kind === "text") out += piece.value;
    else if (piece.kind === "tab") out += "\t";
    else if (piece.kind === "break") out += "\n";
    else if (piece.kind === "noBreakHyphen" || piece.kind === "softHyphen") out += "-";
  }
  return out;
}

function paragraphRuns(para: WmlParagraph): WmlRun[] {
  return para.children.filter((c): c is WmlRun => c.kind === "run");
}

/** Every paragraph of the body in reading order, with its address. */
export function bodyParagraphs(doc: Docx): Array<{ para: WmlParagraph; at: DocPosition }> {
  const out: Array<{ para: WmlParagraph; at: DocPosition }> = [];
  doc.document.body.blocks.forEach((block, index) => {
    if (block.kind === "paragraph") out.push({ para: block, at: { block: index } });
    else if (block.kind === "table") {
      block.rows.forEach((row, r) =>
        row.cells.forEach((cell, c) =>
          cell.paragraphs.forEach((para, p) =>
            out.push({ para, at: { block: index, cell: { row: r, col: c }, para: p } }),
          ),
        ),
      );
    }
  });
  return out;
}

function escapeRegExp(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

// Word's `^` codes in the Find what box (plain-text mode).
const FIND_SPECIALS: Readonly<Record<string, string>> = {
  t: "\\t",
  p: "$",
  l: "\\n",
  m: "\\n",
  "?": ".",
  "#": "\\d",
  $: "\\p{L}",
  w: "\\s+",
  "^": "\\^",
  "-": "-",
  "~": "-",
};

function plainPattern(query: string): string {
  let out = "";
  for (let i = 0; i < query.length; i++) {
    const c = query[i] ?? "";
    const next = query[i + 1];
    if (c === "^" && next !== undefined && FIND_SPECIALS[next] !== undefined) {
      out += FIND_SPECIALS[next];
      i++;
    } else {
      out += escapeRegExp(c);
    }
  }
  return out;
}

/**
 * Word's wildcard syntax as a JavaScript pattern. `*` is lazy and stays within
 * the paragraph, as in Word; `@` is "one or more of the previous item".
 */
function wildcardPattern(query: string): string {
  let out = "";
  for (let i = 0; i < query.length; i++) {
    const c = query[i] ?? "";
    switch (c) {
      case "\\": {
        const next = query[i + 1] ?? "";
        out += escapeRegExp(next);
        i++;
        break;
      }
      case "?":
        out += ".";
        break;
      case "*":
        out += ".*?";
        break;
      case "@":
        out += "+";
        break;
      case "<":
        out += "(?<![\\p{L}\\p{N}])";
        break;
      case ">":
        out += "(?![\\p{L}\\p{N}])";
        break;
      case "[": {
        const close = query.indexOf("]", i + 1);
        if (close < 0) throw new SyntaxError("A [ in the search text has no matching ].");
        const body = query.slice(i + 1, close);
        out += body.startsWith("!") ? `[^${body.slice(1)}]` : `[${body}]`;
        i = close;
        break;
      }
      case "{": {
        const close = query.indexOf("}", i + 1);
        if (close < 0) throw new SyntaxError("A { in the search text has no matching }.");
        // Word writes {n;m} with a semicolon in some locales.
        out += `{${query.slice(i + 1, close).replace(";", ",")}}`;
        i = close;
        break;
      }
      case "(":
      case ")":
        out += c;
        break;
      case "^": {
        const next = query[i + 1];
        if (next !== undefined && FIND_SPECIALS[next] !== undefined) {
          out += FIND_SPECIALS[next];
          i++;
        } else {
          out += "\\^";
        }
        break;
      }
      default:
        out += escapeRegExp(c);
    }
  }
  return out;
}

/** The regular expression a search uses. Throws a `SyntaxError` for a malformed wildcard. */
export function compileFind(options: FindOptions): RegExp {
  let pattern = options.wildcards ? wildcardPattern(options.query) : plainPattern(options.query);
  // Wildcard searches are always case sensitive in Word.
  const flags = `gu${options.matchCase || options.wildcards ? "" : "i"}`;
  if (options.wholeWords && !options.wildcards) {
    pattern = `(?<![\\p{L}\\p{N}_])(?:${pattern})(?![\\p{L}\\p{N}_])`;
  }
  return new RegExp(pattern, flags);
}

function formatMatches(
  wanted: FindFormat,
  run: ResolvedRunFormat,
  paragraphStyle: string | undefined,
  alignment: string | undefined,
): boolean {
  const underlined = run.underline !== undefined && run.underline !== "none";
  const highlighted = run.highlight !== undefined && run.highlight !== "none";
  return (
    (wanted.bold === undefined || wanted.bold === run.bold) &&
    (wanted.italic === undefined || wanted.italic === run.italic) &&
    (wanted.strike === undefined || wanted.strike === run.strike) &&
    (wanted.underline === undefined || wanted.underline === underlined) &&
    (wanted.highlight === undefined || wanted.highlight === highlighted) &&
    (wanted.font === undefined || wanted.font === run.font) &&
    (wanted.sizeHalfPoints === undefined || wanted.sizeHalfPoints === run.sizeHalfPoints) &&
    (wanted.color === undefined || wanted.color.toUpperCase() === run.color?.toUpperCase()) &&
    (wanted.alignment === undefined || wanted.alignment === (alignment ?? "left")) &&
    (wanted.style === undefined ||
      wanted.style === paragraphStyle ||
      wanted.style === run.characterStyle)
  );
}

interface RunSpan {
  readonly run: WmlRun;
  readonly start: number;
  readonly end: number;
}

function runSpans(para: WmlParagraph): { text: string; spans: RunSpan[] } {
  let text = "";
  const spans: RunSpan[] = [];
  for (const run of paragraphRuns(para)) {
    const chars = runChars(run);
    spans.push({ run, start: text.length, end: text.length + chars.length });
    text += chars;
  }
  return { text, spans };
}

/** The regex `findMatches` passes for a search by format alone (empty Find what). */
const FORMAT_ONLY = /(?:)/gu;

function hasFormat(format: FindFormat | undefined): format is FindFormat {
  return format !== undefined && Object.values(format).some((v) => v !== undefined);
}

function matchesIn(
  para: WmlParagraph,
  at: DocPosition,
  regex: RegExp,
  format: FindFormat | undefined,
  styles: StyleResolver,
  defaultStyle: string | undefined,
): FindMatch[] {
  const { text, spans } = runSpans(para);
  const out: FindMatch[] = [];
  const paraFormat = hasFormat(format) ? styles.paragraph(para) : undefined;
  const styleId = getParagraphStyle(para) ?? defaultStyle;
  if (regex === FORMAT_ONLY) {
    if (!format || !paraFormat) return out;
    // Format-only search: each stretch of runs with the format is one match.
    let open: RunSpan | undefined;
    let close = 0;
    for (const span of spans) {
      if (span.end === span.start) continue;
      if (formatMatches(format, styles.run(para, span.run), styleId, paraFormat.alignment)) {
        open ??= span;
        close = span.end;
      } else if (open) {
        out.push({
          at,
          start: open.start,
          end: close,
          text: text.slice(open.start, close),
          groups: [],
        });
        open = undefined;
      }
    }
    if (open)
      out.push({
        at,
        start: open.start,
        end: close,
        text: text.slice(open.start, close),
        groups: [],
      });
    return out;
  }
  let first = 0;
  for (const m of text.matchAll(regex)) {
    const start = m.index;
    const end = start + m[0].length;
    // A zero-length match (an anchor such as ^p alone) selects nothing.
    if (end === start) continue;
    if (format && paraFormat) {
      // Matches come in order, so the first covering span only moves forward.
      while (first < spans.length && (spans[first]?.end ?? 0) <= start) first++;
      let ok = true;
      for (let i = first; ok && i < spans.length && (spans[i]?.start ?? end) < end; i++) {
        const span = spans[i];
        if (span && span.end > span.start)
          ok = formatMatches(format, styles.run(para, span.run), styleId, paraFormat.alignment);
      }
      if (!ok) continue;
    }
    out.push({ at, start, end, text: m[0], groups: m.slice(1).map((g) => g ?? "") });
  }
  return out;
}

/** Every match in the body, in reading order. */
export function findMatches(doc: Docx, options: FindOptions): FindMatch[] {
  if (options.query === "" && !hasFormat(options.format)) return [];
  const regex = options.query === "" ? FORMAT_ONLY : compileFind(options);
  const styles = createStyleResolver(doc);
  const defaultStyle = defaultParagraphStyleId(doc);
  return bodyParagraphs(doc).flatMap(({ para, at }) =>
    matchesIn(para, at, regex, options.format, styles, defaultStyle),
  );
}

/** The style a paragraph without `w:pStyle` has (`Normal` in documents Word writes). */
function defaultParagraphStyleId(doc: Docx): string | undefined {
  const style = stylesPart(doc)?.styles.find(
    (s) => getElementAttr(s, "type") === "paragraph" && getElementAttr(s, "default") === "1",
  );
  return style && getElementAttr(style, "styleId");
}

/** A selection covering a match, for Find Next / Previous. */
export function selectionOf(doc: Docx, match: FindMatch): Selection | undefined {
  const para = bodyParagraphs(doc).find(
    (p) =>
      p.at.block === match.at.block &&
      p.at.cell?.row === match.at.cell?.row &&
      p.at.cell?.col === match.at.cell?.col &&
      (p.at.para ?? 0) === (match.at.para ?? 0),
  )?.para;
  if (!para) return undefined;
  const runs = paragraphRuns(para);
  const position = (abs: number, preferNext: boolean): DocPosition => {
    let cursor = 0;
    for (const [i, run] of runs.entries()) {
      const len = runTextLength(run);
      // A boundary between runs belongs to the next run when starting a range.
      if (abs < cursor + len || (abs === cursor + len && !preferNext)) {
        return { ...match.at, inline: i, offset: abs - cursor };
      }
      cursor += len;
    }
    const last = Math.max(runs.length - 1, 0);
    return { ...match.at, inline: last, offset: runs[last] ? runTextLength(runs[last]) : 0 };
  };
  return { anchor: position(match.start, true), focus: position(match.end, false) };
}

const REPLACE_SPECIALS = /\^[\^&tp]/g;
const REPLACE_SPECIALS_WILDCARD = /\^[\^&tp]|\\[1-9]/g;

/**
 * The replacement text for one match: `^&` is the found text, `^t` a tab,
 * `^^` a caret, and with wildcards `\1` … `\9` the groups of the pattern.
 */
export function expandReplacement(
  replacement: string,
  match: Pick<FindMatch, "text" | "groups">,
  wildcards: boolean,
): string {
  // One pass, so text from the document is never re-read as a special.
  return replacement.replace(wildcards ? REPLACE_SPECIALS_WILDCARD : REPLACE_SPECIALS, (token) => {
    if (token === "^^") return "^";
    if (token === "^&") return match.text;
    if (token === "^t") return "\t";
    if (token === "^p")
      throw new RangeError("A paragraph mark (^p) in Replace with is not supported.");
    return match.groups[Number(token.slice(1)) - 1] ?? "";
  });
}

// Leading, trailing, or doubled spaces collapse in XML unless xml:space="preserve".
const PRESERVE_SPACE = /^\s|\s$|\s\s/;

/**
 * Replace characters `start`…`end` of `para` with `text`, keeping the
 * formatting of the first replaced character (as Word does) and applying
 * `format` on top when given. Returns the run that holds the new text.
 */
export function replaceRange(
  para: WmlParagraph,
  start: number,
  end: number,
  text: string,
  format?: RunFormatting,
): WmlRun | undefined {
  const isolated = isolateParagraphRunRange(para, start, end);
  const [first, ...rest] = isolated;
  if (!first) return undefined;
  const drop = new Set<WmlRun>(rest);
  para.children = para.children.filter((c) => c.kind !== "run" || !drop.has(c));
  const pieces: WmlRun["pieces"][number][] = [];
  for (const [i, part] of text.split("\t").entries()) {
    if (i > 0) pieces.push({ kind: "tab" });
    if (part) pieces.push({ kind: "text", value: part, preserveSpace: PRESERVE_SPACE.test(part) });
  }
  first.pieces = pieces;
  if (format) setRunFormat(first, format);
  return first;
}
