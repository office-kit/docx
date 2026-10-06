/**
 * Read clipboard HTML — from Word, a web page, Google Docs, or this editor's
 * own canvas — into {@link PastedBlock}s for `insertFragmentCommand`.
 * Browser-only: it parses with `DOMParser`.
 *
 * What survives: paragraphs, line breaks and tabs; bold, italic, underline,
 * strikethrough, superscript / subscript and text color; headings; bullet and
 * numbered lists with their levels; and tables. Fonts, sizes, spacing and
 * other layout are left to the destination, as in Word's "Merge Formatting".
 */

import type {
  PastedBlock,
  PastedParagraph,
  PastedRun,
  PastedRunFormat,
} from "./commands/structure.js";

type ListSpec = NonNullable<PastedParagraph["list"]>;

interface Draft {
  runs: PastedRun[];
  heading?: number;
  list?: ListSpec;
}

interface Context {
  readonly format: PastedRunFormat;
  readonly pre: boolean;
  /** The kind and depth of the `<ul>` / `<ol>` being read. */
  readonly list?: ListSpec;
}

const SKIPPED = new Set([
  "SCRIPT",
  "STYLE",
  "HEAD",
  "TITLE",
  "META",
  "LINK",
  "TEMPLATE",
  "NOSCRIPT",
]);
const BLOCKS = new Set([
  "P",
  "DIV",
  "LI",
  "BLOCKQUOTE",
  "PRE",
  "SECTION",
  "ARTICLE",
  "HEADER",
  "FOOTER",
  "ASIDE",
  "NAV",
  "MAIN",
  "FIGURE",
  "FIGCAPTION",
  "ADDRESS",
  "DT",
  "DD",
  "H1",
  "H2",
  "H3",
  "H4",
  "H5",
  "H6",
]);
const HEADING = /^H([1-6])$/;
// Word writes a list paragraph as `mso-list:l0 level2 lfo1`, its marker in a
// span styled `mso-list:Ignore`.
const MSO_LIST_LEVEL = /mso-list:\s*l\d+\s+level(\d+)/i;
const MSO_LIST_IGNORE = /mso-list:\s*ignore/i;
// A numbered marker ends in a delimiter ("1.", "a)", "(iv)") or is a circled
// number ("①"); a bare symbol or letter ("•", "o", "§") is a bullet.
const NUMBERED_MARKER = /^\(?[\p{L}\p{N}]+[.)]$|^\p{No}$/u;
const ZWSP = /​/g;
const MAX_LIST_LEVEL = 8;

export function parseClipboardHtml(html: string): PastedBlock[] {
  const doc = new DOMParser().parseFromString(html, "text/html");
  return readBlocks(doc.body, { format: {}, pre: false });
}

/** The blocks of a container: the clipboard body or a table cell. */
function readBlocks(root: Element, ctx: Context): PastedBlock[] {
  const out: PastedBlock[] = [];
  let draft: Draft | undefined;

  const flush = (): void => {
    if (!draft) return;
    const runs = trimParagraph(draft.runs);
    if (runs.length > 0 || draft.list || draft.heading !== undefined) {
      out.push({
        kind: "paragraph",
        runs,
        ...(draft.heading !== undefined ? { heading: draft.heading } : {}),
        ...(draft.list ? { list: draft.list } : {}),
      });
    }
    draft = undefined;
  };

  /** Start a paragraph; an opened one with nothing in it yet takes the props instead. */
  const open = (props: Omit<Draft, "runs">): void => {
    if (draft && draft.runs.length === 0) {
      Object.assign(draft, props);
      return;
    }
    flush();
    draft = { runs: [], ...props };
  };

  const text = (value: string, format: PastedRunFormat): void => {
    draft ??= { runs: [] };
    const last = draft.runs.at(-1);
    if (last && sameFormat(last.format, format)) {
      draft.runs[draft.runs.length - 1] = { text: last.text + value, format };
    } else {
      draft.runs.push({ text: value, format });
    }
  };

  const walk = (node: Node, ctx: Context): void => {
    if (node.nodeType === Node.TEXT_NODE) {
      const raw = (node.textContent ?? "").replace(ZWSP, "");
      // HTML collapses ASCII whitespace only; a no-break space stays a space.
      const value = ctx.pre ? raw.replace(/\r\n?/g, "\n") : raw.replace(/[ \t\n\r\f]+/g, " ");
      if (value) text(value, ctx.format);
      return;
    }
    if (!(node instanceof HTMLElement) || SKIPPED.has(node.tagName) || isIgnored(node)) return;
    const tag = node.tagName;
    const style = node.getAttribute("style") ?? "";
    const inner: Context = {
      ...ctx,
      format: formatOf(node, ctx.format),
      pre: ctx.pre || tag === "PRE" || /white-space:\s*pre/i.test(style),
    };
    if (tag === "BR") {
      text("\n", ctx.format);
      return;
    }
    if (node.classList.contains("wk-tab")) {
      text("\t", ctx.format);
      return;
    }
    if (tag === "TABLE") {
      flush();
      out.push({ kind: "table", rows: readTable(node, inner) });
      return;
    }
    if (tag === "UL" || tag === "OL") {
      flush();
      const level = ctx.list ? Math.min(ctx.list.level + 1, MAX_LIST_LEVEL) : 0;
      const list: ListSpec = { kind: tag === "OL" ? "numbered" : "bullet", level };
      for (const child of Array.from(node.childNodes)) walk(child, { ...inner, list });
      flush();
      return;
    }
    if (BLOCKS.has(tag)) {
      const heading = HEADING.exec(tag)?.[1];
      const list = tag === "LI" ? ctx.list : (msoList(node, style) ?? editorList(node));
      open({
        ...(heading ? { heading: Number(heading) } : {}),
        ...(list ? { list } : {}),
      });
      for (const child of Array.from(node.childNodes)) walk(child, inner);
      flush();
      return;
    }
    for (const child of Array.from(node.childNodes)) walk(child, inner);
  };

  for (const child of Array.from(root.childNodes)) walk(child, ctx);
  flush();
  return out;
}

function readTable(table: HTMLElement, ctx: Context): PastedParagraph[][][] {
  // Rows of this table only: a nested table's rows belong to its cell.
  const rows = Array.from(table.querySelectorAll("tr")).filter(
    (tr) => tr.closest("table") === table,
  );
  return rows.map((tr) =>
    Array.from(tr.children)
      .filter(
        (cell): cell is HTMLElement =>
          cell instanceof HTMLElement && (cell.tagName === "TD" || cell.tagName === "TH"),
      )
      .map((cell) =>
        // A cell holds paragraphs; a table nested in it is read as its text.
        readBlocks(cell, { format: formatOf(cell, ctx.format), pre: ctx.pre }).flatMap((b) =>
          b.kind === "paragraph" ? [b] : b.rows.flat(2),
        ),
      ),
  );
}

/** Word's list marker spans, and generated text in this editor's own copy. */
function isIgnored(el: HTMLElement): boolean {
  return (
    MSO_LIST_IGNORE.test(el.getAttribute("style") ?? "") ||
    el.classList.contains("wk-list-label") ||
    el.classList.contains("wk-fchar") ||
    el.classList.contains("wk-fcode")
  );
}

/** A Word list paragraph's level and kind (read off its marker). */
function msoList(el: HTMLElement, style: string): ListSpec | undefined {
  const level = MSO_LIST_LEVEL.exec(style)?.[1];
  if (!level) return undefined;
  const marker = Array.from(el.querySelectorAll("span")).find((s) =>
    MSO_LIST_IGNORE.test(s.getAttribute("style") ?? ""),
  );
  return {
    kind: markerKind(marker?.textContent ?? ""),
    level: Math.min(Math.max(Number(level) - 1, 0), MAX_LIST_LEVEL),
  };
}

/** A list paragraph copied from this editor: it carries its rendered label. */
function editorList(el: HTMLElement): ListSpec | undefined {
  const label = el.querySelector(":scope > .wk-list-label");
  return label ? { kind: markerKind(label.textContent ?? ""), level: 0 } : undefined;
}

function markerKind(marker: string): ListSpec["kind"] {
  return NUMBERED_MARKER.test(marker.replace(/\s/g, "")) ? "numbered" : "bullet";
}

/** The formatting inside `el`: what its tag implies, then its inline style. */
function formatOf(el: HTMLElement, outer: PastedRunFormat): PastedRunFormat {
  const f: { -readonly [K in keyof PastedRunFormat]: PastedRunFormat[K] } = { ...outer };
  switch (el.tagName) {
    case "B":
    case "STRONG":
    case "TH":
      f.bold = true;
      break;
    case "I":
    case "EM":
    case "CITE":
    case "VAR":
      f.italic = true;
      break;
    case "U":
    case "INS":
      f.underline = true;
      break;
    case "S":
    case "STRIKE":
    case "DEL":
      f.strike = true;
      break;
    case "SUP":
      f.verticalAlign = "superscript";
      break;
    case "SUB":
      f.verticalAlign = "subscript";
      break;
  }
  const attrColor = el.tagName === "FONT" ? el.getAttribute("color") : null;
  const color = hexColor(el.style.color || attrColor || "");
  if (color !== undefined) {
    if (color === BLACK) delete f.color;
    else f.color = color;
  }
  const weight = el.style.fontWeight;
  if (weight) f.bold = weight === "bold" || weight === "bolder" || Number(weight) >= BOLD_WEIGHT;
  const fontStyle = el.style.fontStyle;
  if (fontStyle) f.italic = fontStyle === "italic" || fontStyle === "oblique";
  const decoration = `${el.style.textDecorationLine} ${el.style.textDecoration}`;
  if (/underline/.test(decoration)) f.underline = true;
  if (/line-through/.test(decoration)) f.strike = true;
  if (/\bnone\b/.test(decoration) && !/underline|line-through/.test(decoration)) {
    delete f.underline;
    delete f.strike;
  }
  const align = el.style.verticalAlign;
  if (align === "super") f.verticalAlign = "superscript";
  else if (align === "sub") f.verticalAlign = "subscript";
  else if (align === "baseline") delete f.verticalAlign;
  return f;
}

const BOLD_WEIGHT = 600;
// Black is every document's default text color; keeping it would pin the
// pasted text black even in a style that colors it.
const BLACK = "000000";

/** `#rgb`, `#rrggbb` or `rgb(r, g, b)` as six hex digits; `undefined` otherwise. */
function hexColor(css: string): string | undefined {
  const value = css.trim().toLowerCase();
  const short = /^#([0-9a-f])([0-9a-f])([0-9a-f])$/.exec(value);
  if (short)
    return `${short[1]}${short[1]}${short[2]}${short[2]}${short[3]}${short[3]}`.toUpperCase();
  const long = /^#([0-9a-f]{6})$/.exec(value);
  if (long?.[1]) return long[1].toUpperCase();
  const rgb = /^rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*(?:,\s*([\d.]+)\s*)?\)$/.exec(value);
  if (!rgb || rgb[4] === "0") return undefined;
  return [rgb[1], rgb[2], rgb[3]]
    .map((n) => Math.min(Number(n), 255).toString(16).padStart(2, "0"))
    .join("")
    .toUpperCase();
}

function sameFormat(a: PastedRunFormat, b: PastedRunFormat): boolean {
  return (
    !!a.bold === !!b.bold &&
    !!a.italic === !!b.italic &&
    !!a.underline === !!b.underline &&
    !!a.strike === !!b.strike &&
    a.verticalAlign === b.verticalAlign &&
    a.color === b.color
  );
}

/**
 * HTML's whitespace rules at a paragraph's edges: collapsed spaces there do
 * not show, and neither does the `<br>` a browser leaves at the end of a block.
 */
function trimParagraph(runs: readonly PastedRun[]): PastedRun[] {
  const out = [...runs];
  const first = out[0];
  if (first) out[0] = { text: first.text.replace(/^ +/, ""), format: first.format };
  const last = out.at(-1);
  if (last) {
    out[out.length - 1] = {
      text: last.text.replace(/ +$/, "").replace(/\n$/, ""),
      format: last.format,
    };
  }
  return out
    .filter((r) => r.text !== "")
    .map((r) => ({ text: r.text.replace(NBSP, " "), format: r.format }));
}

const NBSP = /\u00a0/g;
