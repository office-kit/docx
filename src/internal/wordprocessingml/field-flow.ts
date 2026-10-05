/**
 * The document's paragraph flow and the complex fields in it (ECMA-376 Part 1
 * §17.16.18): which runs open, separate and close each field, how to replace a
 * field's result (inline, or a run of whole paragraphs for TOC / INDEX / TOA /
 * BIBLIOGRAPHY), and an estimate of the page each paragraph lands on.
 */

import { parseXml, serializeXml, type XmlElement, type XmlNode } from "../xml/index.js";

/** Replace an element's children in place (the XML AST types them read-only). */
export function replaceChildren(el: XmlElement, nodes: readonly XmlNode[]): void {
  const children = el.children as XmlNode[];
  children.splice(0, children.length, ...nodes);
}
import { WML_NS } from "./namespaces.js";
import { parseParagraph, parseWmlDocument } from "./parser.js";
import type { WmlBlock, WmlBody, WmlInline, WmlParagraph, WmlRun, WmlRunPiece } from "./types.js";
import { paragraphToElement } from "./writer.js";
import { type FieldInstruction, parseFieldInstruction } from "./field-code.js";

const W_DECL = `xmlns:w="${WML_NS}"`;
const R_DECL = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

/** Escape text for an XML text node or attribute value. */
export function escapeXml(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

/** `<w:r>` XML for plain text; `\t` becomes `<w:tab/>` and `\n` `<w:br/>`. */
export function textRunXml(text: string, rPr = ""): string {
  if (text === "") return "";
  const parts = text.split(/(\t|\n)/);
  const inner = parts
    .map((part) => {
      if (part === "\t") return "<w:tab/>";
      if (part === "\n") return "<w:br/>";
      if (part === "") return "";
      const space = /^\s|\s$/.test(part) ? ' xml:space="preserve"' : "";
      return `<w:t${space}>${escapeXml(part)}</w:t>`;
    })
    .join("");
  return `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ""}${inner}</w:r>`;
}

/**
 * The runs of a complex field: begin, instruction, separate, the result runs,
 * end. Word pads the instruction with spaces, and so do we.
 */
export function fieldRunsXml(instruction: string, resultXml: string, rPr = ""): string {
  const props = rPr ? `<w:rPr>${rPr}</w:rPr>` : "";
  return [
    `<w:r>${props}<w:fldChar w:fldCharType="begin"/></w:r>`,
    `<w:r>${props}<w:instrText xml:space="preserve"> ${escapeXml(instruction)} </w:instrText></w:r>`,
    `<w:r>${props}<w:fldChar w:fldCharType="separate"/></w:r>`,
    resultXml,
    `<w:r>${props}<w:fldChar w:fldCharType="end"/></w:r>`,
  ].join("");
}

/**
 * The skeleton of a field whose result spans paragraphs (TOC, INDEX …): a
 * paragraph opening the field up to its separator, and a paragraph holding
 * the end character. `pPr` goes on the first paragraph.
 */
export function blockFieldSkeletonXml(instruction: string, pPr = ""): string {
  return [
    `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ""}`,
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r>',
    `<w:r><w:instrText xml:space="preserve"> ${escapeXml(instruction)} </w:instrText></w:r>`,
    '<w:r><w:fldChar w:fldCharType="separate"/></w:r>',
    "</w:p>",
    '<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>',
  ].join("");
}

/** A field without a result (XE, TA, NEXT …): begin, instruction, end. */
export function resultlessFieldRunsXml(instruction: string): string {
  return [
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r>',
    `<w:r><w:instrText xml:space="preserve"> ${escapeXml(instruction)} </w:instrText></w:r>`,
    '<w:r><w:fldChar w:fldCharType="end"/></w:r>',
  ].join("");
}

/** Parse `<w:p>` elements (as XML text) into typed paragraphs. */
export function paragraphsFromXml(xml: string): WmlParagraph[] {
  const root = parseXml(`<w:body ${W_DECL} ${R_DECL}>${xml}</w:body>`).root;
  return root.children
    .filter((c): c is XmlElement => c.kind === "element" && c.name.local === "p")
    .map((p) => parseParagraph(p));
}

/** Parse body content (paragraphs and tables, as XML text) into typed blocks. */
export function blocksFromXml(xml: string): WmlBlock[] {
  const doc = parseXml(
    `<w:document ${W_DECL} ${R_DECL}><w:body>${xml}</w:body></w:document>`,
  );
  return parseWmlDocument(doc).body.blocks;
}

/** Serialize one element (and its subtree) to XML text. */
export function elementXml(el: XmlElement): string {
  return serializeXml({ prologue: [], root: el, epilogue: [] });
}

/** Parse one `w:`-prefixed element from XML text (no namespace declaration needed). */
export function elementFromXml(xml: string): XmlElement {
  const root = parseXml(`<w:wrap ${W_DECL} ${R_DECL}>${xml}</w:wrap>`).root;
  const el = root.children.find((c): c is XmlElement => c.kind === "element");
  if (!el) throw new Error("elementFromXml: no element in input");
  return el;
}

/** Parse run-level XML (`<w:r>`, `<w:hyperlink>` …) into typed inlines. */
export function inlinesFromXml(xml: string): WmlInline[] {
  return paragraphsFromXml(`<w:p>${xml}</w:p>`)[0]?.children ?? [];
}

/** A paragraph in the flow, and the array (body blocks or cell paragraphs) that holds it. */
export interface ParagraphRef {
  readonly list: WmlBlock[] | WmlParagraph[];
  readonly index: number;
  readonly paragraph: WmlParagraph;
}

/** Where a field character sits: its paragraph and the index of its run. */
export interface FieldPosition {
  readonly ref: ParagraphRef;
  readonly inline: number;
}

/** A complex field or `<w:fldSimple>` found in the flow. */
export interface FlowField {
  readonly instruction: string;
  readonly parsed: FieldInstruction;
  /** 0 for a top-level field; nested fields count their enclosing ones. */
  readonly depth: number;
  readonly begin: FieldPosition;
  readonly separate: FieldPosition | undefined;
  readonly end: FieldPosition;
  /** Set for `<w:fldSimple>`: the element, whose children are the result. */
  readonly simple?: XmlElement;
}

function isFieldCharRun(inline: WmlInline): inline is WmlRun {
  return inline.kind === "run" && inline.pieces.some((p) => p.kind === "fieldChar");
}

/**
 * Bring a paragraph into the shape the field scanner reads: runs built as raw
 * XML become typed runs, and a run mixing a field character with other content
 * is split so each field character has a run of its own. The XML is unchanged
 * in meaning.
 */
export function normalizeFieldRuns(paragraph: WmlParagraph): void {
  const out: WmlInline[] = [];
  let changed = false;
  for (const inline of paragraph.children) {
    let item = inline;
    if (item.kind === "raw" && item.node.name.uri === WML_NS && item.node.name.local === "r") {
      const wrapped = parseParagraph({
        kind: "element",
        name: { uri: WML_NS, local: "p", prefix: "w" },
        attrs: [],
        children: [item.node],
        xmlSpace: "default",
        selfClosing: false,
      });
      const run = wrapped.children[0];
      if (run) {
        item = run;
        changed = true;
      }
    }
    if (item.kind === "run" && isFieldCharRun(item) && item.pieces.length > 1) {
      changed = true;
      const run: WmlRun = item;
      let group: WmlRunPiece[] = [];
      const flush = (): void => {
        if (group.length === 0) return;
        // Split runs share the original's formatting but not its unknown
        // children or identity attributes.
        out.push({ kind: "run", ...(run.rPr ? { rPr: run.rPr } : {}), pieces: group, extras: [] });
        group = [];
      };
      for (const piece of run.pieces) {
        if (piece.kind === "fieldChar") {
          flush();
          group = [piece];
          flush();
        } else {
          group.push(piece);
        }
      }
      flush();
      continue;
    }
    out.push(item);
  }
  if (changed) paragraph.children = out;
}

const SDT_CONTENT = "sdtContent";

/**
 * Body-level content controls (`<w:sdt>`), which hold Word's Table of
 * Contents and Bibliography, are raw blocks in the typed model. Their
 * paragraphs are opened as typed paragraphs for the duration of `fn`, then
 * written back, so fields inside them are found and updated like any other.
 */
export function withOpenContentControls<T>(body: WmlBody, fn: () => T): T {
  const opened: Array<{ content: XmlElement; paragraphs: WmlParagraph[]; at: number }> = [];
  for (let i = 0; i < body.blocks.length; i++) {
    const block = body.blocks[i] as WmlBlock;
    if (block.kind !== "raw" || block.node.name.local !== "sdt") continue;
    const content = block.node.children.find(
      (c): c is XmlElement => c.kind === "element" && c.name.local === SDT_CONTENT,
    );
    if (!content) continue;
    const elements = content.children.filter((c): c is XmlElement => c.kind === "element");
    if (elements.length === 0 || !elements.every((c) => c.name.local === "p")) continue;
    opened.push({ content, paragraphs: elements.map((p) => parseParagraph(p)), at: i });
  }
  if (opened.length === 0) return fn();
  // Splice the opened paragraphs in place of their content control between
  // two marker paragraphs, run, and put the (possibly changed) paragraphs
  // back inside it.
  const starts = new Map<WmlParagraph, (typeof opened)[number] & { sdt: WmlBlock }>();
  const ends = new Set<WmlParagraph>();
  const flat: WmlBlock[] = [];
  let next = 0;
  body.blocks.forEach((block, i) => {
    const open = opened[next];
    if (open && open.at === i) {
      next++;
      const start: WmlParagraph = { kind: "paragraph", children: [], extras: [] };
      const stop: WmlParagraph = { kind: "paragraph", children: [], extras: [] };
      starts.set(start, { ...open, sdt: block });
      ends.add(stop);
      flat.push(start, ...open.paragraphs, stop);
    } else {
      flat.push(block);
    }
  });
  body.blocks = flat;
  try {
    return fn();
  } finally {
    const rebuilt: WmlBlock[] = [];
    let control: ((typeof opened)[number] & { sdt: WmlBlock }) | undefined;
    let inner: WmlParagraph[] = [];
    for (const block of body.blocks) {
      const opening = block.kind === "paragraph" ? starts.get(block) : undefined;
      if (opening) {
        control = opening;
        inner = [];
      } else if (control && block.kind === "paragraph" && ends.has(block)) {
        replaceChildren(
          control.content,
          inner.map((p) => paragraphToElement(p)),
        );
        // A control whose content was removed entirely (Remove Table of
        // Contents) goes with it.
        if (inner.length > 0) rebuilt.push(control.sdt);
        control = undefined;
      } else if (control) {
        if (block.kind === "paragraph") inner.push(block);
      } else {
        rebuilt.push(block);
      }
    }
    body.blocks = rebuilt;
  }
}

/** Every paragraph in document order: body paragraphs and table cell paragraphs. */
export function paragraphFlow(body: WmlBody): ParagraphRef[] {
  const out: ParagraphRef[] = [];
  body.blocks.forEach((block, index) => {
    if (block.kind === "paragraph") out.push({ list: body.blocks, index, paragraph: block });
    else if (block.kind === "table") {
      for (const row of block.rows) {
        for (const cell of row.cells) {
          cell.paragraphs.forEach((paragraph, i) =>
            out.push({ list: cell.paragraphs, index: i, paragraph }),
          );
        }
      }
    }
  });
  return out;
}

/** Concatenated `<w:instrText>` of a run. */
function instrTextOf(run: WmlRun): string {
  let acc = "";
  for (const piece of run.pieces) if (piece.kind === "instrText") acc += piece.value;
  return acc;
}

/**
 * Every field in the flow in document order (by where it begins), with
 * nesting depth. Normalizes the paragraphs' field runs first.
 */
export function scanFields(flow: readonly ParagraphRef[]): FlowField[] {
  interface Open {
    begin: FieldPosition;
    instruction: string;
    separate: FieldPosition | undefined;
    order: number;
  }
  const stack: Open[] = [];
  const done: Array<FlowField & { order: number }> = [];
  let order = 0;
  for (const ref of flow) {
    normalizeFieldRuns(ref.paragraph);
    ref.paragraph.children.forEach((inline, i) => {
      if (inline.kind === "raw") {
        const node = inline.node;
        if (node.name.uri === WML_NS && node.name.local === "fldSimple") {
          const instruction = node.attrs.find((a) => a.name.local === "instr")?.value ?? "";
          const at = { ref, inline: i };
          done.push({
            instruction,
            parsed: parseFieldInstruction(instruction),
            depth: stack.length,
            begin: at,
            separate: at,
            end: at,
            simple: node,
            order: order++,
          });
        }
        return;
      }
      for (const piece of inline.pieces) {
        if (piece.kind === "fieldChar") {
          const at = { ref, inline: i };
          if (piece.charType === "begin") {
            stack.push({ begin: at, instruction: "", separate: undefined, order: order++ });
          } else if (piece.charType === "separate") {
            const top = stack.at(-1);
            if (top) top.separate = at;
          } else {
            const top = stack.pop();
            if (top) {
              done.push({
                instruction: top.instruction,
                parsed: parseFieldInstruction(top.instruction),
                depth: stack.length,
                begin: top.begin,
                separate: top.separate,
                end: at,
                order: top.order,
              });
            }
          }
        } else if (piece.kind === "instrText") {
          const top = stack.at(-1);
          // Instruction text after the separator belongs to a nested field's
          // result, not to this field's code.
          if (top && !top.separate) top.instruction += instrTextOf({ ...inline, pieces: [piece] });
        }
      }
    });
  }
  return done.toSorted((a, b) => a.order - b.order);
}

/** The visible text of a field's current result. */
export function fieldResultText(field: FlowField): string {
  if (field.simple) return xmlText(field.simple);
  if (!field.separate) return "";
  const { ref, inline } = field.separate;
  if (ref !== field.end.ref) {
    // Multi-paragraph results: their text is not needed by any caller that
    // reads a cached value.
    return "";
  }
  let acc = "";
  for (const child of ref.paragraph.children.slice(inline + 1, field.end.inline)) {
    if (child.kind === "run") {
      for (const p of child.pieces) {
        if (p.kind === "text") acc += p.value;
        else if (p.kind === "tab") acc += "\t";
      }
    } else acc += xmlText(child.node);
  }
  return acc;
}

function xmlText(el: XmlElement): string {
  if (el.name.uri === WML_NS && el.name.local === "t") {
    return el.children.map((c) => (c.kind === "text" ? c.value : "")).join("");
  }
  if (el.name.uri === WML_NS && el.name.local === "instrText") return "";
  return el.children.map((c) => (c.kind === "element" ? xmlText(c) : "")).join("");
}

const SEPARATE_RUN_XML = '<w:r><w:fldChar w:fldCharType="separate"/></w:r>';

/**
 * Replace a field's result with inline content. Valid when the field begins
 * and ends in one paragraph (or is a `<w:fldSimple>`).
 */
export function setInlineResult(field: FlowField, inlines: WmlInline[]): void {
  if (field.simple) {
    replaceChildren(
      field.simple,
      inlines.map((inline) => (inline.kind === "raw" ? inline.node : runElement(inline))),
    );
    return;
  }
  const paragraph = field.end.ref.paragraph;
  if (field.separate && field.separate.ref === field.end.ref) {
    const from = field.separate.inline + 1;
    paragraph.children.splice(from, field.end.inline - from, ...inlines);
    return;
  }
  paragraph.children.splice(field.end.inline, 0, ...inlinesFromXml(SEPARATE_RUN_XML), ...inlines);
}

function runElement(run: WmlRun): XmlElement {
  const p = paragraphToElement({ kind: "paragraph", children: [run], extras: [] });
  return p.children.find((c): c is XmlElement => c.kind === "element") as XmlElement;
}

/**
 * Replace a field's result with whole paragraphs, the way Word lays out a TOC,
 * INDEX, TOA or BIBLIOGRAPHY: the field code opens the first result paragraph
 * and the end character sits in a paragraph of its own after the last. Text
 * before the field and after its end is kept. Both ends must share one
 * paragraph list (body or a single cell).
 */
export function setBlockResult(field: FlowField, paragraphs: WmlParagraph[]): void {
  const begin = field.begin.ref;
  const end = field.end.ref;
  if (begin.list !== end.list) {
    throw new Error(`The ${field.parsed.type} field spans more than one table cell.`);
  }
  const codeEnd = field.separate && field.separate.ref === begin ? field.separate.inline + 1 : -1;
  const code =
    codeEnd >= 0
      ? begin.paragraph.children.slice(field.begin.inline, codeEnd)
      : [
          ...begin.paragraph.children.slice(
            field.begin.inline,
            begin === end ? field.end.inline : begin.paragraph.children.length,
          ),
          ...inlinesFromXml(SEPARATE_RUN_XML),
        ];
  const prefix = begin.paragraph.children.slice(0, field.begin.inline);
  const endRun = end.paragraph.children[field.end.inline] as WmlInline;
  const suffix = end.paragraph.children.slice(field.end.inline + 1);
  const first = paragraphs[0] ?? { kind: "paragraph", children: [], extras: [] };
  first.children = [...prefix, ...code, ...first.children];
  const last: WmlParagraph = {
    kind: "paragraph",
    ...(end.paragraph.pPr ? { pPr: end.paragraph.pPr } : {}),
    children: [endRun, ...suffix],
    extras: [],
  };
  const list = begin.list as WmlParagraph[];
  list.splice(begin.index, end.index - begin.index + 1, first, ...paragraphs.slice(1), last);
}

/** Remove a field and its result entirely (Remove Table of Contents). */
export function removeField(field: FlowField): void {
  const begin = field.begin.ref;
  const end = field.end.ref;
  if (begin === end) {
    begin.paragraph.children.splice(field.begin.inline, field.end.inline - field.begin.inline + 1);
    return;
  }
  const list = begin.list as WmlParagraph[];
  const prefix = begin.paragraph.children.slice(0, field.begin.inline);
  const suffix = end.paragraph.children.slice(field.end.inline + 1);
  const keep: WmlParagraph[] = [];
  if (prefix.length > 0) keep.push({ ...begin.paragraph, children: prefix });
  if (suffix.length > 0 || end.paragraph.pPr?.children.some(isSectPr)) {
    keep.push({ ...end.paragraph, children: suffix });
  }
  list.splice(begin.index, end.index - begin.index + 1, ...keep);
}

function isSectPr(node: { kind: string }): boolean {
  return node.kind === "element" && (node as XmlElement).name.local === "sectPr";
}

/** Look up a child element of a properties element by local name. */
export function childElement(parent: XmlElement | undefined, local: string): XmlElement | undefined {
  return parent?.children.find(
    (c): c is XmlElement => c.kind === "element" && c.name.uri === WML_NS && c.name.local === local,
  );
}

/** A `w:`-namespaced attribute value. */
export function wAttr(el: XmlElement | undefined, local: string): string | undefined {
  return el?.attrs.find((a) => a.name.local === local)?.value;
}

/** The pages a section break of each `w:type` starts (§17.18.77); continuous adds none. */
const NEW_PAGE_SECTION_TYPES: ReadonlySet<string> = new Set(["nextPage", "oddPage", "evenPage"]);

/**
 * Estimate the page each paragraph starts on from the breaks the document
 * itself records: page breaks, "page break before", non-continuous section
 * breaks, and the `<w:lastRenderedPageBreak>` marks Word leaves where its own
 * layout broke pages. Without a layout engine this cannot see pages that
 * overflow; callers with real pagination pass their own page lookup instead.
 */
export function estimatePages(body: WmlBody): Map<WmlParagraph, number> {
  const pages = new Map<WmlParagraph, number>();
  const flow = paragraphFlow(body);
  const sectionTypes: string[] = [];
  for (const ref of flow) {
    const sectPr = childElement(ref.paragraph.pPr, "sectPr");
    if (sectPr) sectionTypes.push(wAttr(childElement(sectPr, "type"), "val") ?? "nextPage");
  }
  sectionTypes.push(wAttr(childElement(body.sectPr, "type"), "val") ?? "nextPage");
  let page = 1;
  // True at the top of a page, so a rendered-break mark right after an
  // explicit break does not count the same page twice.
  let fresh = true;
  let section = 0;
  const newPage = (): void => {
    if (!fresh) page++;
    fresh = true;
  };
  for (const ref of flow) {
    const pPr = ref.paragraph.pPr;
    const before = childElement(pPr, "pageBreakBefore");
    if (before && wAttr(before, "val") !== "0" && wAttr(before, "val") !== "false") newPage();
    // A paragraph is on the page where its first content lands, so a
    // rendered-break mark before any text moves the paragraph itself.
    let placed = false;
    const place = (): void => {
      if (!placed) pages.set(ref.paragraph, page);
      placed = true;
    };
    for (const inline of ref.paragraph.children) {
      if (inline.kind !== "run") {
        if (inline.node.name.local !== "bookmarkStart" && inline.node.name.local !== "bookmarkEnd") {
          place();
          fresh = false;
        }
        continue;
      }
      for (const piece of inline.pieces) {
        if (piece.kind === "break" && piece.breakType === "page") {
          place();
          fresh = false;
          newPage();
        } else if (piece.kind === "lastRenderedPageBreak") {
          newPage();
        } else if (piece.kind !== "fieldChar" && piece.kind !== "instrText") {
          place();
          fresh = false;
        }
      }
    }
    place();
    if (childElement(pPr, "sectPr")) {
      section++;
      if (NEW_PAGE_SECTION_TYPES.has(sectionTypes[section] ?? "nextPage")) newPage();
    }
  }
  return pages;
}

/** The plain text of a paragraph's runs and hyperlinks, without field codes. */
export function visibleParagraphText(paragraph: WmlParagraph): string {
  return visibleInlinesText(paragraph.children);
}

/** The visible text of a slice of a paragraph's inlines. */
export function visibleInlinesText(inlines: readonly WmlInline[]): string {
  let acc = "";
  // Field codes (between begin and separate) are not visible text.
  const inCode: boolean[] = [];
  for (const inline of inlines) {
    if (inline.kind === "raw") {
      if (!inCode.at(-1)) acc += xmlText(inline.node);
      continue;
    }
    for (const piece of inline.pieces) {
      if (piece.kind === "fieldChar") {
        if (piece.charType === "begin") {
          inCode.push(true);
        } else if (piece.charType === "separate") {
          inCode[inCode.length - 1] = false;
        } else {
          inCode.pop();
        }
        continue;
      }
      if (inCode.at(-1)) continue;
      if (piece.kind === "text") acc += piece.value;
      else if (piece.kind === "tab") acc += "\t";
    }
  }
  return acc;
}
