/**
 * Review markup on the canvas: tracked insertions and deletions, changed
 * paragraphs (for the change bar), comment ranges, editable exceptions of a
 * protected document, proofing language and the Outline view's levels, as classes and `data-*` attributes on the rendered
 * runs and paragraphs.
 *
 * Which of them show — Simple Markup, All Markup, No Markup, Original — is
 * left to the page's CSS (keyed on the markup mode), so switching the
 * display does not re-render the document.
 */

import {
  type Docx,
  getRunLanguage,
  getRunProp,
  revisions,
  type WmlParagraph,
  type WmlRun,
  type XmlElement,
} from "@office-kit/docx";
import { outlineLevelOf } from "./outline.js";
import { WML_NS } from "./wml-ns.js";

/** How many author colours the canvas cycles through (Word's "By author"). */
export const AUTHOR_COLORS = 8;

export interface RunDecoration {
  /** Extra classes for the run's span (leading space included). */
  readonly classes: string;
  /** Extra attributes for the run's span (leading space included). */
  readonly attrs: string;
  /** A deleted run: shown read-only and never synced back as typed text. */
  readonly deleted: boolean;
}

export interface ReviewDecorations {
  /** Call for every inline of a paragraph, in order: tracks open comment ranges. */
  inline(node: XmlElement): void;
  run(run: WmlRun): RunDecoration;
  paragraph(para: WmlParagraph): { readonly classes: string; readonly attrs: string };
}

function wAttr(el: { readonly attrs: XmlElement["attrs"] }, local: string): string | undefined {
  return el.attrs.find((a) => a.name.uri === WML_NS && a.name.local === local)?.value;
}

function escapeAttr(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;");
}

function hasChild(el: XmlElement | undefined, local: string): boolean {
  return !!el?.children.some(
    (c) => c.kind === "element" && c.name.uri === WML_NS && c.name.local === local,
  );
}

function markRPr(para: WmlParagraph): XmlElement | undefined {
  return para.pPr?.children.find(
    (c): c is XmlElement => c.kind === "element" && c.name.uri === WML_NS && c.name.local === "rPr",
  );
}

/**
 * Decorations for one render of `doc`. Author colours follow the order in
 * which authors first appear, as Word assigns them.
 */
export function reviewDecorations(doc: Docx): ReviewDecorations {
  const authors = new Map<string, number>();
  for (const r of revisions(doc)) {
    if (r.author !== undefined && !authors.has(r.author)) authors.set(r.author, authors.size);
  }
  const colorOf = (author: string | undefined): number =>
    (authors.get(author ?? "") ?? 0) % AUTHOR_COLORS;
  const openComments = new Set<string>();
  // Exception ranges stay editable while protection locks the rest.
  const openPermissions = new Set<string>();

  return {
    inline(node) {
      if (node.name.uri !== WML_NS) return;
      const id = wAttr(node, "id");
      if (id === undefined) return;
      if (node.name.local === "commentRangeStart") openComments.add(id);
      else if (node.name.local === "commentRangeEnd") openComments.delete(id);
      else if (node.name.local === "permStart") openPermissions.add(id);
      else if (node.name.local === "permEnd") openPermissions.delete(id);
    },
    run(run) {
      const classes: string[] = [];
      const attrs: string[] = [];
      const revision = run.revision;
      if (revision) {
        const author = wAttr(revision, "author");
        classes.push(revision.kind === "ins" ? "wk-ins" : "wk-del", `wk-author-${colorOf(author)}`);
        const title = [author, wAttr(revision, "date")?.slice(0, 10)].filter(Boolean).join(" · ");
        if (title) attrs.push(`title="${escapeAttr(title)}"`);
        attrs.push(`data-wk-rev="${escapeAttr(wAttr(revision, "id") ?? "")}"`);
      }
      if (hasChild(run.rPr, "rPrChange")) classes.push("wk-fmt-change");
      if (openComments.size > 0) {
        classes.push("wk-commented");
        attrs.push(`data-wk-comment="${escapeAttr([...openComments].join(" "))}"`);
      }
      if (openPermissions.size > 0) {
        classes.push("wk-perm");
        attrs.push('contenteditable="true"');
      }
      const language = getRunLanguage(run).latin;
      if (language) attrs.push(`lang="${escapeAttr(language)}"`);
      if (getRunProp(run, "noProof").present && getRunProp(run, "noProof").val !== "0") {
        attrs.push('spellcheck="false"');
      }
      return {
        classes: classes.map((c) => ` ${c}`).join(""),
        attrs: attrs.map((a) => ` ${a}`).join(""),
        deleted: revision?.kind === "del",
      };
    },
    paragraph(para) {
      const mark = markRPr(para);
      const changed =
        hasChild(para.pPr, "pPrChange") ||
        hasChild(mark, "ins") ||
        hasChild(mark, "del") ||
        para.children.some(
          (c) => c.kind === "run" && (c.revision !== undefined || hasChild(c.rPr, "rPrChange")),
        );
      const classes = [
        changed ? " wk-changed" : "",
        hasChild(mark, "del") ? " wk-mark-del" : "",
        hasChild(mark, "ins") ? " wk-mark-ins" : "",
      ].join("");
      return { classes, attrs: ` data-wk-outline="${outlineLevelOf(para)}"` };
    },
  };
}

/** The visible text of a deleted run (its `w:delText`). */
export function deletedRunText(run: WmlRun): string {
  let out = "";
  for (const piece of run.pieces) {
    if (piece.kind === "delText") out += piece.value;
    else if (piece.kind === "tab") out += "\t";
  }
  return out;
}
