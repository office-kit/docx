/**
 * Mail merge session state: the recipient list (Word reads it from the data
 * source file each time; here it is kept in memory after the user picks the
 * file), which records are included, and the previewed record.
 */

import { commands } from "@office-kit/docx-editor";
import { parseRecipientCsv, recipientInclusion, type RecipientList } from "@office-kit/docx";
import type { EditorSession } from "../session.svelte";

/** The root class that turns on Highlight Merge Fields (see MailingsTab's styles). */
const HIGHLIGHT_CLASS = "wk-highlight-merge";

/** Pick the delimiter of a recipient file: tab-separated text or CSV. */
export function parseRecipientFile(name: string, text: string): RecipientList {
  const firstLine = text.split(/\r?\n/, 1)[0] ?? "";
  const tabbed =
    /\.(txt|tsv|tab)$/i.test(name) || (firstLine.includes("\t") && !firstLine.includes(","));
  return parseRecipientCsv(text, tabbed ? "\t" : ",");
}

class MergeState {
  list = $state<RecipientList | null>(null);
  path = $state("");
  included = $state<boolean[]>([]);
  /** The previewed record, 0-based into `list.records`. */
  index = $state(0);
  preview = $state(false);
  highlight = $state(false);
  /** Answers to ASK / FILLIN prompts, by prompt. */
  answers = $state<Record<string, string>>({});

  /** Select Recipients: attach the list to the document and remember it. */
  attach(session: EditorSession, path: string, list: RecipientList): void {
    session.apply(commands.selectRecipientsCommand, { path, list });
    if (session.status || !session.model) return;
    this.list = list;
    this.path = path;
    this.index = 0;
    this.included = recipientInclusion(session.model.doc, list);
    if (this.preview) this.show(session);
  }

  /** Show the current record in the merge fields, or the «placeholders» when preview is off. */
  show(session: EditorSession): void {
    const list = this.list ?? undefined;
    session.apply(commands.previewResultsCommand, {
      list,
      index: this.preview && list ? this.index : undefined,
      options: { answers: this.answers },
    });
  }

  go(session: EditorSession, index: number): void {
    const count = this.list?.records.length ?? 0;
    if (count === 0) return;
    this.index = Math.min(Math.max(index, 0), count - 1);
    this.preview = true;
    this.show(session);
  }

  setHighlight(on: boolean): void {
    this.highlight = on;
    document.documentElement.classList.toggle(HIGHLIGHT_CLASS, on);
  }
}

export const merge = new MergeState();
