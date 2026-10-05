/**
 * References tab state that outlives one document: Word's Source Manager
 * "Master List" (kept in this browser, like Word's Sources.xml in the user
 * profile), plus the page-number source for generated tables.
 */

import type { BibliographySource, PageNumberProvider } from "@office-kit/docx";
import type { EditorSession } from "../session.svelte";

const MASTER_LIST_KEY = "word-kit:bibliography-master-list";

function loadMasterList(): BibliographySource[] {
  try {
    const raw = localStorage.getItem(MASTER_LIST_KEY);
    const parsed: unknown = raw ? JSON.parse(raw) : [];
    // Stored by this page only; entries without a tag or type cannot be cited.
    return Array.isArray(parsed)
      ? parsed.filter(
          (s): s is BibliographySource =>
            typeof s === "object" && s !== null && typeof s.tag === "string" && typeof s.type === "string",
        )
      : [];
  } catch {
    // Private windows and blocked storage: the master list starts empty.
    return [];
  }
}

class MasterList {
  sources = $state<BibliographySource[]>(loadMasterList());

  /** Add or replace (by tag) and persist. */
  save(source: BibliographySource): void {
    this.sources = [...this.sources.filter((s) => s.tag !== source.tag), source];
    this.persist();
  }

  remove(tag: string): void {
    this.sources = this.sources.filter((s) => s.tag !== tag);
    this.persist();
  }

  private persist(): void {
    try {
      localStorage.setItem(MASTER_LIST_KEY, JSON.stringify(this.sources));
    } catch {
      // Storage full or blocked: the list still works for this visit.
    }
  }
}

export const masterList = new MasterList();

/** The source being edited in the Create Source dialog, and what OK does with it. */
class SourceEditor {
  editing = $state<BibliographySource | null>(null);
  /** Insert a citation to the source after saving it (Insert Citation ▸ Add New Source). */
  citeAfterSave = $state(false);
}

export const sourceEditor = new SourceEditor();

/**
 * The canvas's laid-out page of a paragraph when it reports one, so Update
 * Table writes the page numbers the user sees; generated tables fall back to
 * page estimates from the document's breaks otherwise.
 */
export function pageProvider(session: EditorSession): PageNumberProvider | undefined {
  const withPages: EditorSession & { pageOfParagraph?: PageNumberProvider } = session;
  return withPages.pageOfParagraph;
}
