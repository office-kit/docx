/**
 * What the Insert-tab dialogs are opened *for*: the equation being edited,
 * whether the Link dialog edits an existing link, and the user name comments
 * and USERNAME fields use. The dialogs themselves are always mounted (see
 * InsertDialogs.svelte), so any tab — or a double-click on the canvas — can
 * open them through `session.openDialog(id)`.
 */

import type { FieldContext } from "@office-kit/docx";
import type { EquationRef } from "@office-kit/docx-editor";
import type { EditorSession } from "../session.svelte";

/** Dialog ids (stable: other tabs open these). */
export const DIALOG = {
  link: "insert.link",
  bookmark: "insert.bookmark",
  crossReference: "insert.crossReference",
  comment: "insert.comment",
  field: "insert.field",
  dateTime: "insert.dateTime",
  symbol: "insert.symbol",
  equation: "insert.equation",
  dropCap: "insert.dropCap",
  pageNumberFormat: "insert.pageNumberFormat",
  signatureLine: "insert.signatureLine",
} as const;

const USER_NAME_KEY = "wk-editor-user-name";

function storedUserName(): string {
  try {
    return localStorage.getItem(USER_NAME_KEY) ?? "";
  } catch {
    // Storage blocked (private mode, sandboxed preview): fall back to empty.
    return "";
  }
}

class InsertUi {
  /** Set before opening the equation dialog to edit instead of insert. */
  equation = $state<EquationRef | null>(null);
  /** Word's user name (Preferences ▸ User Information), remembered per browser. */
  userName = $state(typeof localStorage === "undefined" ? "" : storedUserName());

  rememberUserName(name: string): void {
    this.userName = name;
    try {
      localStorage.setItem(USER_NAME_KEY, name);
    } catch {
      // Not persisted when storage is blocked; the name still applies this session.
    }
  }

  /** Initials Word derives from the user name ("Ada Lovelace" → "AL"). */
  initials(): string {
    return this.userName
      .split(/\s+/)
      .filter(Boolean)
      .map((w) => w[0]?.toUpperCase() ?? "")
      .join("");
  }
}

export const insertUi = new InsertUi();

/** What field results computed in the browser know: layout, file and user. */
export function fieldContext(session: EditorSession): FieldContext {
  return {
    now: new Date(),
    page: session.currentPage,
    pageCount: session.pageCount,
    fileName: session.fileName,
    userName: insertUi.userName,
    userInitials: insertUi.initials(),
  };
}
