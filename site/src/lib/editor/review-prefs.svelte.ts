/**
 * Review and View preferences that are not part of the document: who you are
 * (Word ▸ Preferences ▸ User Information), how markup is shown (Show Markup,
 * Track Changes Options), Focus, Multiple Pages and the Outline view's display.
 */

/** Track Changes Options ▸ Insertions. */
export type InsertionMark = "underline" | "doubleUnderline" | "bold" | "italic" | "colorOnly";
/** Track Changes Options ▸ Deletions. */
export type DeletionMark = "strikethrough" | "doubleStrikethrough" | "hidden";
/** Track Changes Options ▸ Changed lines. */
export type ChangedLines = "outside" | "left" | "right" | "none";

const USER_KEY = "wk-editor-user";

interface StoredUser {
  readonly name: string;
  readonly initials: string;
}

function initialsOf(name: string): string {
  return name
    .split(/\s+/)
    .filter(Boolean)
    .map((part) => part[0]?.toUpperCase() ?? "")
    .join("")
    .slice(0, 3);
}

/** The saved user name, if the browser keeps one (storage can be blocked). */
function loadUser(): StoredUser | undefined {
  try {
    const raw = localStorage.getItem(USER_KEY);
    if (!raw) return undefined;
    const parsed: unknown = JSON.parse(raw);
    if (
      typeof parsed === "object" &&
      parsed !== null &&
      "name" in parsed &&
      "initials" in parsed &&
      typeof parsed.name === "string" &&
      typeof parsed.initials === "string"
    ) {
      return { name: parsed.name, initials: parsed.initials };
    }
  } catch {
    // No storage (private window, SSR): fall back to the default user.
  }
  return undefined;
}

export class ReviewPrefs {
  userName = $state("Author");
  userInitials = $state("A");
  showComments = $state(true);
  showInsertionsDeletions = $state(true);
  showFormatting = $state(true);
  markupAreaHighlight = $state(true);
  /** The comment selected in the Comments pane (Delete acts on it). */
  activeComment = $state<number | null>(null);
  insertionMark = $state<InsertionMark>("underline");
  deletionMark = $state<DeletionMark>("strikethrough");
  changedLines = $state<ChangedLines>("outside");
  /** View ▸ Multiple Pages: pages side by side. */
  pagesAcross = $state(1);
  /** Outlining ▸ Show Level (1–9; 10 shows all levels and body text). */
  outlineShowLevel = $state(10);
  outlineFirstLineOnly = $state(false);
  outlineShowFormatting = $state(true);
  /** View ▸ Focus: the page alone, chrome hidden. */
  focus = $state(false);
  /** Accessibility checker results update while you work (status bar). */
  keepAccessibilityRunning = $state(true);

  constructor() {
    const user = loadUser();
    if (user) {
      this.userName = user.name;
      this.userInitials = user.initials;
    }
  }

  setUser(name: string, initials: string): void {
    this.userName = name.trim() || "Author";
    this.userInitials = initials.trim() || initialsOf(this.userName);
    try {
      localStorage.setItem(
        USER_KEY,
        JSON.stringify({ name: this.userName, initials: this.userInitials }),
      );
    } catch {
      // Not persisted without storage; the session still uses the new name.
    }
  }
}
