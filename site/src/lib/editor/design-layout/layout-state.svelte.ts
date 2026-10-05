/**
 * State the Layout tab shares with its dialogs for the editing session:
 * Word's "Last Custom Setting" margins and which Page Setup tab to open on.
 */

import type { PageMargins } from "@office-kit/docx";

export type PageSetupTab = "margins" | "paper" | "layout" | "grid";

class LayoutState {
  /** The margins last applied from the Page Setup dialog (Margins ▸ Last Custom Setting). */
  lastCustomMargins = $state<PageMargins | null>(null);
  pageSetupTab = $state<PageSetupTab>("margins");
}

export const layoutState = new LayoutState();
