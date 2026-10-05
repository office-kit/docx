/**
 * The ribbon's tabs, in Word's order. Contextual tabs (Table Design, Picture
 * Format …) appear after View only while `when` holds for the selection, as in
 * Word, where they are drawn in the accent color.
 */

import type { Component } from "svelte";
import { tableSelection } from "@office-kit/docx-editor";
import type { MessageKey } from "../i18n/index.svelte";
import type { EditorSession } from "../session.svelte";
import ChartDesignTab from "./tabs/ChartDesignTab.svelte";
import DesignTab from "./tabs/DesignTab.svelte";
import DrawTab from "./tabs/DrawTab.svelte";
import HeaderFooterTab from "./tabs/HeaderFooterTab.svelte";
import HomeTab from "./tabs/HomeTab.svelte";
import InsertTab from "./tabs/InsertTab.svelte";
import LayoutTab from "./tabs/LayoutTab.svelte";
import MailingsTab from "./tabs/MailingsTab.svelte";
import OutliningTab from "./tabs/OutliningTab.svelte";
import PictureFormatTab from "./tabs/PictureFormatTab.svelte";
import ReferencesTab from "./tabs/ReferencesTab.svelte";
import ReviewTab from "./tabs/ReviewTab.svelte";
import ShapeFormatTab from "./tabs/ShapeFormatTab.svelte";
import SmartArtDesignTab from "./tabs/SmartArtDesignTab.svelte";
import TableDesignTab from "./tabs/TableDesignTab.svelte";
import TableLayoutTab from "./tabs/TableLayoutTab.svelte";
import ViewTab from "./tabs/ViewTab.svelte";

export interface RibbonTab {
  readonly id: string;
  readonly label: MessageKey;
  readonly component: Component;
  /** Contextual tabs only: whether the selection calls for this tab. */
  readonly when?: (session: EditorSession) => boolean;
}

// `tick` is read so the tab appears and disappears as the caret moves.
const inTable = (s: EditorSession): boolean =>
  s.tick >= 0 && !!s.model && !!tableSelection(s.model);
const SHAPE_FORMAT_KINDS: ReadonlySet<string> = new Set(["shape", "textBox", "smartArt", "ink"]);

export const TABS: readonly RibbonTab[] = [
  // Word puts Outlining first, ahead of Home, while the Outline view is on.
  {
    id: "outlining",
    label: "tab.outlining",
    component: OutliningTab,
    when: (session) => session.viewMode === "outline",
  },
  { id: "home", label: "tab.home", component: HomeTab },
  { id: "insert", label: "tab.insert", component: InsertTab },
  { id: "draw", label: "tab.draw", component: DrawTab },
  { id: "design", label: "tab.design", component: DesignTab },
  { id: "layout", label: "tab.layout", component: LayoutTab },
  { id: "references", label: "tab.references", component: ReferencesTab },
  { id: "mailings", label: "tab.mailings", component: MailingsTab },
  { id: "review", label: "tab.review", component: ReviewTab },
  { id: "view", label: "tab.view", component: ViewTab },
  {
    id: "smartArtDesign",
    label: "draw.tab.smartArtDesign",
    component: SmartArtDesignTab,
    when: (s) => s.selectedObject?.kind === "smartArt",
  },
  {
    id: "shapeFormat",
    label: "tab.shapeFormat",
    component: ShapeFormatTab,
    when: (s) => SHAPE_FORMAT_KINDS.has(s.selectedObject?.kind ?? ""),
  },
  {
    id: "pictureFormat",
    label: "tab.pictureFormat",
    component: PictureFormatTab,
    when: (s) => s.selectedObject?.kind === "picture",
  },
  {
    id: "chartDesign",
    label: "tab.chartDesign",
    component: ChartDesignTab,
    when: (s) => s.selectedObject?.kind === "chart",
  },
  { id: "tableDesign", label: "tab.tableDesign", component: TableDesignTab, when: inTable },
  { id: "tableLayout", label: "tab.tableLayout", component: TableLayoutTab, when: inTable },
  {
    id: "headerFooter",
    label: "tab.headerFooter",
    component: HeaderFooterTab,
    when: (s) => s.story?.kind === "header" || s.story?.kind === "footer",
  },
];
