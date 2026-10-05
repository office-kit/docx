/**
 * The ribbon's tabs, in Word's order. Contextual tabs (Table Design, Picture
 * Format …) appear after View only while `when` holds for the selection, as in
 * Word, where they are drawn in the accent color.
 */

import type { Component } from "svelte";
import type { MessageKey } from "../i18n/index.svelte";
import type { EditorSession } from "../session.svelte";
import DesignTab from "./tabs/DesignTab.svelte";
import DrawTab from "./tabs/DrawTab.svelte";
import HomeTab from "./tabs/HomeTab.svelte";
import InsertTab from "./tabs/InsertTab.svelte";
import LayoutTab from "./tabs/LayoutTab.svelte";
import MailingsTab from "./tabs/MailingsTab.svelte";
import PictureFormatTab from "./tabs/PictureFormatTab.svelte";
import ReferencesTab from "./tabs/ReferencesTab.svelte";
import ReviewTab from "./tabs/ReviewTab.svelte";
import ShapeFormatTab from "./tabs/ShapeFormatTab.svelte";
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

const never = (): boolean => false;

export const TABS: readonly RibbonTab[] = [
  { id: "home", label: "tab.home", component: HomeTab },
  { id: "insert", label: "tab.insert", component: InsertTab },
  { id: "draw", label: "tab.draw", component: DrawTab },
  { id: "design", label: "tab.design", component: DesignTab },
  { id: "layout", label: "tab.layout", component: LayoutTab },
  { id: "references", label: "tab.references", component: ReferencesTab },
  { id: "mailings", label: "tab.mailings", component: MailingsTab },
  { id: "review", label: "tab.review", component: ReviewTab },
  { id: "view", label: "tab.view", component: ViewTab },
  { id: "shapeFormat", label: "tab.shapeFormat", component: ShapeFormatTab, when: never },
  { id: "pictureFormat", label: "tab.pictureFormat", component: PictureFormatTab, when: never },
  { id: "tableDesign", label: "tab.tableDesign", component: TableDesignTab, when: never },
  { id: "tableLayout", label: "tab.tableLayout", component: TableLayoutTab, when: never },
];
