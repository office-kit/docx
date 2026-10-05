/**
 * Task panes the ribbon can dock beside the page, by id. A ribbon button opens
 * one with `session.togglePane(side, id)`.
 */

import type { Component } from "svelte";
import AccessibilityPane from "./AccessibilityPane.svelte";
import CommentsPane from "./CommentsPane.svelte";
import NavigationPane from "./NavigationPane.svelte";
import StylesPane from "./StylesPane.svelte";
import RestrictEditingPane from "./RestrictEditingPane.svelte";
import ReviewingPane from "./ReviewingPane.svelte";
import CitationsPane from "./CitationsPane.svelte";
import NotesPane from "./NotesPane.svelte";
import XmlPane from "./XmlPane.svelte";

export const PANES: Readonly<Record<string, Component>> = {
  navigation: NavigationPane,
  styles: StylesPane,
  xml: XmlPane,
  comments: CommentsPane,
  reviewing: ReviewingPane,
  accessibility: AccessibilityPane,
  restrictEditing: RestrictEditingPane,
  notes: NotesPane,
  citations: CitationsPane,
};
