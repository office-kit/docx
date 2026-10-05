/**
 * Task panes the ribbon can dock beside the page, by id. A ribbon button opens
 * one with `session.togglePane(side, id)`.
 */

import type { Component } from "svelte";
import CitationsPane from "./CitationsPane.svelte";
import NavigationPane from "./NavigationPane.svelte";
import NotesPane from "./NotesPane.svelte";
import XmlPane from "./XmlPane.svelte";

export const PANES: Readonly<Record<string, Component>> = {
  navigation: NavigationPane,
  xml: XmlPane,
  notes: NotesPane,
  citations: CitationsPane,
};
