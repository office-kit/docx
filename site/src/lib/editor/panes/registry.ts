/**
 * Task panes the ribbon can dock beside the page, by id. A ribbon button opens
 * one with `session.togglePane(side, id)`.
 */

import type { Component } from "svelte";
import AltTextPane from "./AltTextPane.svelte";
import FormatPicturePane from "./FormatPicturePane.svelte";
import NavigationPane from "./NavigationPane.svelte";
import SelectionPane from "./SelectionPane.svelte";
import XmlPane from "./XmlPane.svelte";

export const PANES: Readonly<Record<string, Component>> = {
  navigation: NavigationPane,
  xml: XmlPane,
  altText: AltTextPane,
  formatPicture: FormatPicturePane,
  selection: SelectionPane,
};
