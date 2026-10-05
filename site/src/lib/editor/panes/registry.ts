/**
 * Task panes the ribbon can dock beside the page, by id. A ribbon button opens
 * one with `session.togglePane(side, id)`.
 */

import type { Component } from "svelte";
import FormatShapePane from "./FormatShapePane.svelte";
import NavigationPane from "./NavigationPane.svelte";
import SmartArtTextPane from "./SmartArtTextPane.svelte";
import XmlPane from "./XmlPane.svelte";

export const PANES: Readonly<Record<string, Component>> = {
  navigation: NavigationPane,
  xml: XmlPane,
  formatShape: FormatShapePane,
  smartArtText: SmartArtTextPane,
};
