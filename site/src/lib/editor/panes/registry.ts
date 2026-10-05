/**
 * Task panes the ribbon can dock beside the page, by id. A ribbon button opens
 * one with `session.togglePane(side, id)`.
 */

import type { Component } from "svelte";
import NavigationPane from "./NavigationPane.svelte";
import StylesPane from "./StylesPane.svelte";
import XmlPane from "./XmlPane.svelte";

export const PANES: Readonly<Record<string, Component>> = {
  navigation: NavigationPane,
  styles: StylesPane,
  xml: XmlPane,
};
