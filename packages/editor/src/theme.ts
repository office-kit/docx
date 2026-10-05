/**
 * The document's theme as the editor uses it: colours for theme-colour
 * references (`w:themeColor`) and the heading / body fonts. A document
 * without a theme part behaves as Word's: theme references fall back to the
 * Office theme.
 */

import { type Docx, getTheme, THEMES, type ThemeInfo } from "@office-kit/docx";

function officeTheme(): ThemeInfo {
  const office = THEMES[0];
  if (!office) throw new Error("The built-in theme list is empty.");
  return {
    name: office.name,
    colors: office.colors,
    fonts: office.fonts,
    effects: office.effects.name,
  };
}

/** The document's theme, or Word's Office theme when it has none. */
export function resolveTheme(doc: Docx): ThemeInfo {
  return getTheme(doc) ?? officeTheme();
}
