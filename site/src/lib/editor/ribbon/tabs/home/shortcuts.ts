/**
 * Word for Mac's Home keyboard shortcuts. ⌘B / ⌘I / ⌘U, undo and typing are
 * the canvas' own; ⌘F / ⌘H open the Navigation pane in the page.
 */

import { BULLET_PRESETS, commands, type RunToggle } from "@office-kit/docx-editor";
import type { EditorSession } from "../../../session.svelte";
import { homeState, selectionFormats } from "./state.svelte";

type Action = (session: EditorSession) => void;

const HEADINGS = ["Heading1", "Heading2", "Heading3"] as const;

function toggleEffect(session: EditorSession, effect: RunToggle): void {
  const model = session.model;
  if (!model) return;
  const runs = selectionFormats(model).runs;
  const on = runs.length > 0 && runs.every((r) => r.toggles.has(effect));
  session.apply(commands.fontFormatCommand, { effects: { [effect]: !on } });
}

function toggleVertAlign(session: EditorSession, val: "superscript" | "subscript"): void {
  const model = session.model;
  if (!model) return;
  const runs = selectionFormats(model).runs;
  const on = runs.length > 0 && runs.every((r) => r.vertAlign === val);
  session.apply(commands.setVertAlignCommand, { val: on ? undefined : val });
}

/** Shift+F3 cycles UPPERCASE → lowercase → Capitalize Each Word, from the selection's current case. */
function cycleCase(session: EditorSession): void {
  const text = document.getSelection()?.toString() ?? "";
  const mode =
    text === text.toUpperCase() && text !== text.toLowerCase()
      ? "lower"
      : text === text.toLowerCase()
        ? "title"
        : "upper";
  session.apply(commands.changeCaseCommand, { mode });
}

function toggleDoubleUnderline(session: EditorSession): void {
  const model = session.model;
  if (!model) return;
  const on = selectionFormats(model).runs.every((r) => r.underline === "double");
  if (on) session.apply(commands.fontFormatCommand, { underline: null });
  else session.apply(commands.setUnderlineStyleCommand, { style: "double" });
}

function copyFormat(session: EditorSession): void {
  const format = session.model && commands.copyFormat(session.model);
  if (format) homeState(session).copiedFormat = format;
}

function pasteFormat(session: EditorSession): void {
  const format = homeState(session).copiedFormat;
  if (format) session.apply(commands.pasteFormatCommand, format);
}

/** Shortcut → action, keyed by modifiers and `KeyboardEvent.code` (layout-independent, as Word's are). */
const SHORTCUTS: Readonly<Record<string, Action>> = {
  "mod+KeyL": (s) => s.apply(commands.alignLeftCommand, undefined),
  "mod+KeyE": (s) => s.apply(commands.alignCenterCommand, undefined),
  "mod+KeyR": (s) => s.apply(commands.alignRightCommand, undefined),
  "mod+KeyJ": (s) => s.apply(commands.alignJustifyCommand, undefined),
  "mod+Digit8": (s) => (s.showMarks = !s.showMarks),
  "mod+KeyD": (s) => s.openDialog("home.font"),
  "mod+alt+KeyM": (s) => s.openDialog("home.paragraph"),
  "mod+shift+Period": (s) => s.apply(commands.growFontCommand, { direction: 1 }),
  "mod+shift+Comma": (s) => s.apply(commands.growFontCommand, { direction: -1 }),
  "mod+BracketRight": (s) => s.apply(commands.nudgeFontCommand, { direction: 1 }),
  "mod+BracketLeft": (s) => s.apply(commands.nudgeFontCommand, { direction: -1 }),
  "mod+Equal": (s) => toggleVertAlign(s, "subscript"),
  "mod+shift+Equal": (s) => toggleVertAlign(s, "superscript"),
  "mod+shift+KeyA": (s) => toggleEffect(s, "caps"),
  "mod+shift+KeyK": (s) => toggleEffect(s, "smallCaps"),
  "mod+shift+KeyH": (s) => toggleEffect(s, "vanish"),
  "mod+shift+KeyD": toggleDoubleUnderline,
  "ctrl+Space": (s) => s.apply(commands.clearCharacterFormattingCommand, undefined),
  "mod+alt+Digit1": (s) => s.apply(commands.applyStyleCommand, { styleId: HEADINGS[0] }),
  "mod+alt+Digit2": (s) => s.apply(commands.applyStyleCommand, { styleId: HEADINGS[1] }),
  "mod+alt+Digit3": (s) => s.apply(commands.applyStyleCommand, { styleId: HEADINGS[2] }),
  "mod+shift+KeyN": (s) => {
    const normal = s.model && commands.defaultParagraphStyleId(s.model);
    if (normal) s.apply(commands.applyStyleCommand, { styleId: normal });
  },
  "mod+shift+KeyL": (s) => {
    const preset = BULLET_PRESETS[0];
    if (preset) s.apply(commands.applyListPresetCommand, { levels: preset.levels });
  },
  "mod+shift+KeyC": copyFormat,
  "mod+shift+KeyV": pasteFormat,
  "mod+Digit1": (s) => s.apply(commands.setLineSpacingCommand, { multiple: 1 }),
  "mod+Digit2": (s) => s.apply(commands.setLineSpacingCommand, { multiple: 2 }),
  "mod+Digit5": (s) => s.apply(commands.setLineSpacingCommand, { multiple: 1.5 }),
  "mod+Digit0": (s) => s.apply(commands.toggleParagraphSpaceCommand, { side: "before" }),
  "shift+F3": cycleCase,
  "mod+alt+shift+KeyS": (s) => s.togglePane("right", "styles"),
};

function chord(e: KeyboardEvent): string {
  // Word for Mac uses ⌘ where Windows uses Ctrl; either works here. Ctrl+Space
  // is the same on both.
  const mod = e.metaKey || (e.ctrlKey && e.code !== "Space");
  return [
    mod && "mod",
    !mod && e.ctrlKey && "ctrl",
    e.altKey && "alt",
    e.shiftKey && "shift",
    e.code,
  ]
    .filter(Boolean)
    .join("+");
}

/** Run the Home shortcut for a key press; returns whether there was one. */
export function handleHomeShortcut(e: KeyboardEvent, session: EditorSession): boolean {
  const action = SHORTCUTS[chord(e)];
  // Shortcuts act on the document: not while a dialog is open or a form
  // field (search box, font box …) has the keyboard.
  const inField = e.target instanceof Element && !!e.target.closest("input, textarea, select");
  if (!action || !session.model || session.dialog || inField) return false;
  action(session);
  return true;
}
