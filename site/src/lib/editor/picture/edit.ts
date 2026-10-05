/** Patch helpers for the selected picture's effects and color adjustments. */

import { commands } from "@office-kit/docx-editor";
import type { PictureColorAdjustments, PictureEffects } from "@office-kit/docx";
import type { EditorSession } from "../session.svelte";
import { applyToSelected, selected } from "./state";

type Mutable<T> = { -readonly [K in keyof T]?: T[K] };

/** `base` with `patch` applied; an `undefined` member removes that property. */
function patched<T extends object>(
  base: T,
  patch: { readonly [K in keyof T]?: T[K] | undefined },
): T {
  const next: Mutable<T> = { ...base };
  for (const key of Object.keys(patch) as Array<keyof T>) {
    const value = patch[key];
    if (value === undefined) delete next[key];
    else next[key] = value;
  }
  // Every member of `next` came from a T or the patch's T-typed values.
  return next as T;
}

export function patchEffects(
  session: EditorSession,
  patch: { readonly [K in keyof PictureEffects]?: PictureEffects[K] | undefined },
): void {
  const effects = selected(session)?.info.picture?.effects;
  if (!effects) return;
  applyToSelected(session, commands.pictureEffectsCommand, { effects: patched(effects, patch) });
}

export function patchAdjustments(
  session: EditorSession,
  patch: { readonly [K in keyof PictureColorAdjustments]?: PictureColorAdjustments[K] | undefined },
): void {
  const adjustments = selected(session)?.info.picture?.adjustments;
  if (!adjustments) return;
  applyToSelected(session, commands.pictureAdjustCommand, {
    adjustments: patched(adjustments, patch),
  });
}

/** Recolor presets replace the previous recolor (and Washout's brightness / contrast). */
export const RECOLOR_KEYS = {
  grayscale: undefined,
  biLevelThreshold: undefined,
  duotone: undefined,
} as const;
