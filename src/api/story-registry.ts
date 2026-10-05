/**
 * Bookkeeping shared by `story.ts` and the save path in `docx.ts`, kept in a
 * module of its own so `docx.ts` does not import `story.ts` (and the story
 * code stays out of bundles that never touch stories).
 *
 * Not public API: nothing here is re-exported from `index.ts`.
 */

import type { Docx } from "./docx.js";

interface StoryEntry {
  /** The package part the story lives in. */
  readonly partName: string;
  /** Write the story's typed content back into the part's XML. */
  readonly flush: () => void;
}

// A story's typed content is authoritative once parsed, so it must be written
// back into its part before every save / clone. Keyed weakly so a discarded
// document (an undo snapshot, say) takes its stories with it.
const entries = new WeakMap<Docx, Map<string, StoryEntry>>();
// Story views are proxies over a real document; saving a view must save the
// document it wraps, never the story it shows as its body.
const viewOwners = new WeakMap<Docx, Docx>();

export function registerStory(doc: Docx, key: string, partName: string, flush: () => void): void {
  let map = entries.get(doc);
  if (!map) {
    map = new Map();
    entries.set(doc, map);
  }
  map.set(key, { partName, flush });
}

/** Whether a parsed story is still live (not discarded by a raw edit). */
export function hasStory(doc: Docx, key: string): boolean {
  return entries.get(doc)?.has(key) ?? false;
}

/** Write every parsed story of `doc` back into its part's XML. */
export function runStoryFlushes(doc: Docx): void {
  const map = entries.get(doc);
  if (!map) return;
  for (const entry of map.values()) entry.flush();
}

/**
 * Forget the parsed stories of a part whose raw XML was edited directly: the
 * raw edit wins, and the story is parsed again from it on next use.
 */
export function discardStoriesOfPart(doc: Docx, partName: string): void {
  const map = entries.get(doc);
  if (!map) return;
  for (const [key, entry] of map) if (entry.partName === partName) map.delete(key);
}

export function registerStoryView(view: Docx, owner: Docx): void {
  viewOwners.set(view, owner);
}

/** The real document behind a story view, or `doc` itself. */
export function storyViewOwner(doc: Docx): Docx {
  return viewOwners.get(doc) ?? doc;
}
