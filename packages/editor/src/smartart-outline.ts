/**
 * The SmartArt Text Pane edits the diagram's data model as a bulleted outline.
 * These helpers convert between the node tree and outline rows, and implement
 * SmartArt Design ▸ Add Shape, Promote, Demote, Move Up and Move Down on rows.
 */

import type { SmartArtNode } from "@office-kit/docx";

export interface OutlineRow {
  readonly level: number;
  readonly text: string;
}

export function outlineRows(
  nodes: readonly SmartArtNode[],
  level = 0,
  out: OutlineRow[] = [],
): OutlineRow[] {
  for (const node of nodes) {
    out.push({ level, text: node.text });
    if (node.children) outlineRows(node.children, level + 1, out);
  }
  return out;
}

/**
 * Rebuild the tree. A row indented more than one level below its predecessor
 * is attached one level down, as Word's Text Pane does.
 */
export function outlineToNodes(rows: readonly OutlineRow[]): SmartArtNode[] {
  interface Draft {
    text: string;
    children: Draft[];
  }
  const root: Draft[] = [];
  const stack: Draft[][] = [root];
  for (const row of rows) {
    const level = Math.min(row.level, stack.length - 1);
    stack.length = level + 1;
    const draft: Draft = { text: row.text, children: [] };
    stack[level]?.push(draft);
    stack.push(draft.children);
  }
  const finish = (list: readonly Draft[]): SmartArtNode[] =>
    list.map((d) =>
      d.children.length ? { text: d.text, children: finish(d.children) } : { text: d.text },
    );
  return finish(root);
}

/** The row and its descendants (the rows below it at a deeper level). */
function subtreeEnd(rows: readonly OutlineRow[], index: number): number {
  const level = rows[index]?.level ?? 0;
  let end = index + 1;
  while (end < rows.length && (rows[end]?.level ?? 0) > level) end++;
  return end;
}

export type OutlineOp =
  | "addAfter"
  | "addBefore"
  | "addAbove"
  | "addBelow"
  | "promote"
  | "demote"
  | "moveUp"
  | "moveDown";

/** Apply an outline operation; returns the new rows and the row to select. */
export function applyOutlineOp(
  rows: readonly OutlineRow[],
  index: number,
  op: OutlineOp,
  placeholder: string,
): { rows: OutlineRow[]; index: number } {
  const next = [...rows];
  const row = next[index];
  if (!row) return { rows: next, index };
  const end = subtreeEnd(next, index);
  switch (op) {
    case "addAfter":
      next.splice(end, 0, { level: row.level, text: placeholder });
      return { rows: next, index: end };
    case "addBefore":
      next.splice(index, 0, { level: row.level, text: placeholder });
      return { rows: next, index };
    case "addAbove": {
      // A new parent: the shape and its subtree move one level down under it.
      next.splice(index, 0, { level: row.level, text: placeholder });
      for (let i = index + 1; i <= end; i++) {
        const r = next[i];
        if (r) next[i] = { ...r, level: r.level + 1 };
      }
      return { rows: next, index };
    }
    case "addBelow":
      next.splice(index + 1, 0, { level: row.level + 1, text: placeholder });
      return { rows: next, index: index + 1 };
    case "promote":
    case "demote": {
      const delta = op === "promote" ? -1 : 1;
      const prevLevel = next[index - 1]?.level ?? -1;
      if (op === "promote" ? row.level === 0 : row.level > prevLevel) return { rows: next, index };
      for (let i = index; i < end; i++) {
        const r = next[i];
        if (r) next[i] = { ...r, level: r.level + delta };
      }
      return { rows: next, index };
    }
    case "moveUp":
    case "moveDown": {
      const block = next.slice(index, end);
      if (op === "moveUp") {
        let start = index - 1;
        while (start > 0 && (next[start]?.level ?? 0) > row.level) start--;
        if (start < 0 || (next[start]?.level ?? 0) !== row.level) return { rows: next, index };
        next.splice(index, block.length);
        next.splice(start, 0, ...block);
        return { rows: next, index: start };
      }
      const sibling = next[end];
      if (!sibling || sibling.level !== row.level) return { rows: next, index };
      const siblingEnd = subtreeEnd(next, end);
      next.splice(index, block.length);
      const at = siblingEnd - block.length;
      next.splice(at, 0, ...block);
      return { rows: next, index: at };
    }
  }
}
