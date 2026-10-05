<script lang="ts">
  import type {
    Command,
    EditorModel,
    DocPosition,
    Selection as EditorSelection,
  } from '@office-kit/docx-editor';
  import {
    renderDocumentHtml,
    positionFromDom,
    readDomSelection,
    runAtPath,
    setSimpleRunText,
    runCommand,
    commands,
    orderSelection,
    pageGeometry,
    isTrackingRevisions,
  } from '@office-kit/docx-editor';
  import { getSession } from './session.svelte';

  interface Props {
    model: EditorModel;
    /** Bumped by the parent whenever a command mutates the doc structurally. */
    version: number;
    /** Zoom factor (1 = 100%). */
    zoom?: number;
    /** Called after selection changes so the ribbon can refresh active states. */
    onselectionchange?: () => void;
    /** Called after the canvas itself mutates the doc (typing / Enter / paste). */
    onedit?: () => void;
    /** Called when a canvas gesture's command was rejected (document unchanged). */
    onerror?: (message: string) => void;
  }

  let { model, version, zoom = 1, onselectionchange, onedit, onerror }: Props = $props();

  let canvas = $state<HTMLDivElement | null>(null);
  const session = getSession();

  const TWIPS_PER_POINT = 20;
  // Page size and margins as CSS, re-read on every structural change (page
  // setup commands bump `version`).
  const pageCss = $derived.by(() => {
    if (version < 0) return '';
    const page = pageGeometry(model.doc);
    const pt = (twips: number): string => `${twips / TWIPS_PER_POINT}pt`;
    return [
      `--page-w:${pt(page.width)}`,
      `--page-h:${pt(page.height)}`,
      `--m-top:${pt(page.top)}`,
      `--m-right:${pt(page.right)}`,
      `--m-bottom:${pt(page.bottom)}`,
      `--m-left:${pt(page.left)}`,
    ].join(';');
  });
  // Bumped by structural edits made *inside* the canvas (Enter / Backspace /
  // undo / paste) to force a re-render + caret restore, without the parent
  // having to bump `version`. Plain typing does NOT bump this — the browser
  // owns the caret during input and re-rendering mid-keystroke would reset it.
  let renderTick = $state(0);

  // Whether consecutive keystrokes are being coalesced into one undo step.
  // The first `input` of a burst snapshots the document (the AST still holds
  // the pre-keystroke text at that point); the burst ends on any structural
  // edit, caret jump, blur, or parent re-render, so the next burst gets its
  // own undo step instead of being folded into an unrelated command.
  let typingOpen = false;
  // Set when a typed character replaced a range: the re-render that follows
  // must not end the burst, so the rest of the word joins the same undo step
  // (Word undoes "typing over a selection" as one step).
  let keepBurstAcrossRender = false;

  function endTyping(): void {
    typingOpen = false;
  }

  // An IME composition that starts over a *range*. compositionstart cannot be
  // cancelled, so the browser replaces the range in the DOM natively, across
  // runs and paragraphs, in a shape the per-run sync cannot map back. While
  // this is set, DOM→model sync and selection tracking are frozen; when the
  // composition ends, the composed text replaces the remembered range through
  // one insertTextCommand (one atomic undo step) and the re-render discards
  // whatever the browser did to the DOM.
  let imeRange: EditorSelection | null = null;
  let imeFinalize: { timer: ReturnType<typeof setTimeout>; run: () => void } | null = null;

  function flushImeFinalize(): void {
    if (!imeFinalize) return;
    clearTimeout(imeFinalize.timer);
    const { run } = imeFinalize;
    imeFinalize = null;
    run();
  }

  function onCompositionStart(): void {
    // A new composition right after the previous one: apply that one first.
    flushImeFinalize();
    // While tracking, a composition at a caret must also become one tracked insert.
    if (!hasRangeSelection() && !isTrackingRevisions(model)) return;
    endTyping();
    imeRange = model.selection;
  }

  function onCompositionEnd(e: CompositionEvent): void {
    const range = imeRange;
    if (!range) return;
    const text = e.data ?? '';
    const run = (): void => {
      imeRange = null;
      model.setSelection(range);
      // On failure the model is unchanged but the DOM is not: re-render anyway.
      if (!exec(commands.insertTextCommand, { text })) rerender();
    };
    // Deferred one task: Safari dispatches the composition's final `input`
    // (and `beforeinput`) *after* compositionend, and they must still see the
    // freeze rather than being synced or replayed as a second insert.
    imeFinalize = { timer: setTimeout(flushImeFinalize), run };
  }

  /** Re-render after a structural edit the canvas performed itself. */
  function rerender(): void {
    endTyping();
    renderTick++;
  }

  // Re-render the canvas HTML only when `version` changes (a structural/format
  // edit), never on every keystroke — that would reset the caret. Character
  // typing is reconciled back into the AST by syncFromDom() on input. After a
  // structural re-render we restore the caret to the model's selection so edits
  // like Enter / Backspace / undo don't dump the caret at the top.
  $effect(() => {
    if (!canvas) return;
    // Depend on both the parent's structural `version` and the canvas-local
    // `renderTick` so either source of structural change re-renders.
    canvas.dataset.version = `${version}.${renderTick}`;
    if (keepBurstAcrossRender) keepBurstAcrossRender = false;
    else endTyping();
    canvas.innerHTML = renderDocumentHtml(model.doc);
    restoreCaret();
  });

  /** Attribute selector for the paragraph a position lives in. */
  function paragraphSelector(pos: DocPosition): string {
    return pos.cell
      ? `[data-wk-block="${pos.block}"][data-wk-cell="${pos.cell.row},${pos.cell.col}"][data-wk-para="${pos.para ?? 0}"]`
      : `[data-wk-block="${pos.block}"]:not([data-wk-cell])`;
  }

  /** The run span (or paragraph fallback) that hosts a document position. */
  function hostFor(pos: DocPosition): HTMLElement | null {
    if (!canvas) return null;
    const span = canvas.querySelector<HTMLElement>(
      `.wk-run${paragraphSelector(pos)}[data-wk-inline="${pos.inline ?? 0}"]`,
    );
    if (span) return span;
    // Empty paragraph: no run span, caret lives in the <p> itself.
    return canvas.querySelector<HTMLElement>(`.wk-p${paragraphSelector(pos)}`);
  }

  /** Map a document position to a concrete DOM (textNode, offset) caret point. */
  function domPoint(pos: DocPosition): { node: Node; offset: number } | null {
    const host = hostFor(pos);
    if (!host) return null;
    const text = host.firstChild;
    if (text && text.nodeType === Node.TEXT_NODE) {
      const raw = text.textContent ?? '';
      // The renderer uses a zero-width space as an empty placeholder; clamp the
      // caret to real content length so it lands correctly.
      const len = raw.replace(/​/g, '').length || raw.length;
      return { node: text, offset: Math.min(pos.offset ?? 0, len) };
    }
    return { node: host, offset: 0 };
  }

  function restoreCaret(): void {
    const sel = model.selection;
    if (!sel) return;
    const focus = domPoint(sel.focus);
    if (!focus) return;
    const anchor = domPoint(sel.anchor) ?? focus;
    const range = document.createRange();
    range.setStart(anchor.node, anchor.offset);
    range.setEnd(focus.node, focus.offset);
    const domSel = window.getSelection();
    if (!domSel) return;
    domSel.removeAllRanges();
    domSel.addRange(range);
  }

  /** Push edited run text from the DOM back into the AST (no re-render). */
  function syncFromDom(): void {
    if (!canvas) return;
    const spans = canvas.querySelectorAll<HTMLElement>('.wk-run[data-wk-block]');
    for (const span of spans) {
      // Same anchor parsing as DOM-selection mapping, so typing and the caret
      // can never disagree about which run a span is.
      const pos = positionFromDom(span, 0);
      if (!pos) continue;
      const run = runAtPath(model.doc, pos);
      if (!run) continue;
      const text = (span.textContent ?? '').replace(/​/g, '');
      setSimpleRunText(run, text);
    }
  }

  function onInput(): void {
    if (imeRange) return;
    if (!typingOpen) {
      model.beginEdit();
      typingOpen = true;
    }
    syncFromDom();
    model.commit();
    onedit?.();
  }

  /**
   * Run a command from a canvas gesture; failures are reported, not thrown.
   * Returns whether it applied.
   */
  function exec<P>(cmd: Command<P>, params: P): boolean {
    endTyping();
    try {
      runCommand(model, cmd, params);
    } catch (err) {
      // Rolled back atomically, and the gesture's default was prevented, so
      // the DOM still matches the model: no re-render needed.
      onerror?.(`${cmd.label}: ${(err as Error).message}`);
      return false;
    }
    rerender();
    onedit?.();
    return true;
  }

  // Keys that move the caret end a typing burst (Word starts a new undo step
  // after the caret jumps).
  const CARET_KEYS: ReadonlySet<string> = new Set([
    'ArrowLeft',
    'ArrowRight',
    'ArrowUp',
    'ArrowDown',
    'Home',
    'End',
    'PageUp',
    'PageDown',
  ]);

  function onSelChange(): void {
    if (imeRange) return;
    const sel = readDomSelection(document);
    if (sel) model.setSelection(sel);
    onselectionchange?.();
  }

  /** Whether the caret is a collapsed cursor at the very start of its paragraph. */
  function caretAtParagraphStart(): boolean {
    const sel = model.selection;
    if (!sel) return false;
    const domSel = window.getSelection();
    if (!domSel || !domSel.isCollapsed) return false;
    return (sel.focus.inline ?? 0) === 0 && (sel.focus.offset ?? 0) === 0;
  }

  /** Whether the caret is a collapsed cursor at the very end of its paragraph. */
  function caretAtParagraphEnd(): boolean {
    const domSel = window.getSelection();
    if (!domSel || !domSel.isCollapsed || !canvas) return false;
    const pos = model.selection?.focus;
    if (!pos) return false;
    const host = hostFor(pos);
    if (!host) return false;
    // Last run span in the same paragraph?
    const runs = canvas.querySelectorAll<HTMLElement>(`.wk-run${paragraphSelector(pos)}`);
    const isLast = runs.length === 0 || runs[runs.length - 1] === host;
    const len = (host.textContent ?? '').replace(/​/g, '').length;
    return isLast && (pos.offset ?? 0) >= len;
  }

  function onKeydown(e: KeyboardEvent): void {
    // While an IME is composing (Japanese/Chinese input), Enter confirms the
    // conversion and Backspace edits the candidate — they must not reach the
    // paragraph commands. Safari reports the composing keydown only through
    // the legacy keyCode 229.
    if (e.isComposing || e.keyCode === 229) return;
    // A key after a range composition must act on the applied result.
    flushImeFinalize();
    if (CARET_KEYS.has(e.key)) endTyping();
    const mod = e.ctrlKey || e.metaKey;

    // Undo / redo.
    if (mod && e.key.toLowerCase() === 'z') {
      e.preventDefault();
      // Typed text is already in the AST (onInput syncs every keystroke), so
      // undo/redo can act on the model directly.
      endTyping();
      if (e.shiftKey) model.redo();
      else model.undo();
      rerender();
      onedit?.();
      return;
    }
    if (mod && e.key.toLowerCase() === 'y') {
      e.preventDefault();
      endTyping();
      model.redo();
      rerender();
      onedit?.();
      return;
    }

    // Enter splits the paragraph at the caret (Shift+Enter = soft line break).
    if (e.key === 'Enter') {
      e.preventDefault();
      if (e.shiftKey) exec(commands.insertLineBreakCommand, { kind: 'line' });
      else exec(commands.splitParagraphCommand, undefined);
      return;
    }

    // Backspace at paragraph start merges into the previous paragraph.
    if (
      e.key === 'Backspace' &&
      caretAtParagraphStart() &&
      (commands.mergeBackCommand.isEnabled?.(model) ?? false)
    ) {
      e.preventDefault();
      exec(commands.mergeBackCommand, undefined);
      return;
    }

    // Delete at paragraph end pulls the next paragraph up into this one.
    if (e.key === 'Delete' && caretAtParagraphEnd()) {
      const here = model.selection?.focus;
      if (here?.cell && canvas) {
        // Inside a cell only the next paragraph *of the same cell* may be
        // pulled up; at the cell's last paragraph Delete does nothing (as in
        // Word) instead of letting the browser merge DOM nodes on its own.
        e.preventDefault();
        const next = { block: here.block, cell: here.cell, para: (here.para ?? 0) + 1 };
        const nextExists = !!canvas.querySelector(
          `.wk-p[data-wk-block="${next.block}"][data-wk-cell="${next.cell.row},${next.cell.col}"][data-wk-para="${next.para}"]`,
        );
        if (nextExists) {
          model.setSelection({ anchor: next, focus: next });
          if (!exec(commands.mergeBackCommand, undefined)) model.setSelection({ anchor: here, focus: here });
        }
        return;
      }
      if (here && canvas) {
        const nextBlock = here.block + 1;
        // Only a top-level paragraph can be pulled up (not a table / raw block).
        const nextExists = !!canvas.querySelector(`.wk-p[data-wk-block="${nextBlock}"]:not([data-wk-cell])`);
        if (nextExists) {
          e.preventDefault();
          // Merge the next paragraph into this one, keeping the caret at the join.
          model.setSelection({ anchor: { block: nextBlock }, focus: { block: nextBlock } });
          if (!exec(commands.mergeBackCommand, undefined)) model.setSelection({ anchor: here, focus: here });
          return;
        }
      }
    }

    // Formatting shortcuts.
    if (mod && ['b', 'i', 'u'].includes(e.key.toLowerCase())) {
      e.preventDefault();
      const cmd =
        e.key.toLowerCase() === 'b'
          ? commands.toggleBoldCommand
          : e.key.toLowerCase() === 'i'
            ? commands.toggleItalicCommand
            : commands.toggleUnderlineCommand;
      exec(cmd, undefined);
    }
  }

  function hasRangeSelection(): boolean {
    const sel = model.selection;
    return !!sel && !orderSelection(sel).collapsed;
  }

  /**
   * Typing / deleting over a *range* is routed to commands: the browser's own
   * handling would restructure the DOM across runs and paragraphs in ways the
   * per-run text sync cannot map back. Collapsed-caret typing stays native
   * (reconciled by onInput) so the caret and IME behave normally.
   */
  function onBeforeInput(e: InputEvent): void {
    // While tracking changes every edit is recorded by a command, never synced natively.
    if (imeRange || e.isComposing || (!hasRangeSelection() && !isTrackingRevisions(model))) return;
    if (e.inputType === 'insertText') {
      e.preventDefault();
      // The replacement's undo snapshot (taken by runCommand) is the burst's
      // starting point; following keystrokes sync into it without a new one.
      if (exec(commands.insertTextCommand, { text: e.data ?? '' })) {
        typingOpen = true;
        keepBurstAcrossRender = true;
      }
    } else if (e.inputType.startsWith('delete')) {
      e.preventDefault();
      if (hasRangeSelection()) exec(commands.deleteSelectionCommand, undefined);
      else exec(commands.trackedDeleteCommand, { direction: e.inputType.endsWith('Forward') ? 1 : -1 });
    }
  }

  function onCut(e: ClipboardEvent): void {
    if (!hasRangeSelection() || !e.clipboardData) return;
    e.preventDefault();
    e.clipboardData.setData('text/plain', window.getSelection()?.toString() ?? '');
    exec(commands.deleteSelectionCommand, undefined);
  }

  function onPaste(e: ClipboardEvent): void {
    flushImeFinalize();
    const text = e.clipboardData?.getData('text/plain');
    if (text === undefined) return;
    e.preventDefault();
    exec(commands.insertTextCommand, { text });
  }

  $effect(() => {
    document.addEventListener('selectionchange', onSelChange);
    return () => document.removeEventListener('selectionchange', onSelChange);
  });
</script>

<div class="wk-page" style="zoom: {zoom}; {pageCss}">
  <div
    bind:this={canvas}
    class="wk-canvas"
    contenteditable={!session.readOnly}
    spellcheck={session.spellcheck}
    role="textbox"
    tabindex="0"
    aria-multiline="true"
    aria-label="Document editor"
    oninput={onInput}
    onkeydown={onKeydown}
    onmousedown={endTyping}
    onblur={endTyping}
    onpaste={onPaste}
    onbeforeinput={onBeforeInput}
    oncut={onCut}
    oncompositionstart={onCompositionStart}
    oncompositionend={onCompositionEnd}
  ></div>
</div>

<style>
  .wk-page {
    display: flex;
    justify-content: center;
  }
  /*
   * The page as Word draws it: the section's paper size and margins, a thin
   * gray edge with a soft shadow, and the L-shaped crop marks Word puts at the
   * corners of the text area.
   */
  .wk-canvas {
    --mark: 13.5pt;
    --mark-color: #a6a6a6;
    box-sizing: border-box;
    flex: none;
    width: var(--page-w);
    min-height: var(--page-h);
    padding: var(--m-top) var(--m-right) var(--m-bottom) var(--m-left);
    background-color: #fff;
    background-repeat: no-repeat;
    background-image:
      linear-gradient(var(--mark-color), var(--mark-color)),
      linear-gradient(var(--mark-color), var(--mark-color)),
      linear-gradient(var(--mark-color), var(--mark-color)),
      linear-gradient(var(--mark-color), var(--mark-color)),
      linear-gradient(var(--mark-color), var(--mark-color)),
      linear-gradient(var(--mark-color), var(--mark-color)),
      linear-gradient(var(--mark-color), var(--mark-color)),
      linear-gradient(var(--mark-color), var(--mark-color));
    background-size:
      var(--mark) 1px, 1px var(--mark),
      var(--mark) 1px, 1px var(--mark),
      var(--mark) 1px, 1px var(--mark),
      var(--mark) 1px, 1px var(--mark);
    background-position:
      calc(var(--m-left) - var(--mark)) var(--m-top),
      calc(var(--m-left) - 1px) calc(var(--m-top) - var(--mark)),
      calc(100% - var(--m-right) + var(--mark)) var(--m-top),
      calc(100% - var(--m-right) + 1px) calc(var(--m-top) - var(--mark)),
      calc(var(--m-left) - var(--mark)) calc(100% - var(--m-bottom)),
      calc(var(--m-left) - 1px) calc(100% - var(--m-bottom) + var(--mark)),
      calc(100% - var(--m-right) + var(--mark)) calc(100% - var(--m-bottom)),
      calc(100% - var(--m-right) + 1px) calc(100% - var(--m-bottom) + var(--mark));
    border: 1px solid #c6c6c6;
    box-shadow: 0 1px 6px rgba(0, 0, 0, 0.12);
    outline: none;
    color: #000;
    /* Only for the caret in an empty document; every paragraph and run carries
       its style-resolved font and size. */
    font-family: Calibri, Carlito, 'Segoe UI', system-ui, sans-serif;
    font-size: 11pt;
    line-height: 1.2;
  }

  .wk-canvas :global(.wk-p) {
    margin: 0;
    min-height: 1.2em;
    /* Runs render tabs and <w:br/> as \t / \n; keep them visible. */
    white-space: pre-wrap;
  }

  .wk-canvas :global(.wk-link) {
    color: #1a56c4;
    text-decoration: underline;
  }

  .wk-canvas :global(.wk-inline-raw) {
    background: #f3f4f6;
  }

  /* Width, columns, borders and cell margins come from the document. */
  .wk-canvas :global(.wk-table) {
    border-collapse: collapse;
    table-layout: fixed;
    margin: 0;
  }

  .wk-canvas :global(.wk-td) {
    vertical-align: top;
  }

  .wk-canvas :global(.wk-raw) {
    color: #888;
    font-style: italic;
    user-select: none;
    background: #f3f4f6;
    padding: 2px 6px;
    border-radius: 3px;
    display: inline-block;
  }
</style>
