<script lang="ts">
  import { untrack } from 'svelte';
  import type {
    Command,
    EditorModel,
    DocPosition,
    Selection as EditorSelection,
  } from '@office-kit/docx-editor';
  import ShapeOverlay from './shapes/ShapeOverlay.svelte';
  import {
    paragraphAt,
    positionFromDom,
    readDomSelection,
    parseClipboardHtml,
    releasePendingFormat,
    isEmptyParagraph,
    runAtPath,
    setSimpleRunText,
    runCommand,
    commands,
    orderSelection,
    caretAt,
    sameStory,
    isTrackingRevisions,
  } from '@office-kit/docx-editor';
  import { getParagraphNumbering, parseStoryKey, type WmlParagraph } from '@office-kit/docx';
  import { getSession } from './session.svelte';
  import { t } from './i18n/index.svelte';
  import { PageLayout } from './canvas/page-layout';
  import { decorateTableSelection, handleTableKey, tableCanvasStyle, tablePointerDown, tablePointerUp, tableToolAttr } from './table-canvas';
  import { copiesOf, domPoint, edgePoint, focusPage, focusParagraph, mirrorParagraph, paragraphJumpPoint, paragraphSelector, storyRange, type ParagraphJump } from './canvas/caret';
  import ObjectSelection from './picture/ObjectSelection.svelte';

  interface Props {
    model: EditorModel;
  }

  let { model }: Props = $props();
  const session = getSession();

  let canvas = $state<HTMLDivElement | null>(null);
  let galley = $state<HTMLDivElement | null>(null);
  let layout: PageLayout | undefined;

  // Bumped by structural edits made *inside* the canvas (Enter / Backspace /
  // undo / paste) to force a re-render + caret restore, without a session
  // version bump. Plain typing does NOT bump this — the browser owns the
  // caret during input and re-rendering mid-keystroke would reset it.
  let renderTick = $state(0);

  const editingHeaderFooter = $derived(session.story?.kind === 'header' || session.story?.kind === 'footer');
  // The page a header/footer was opened on: its copy there gets the caret.
  let headerFooterPage: number | undefined;

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
  // Any composition in progress: re-pagination waits for it to end.
  let composing = false;

  function flushImeFinalize(): void {
    if (!imeFinalize) return;
    clearTimeout(imeFinalize.timer);
    const { run } = imeFinalize;
    imeFinalize = null;
    run();
  }

  function onCompositionStart(): void {
    composing = true;
    // A new composition right after the previous one: apply that one first.
    flushImeFinalize();
    // While tracking, a composition at a caret must also become one tracked
    // insert; so must one where native typing could not be synced.
    if (!hasRangeSelection() && !isTrackingRevisions(model) && caretInModelRun()) return;
    endTyping();
    imeRange = model.selection;
  }

  function onCompositionEnd(e: CompositionEvent): void {
    composing = false;
    const range = imeRange;
    if (!range) {
      scheduleRelayout();
      return;
    }
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

  function layoutOptions() {
    return { showGridlines: session.showGridlines, editingHeaderFooter };
  }

  /** Tell the session what the layout produced (page count, page per paragraph …). */
  function reportLayout(): void {
    const pages = layout?.pages ?? [];
    session.pageCount = Math.max(1, pages.length);
    session.pages = pages.map((p) => ({ section: p.section, kind: p.kind }));
    let index: Map<WmlParagraph, number> | undefined;
    session.pageOfParagraph =
      pages.length === 0
        ? undefined
        : (paragraph) => {
            if (!index) {
              index = new Map();
              model.document.document.body.blocks.forEach((b, i) => {
                if (b.kind === 'paragraph') index?.set(b, i);
              });
            }
            const block = index.get(paragraph);
            return block === undefined ? 1 : (layout?.pageOfBlock(block) ?? 0) + 1;
          };
    labelHeaderFooterZones();
    updateCurrentPage();
  }

  const ZONE_LABELS = {
    header: { first: 'hf.firstHeader', even: 'hf.evenHeader', odd: 'hf.oddHeader', default: 'hf.header' },
    footer: { first: 'hf.firstFooter', even: 'hf.evenFooter', odd: 'hf.oddFooter', default: 'hf.footer' },
  } as const;

  /** Word's tab on the header/footer boundary: "First Page Header -Section 2-" … */
  function labelHeaderFooterZones(): void {
    if (!canvas || !editingHeaderFooter || !layout) return;
    const multiSection = layout.sections.length > 1;
    for (const zone of canvas.querySelectorAll<HTMLElement>('.wk-hf')) {
      const labels = ZONE_LABELS[zone.dataset.kind === 'footer' ? 'footer' : 'header'];
      const type = zone.dataset.pageKind;
      // With different odd & even pages, Word names the default one "Odd Page".
      const key =
        type === 'first' ? labels.first : type === 'even' ? labels.even : layout.evenAndOddHeaders ? labels.odd : labels.default;
      const name = t(key);
      zone.dataset.label = multiSection ? `${name} -${t('hf.section')} ${Number(zone.dataset.section) + 1}-` : name;
    }
  }

  // Re-render only on structural edits (session version, canvas renderTick)
  // and view changes, never on every keystroke — that would reset the caret.
  // Character typing is reconciled back into the AST by syncFromDom() on
  // input. After a structural re-render the caret is restored to the model's
  // selection so Enter / Backspace / undo don't dump it at the top.
  $effect(() => {
    if (!canvas || !galley) return;
    const stamp = `${session.version}.${renderTick}`;
    const view = session.viewMode;
    const options = layoutOptions();
    untrack(() => {
      if (!canvas || !galley) return;
      canvas.dataset.version = stamp;
      if (keepBurstAcrossRender) keepBurstAcrossRender = false;
      else endTyping();
      layout ??= new PageLayout(canvas, galley);
      if (view === 'print') layout.build(model.document, options);
      else layout.buildFlow(model.document);
      reportLayout();
      restoreCaret();
      decorateTableSelection(canvas, model);
    });
  });

  function restoreCaret(): void {
    const sel = model.selection;
    if (!sel || !canvas) return;
    const prefer = sel.focus.story && editingHeaderFooter ? headerFooterPage : undefined;
    const focus = domPoint(canvas, sel.focus, prefer);
    if (!focus) return;
    const anchor = domPoint(canvas, sel.anchor, prefer) ?? focus;
    const range = document.createRange();
    range.setStart(anchor.node, anchor.offset);
    range.setEnd(focus.node, focus.offset);
    const domSel = window.getSelection();
    if (!domSel) return;
    domSel.removeAllRanges();
    domSel.addRange(range);
  }

  /** Push the edited paragraph's run text from the DOM back into the AST (no re-render). */
  function syncFromDom(): DocPosition | null {
    if (!canvas) return null;
    const para = focusParagraph();
    if (!para || !canvas.contains(para)) return null;
    for (const span of para.querySelectorAll<HTMLElement>('.wk-run[data-wk-block]')) {
      // Same anchor parsing as DOM-selection mapping, so typing and the caret
      // can never disagree about which run a span is.
      const pos = positionFromDom(span, 0);
      if (!pos) continue;
      const run = runAtPath(model.doc, pos);
      if (!run) continue;
      const text = (span.textContent ?? '').replace(/​/g, '');
      setSimpleRunText(run, text);
    }
    const pos = positionFromDom(para, 0);
    if (pos) mirrorParagraph(canvas, para, pos);
    return pos;
  }

  // Typing re-paginates once the burst pauses, from the edited block on.
  const RELAYOUT_DELAY_MS = 150;
  let relayoutTimer: ReturnType<typeof setTimeout> | undefined;
  let pendingBlock: number | undefined;

  function scheduleRelayout(block?: number): void {
    if (block !== undefined) pendingBlock = block;
    clearTimeout(relayoutTimer);
    relayoutTimer = setTimeout(() => {
      const target = pendingBlock;
      pendingBlock = undefined;
      if (composing || imeRange || target === undefined || !layout || session.viewMode !== 'print') return;
      if (!layout.remeasure(target)) return;
      layout.distribute(layoutOptions());
      reportLayout();
      restoreCaret();
    }, RELAYOUT_DELAY_MS);
  }

  function onInput(): void {
    if (imeRange) return;
    if (!typingOpen) {
      model.beginEdit();
      typingOpen = true;
    }
    const pos = syncFromDom();
    model.commit();
    if (pos && !pos.story) scheduleRelayout(pos.block);
    session.edited();
    session.status = '';
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
      session.status = `${cmd.label}: ${(err as Error).message}`;
      return false;
    }
    rerender();
    session.edited();
    session.status = '';
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

  function updateCurrentPage(): void {
    const page = focusPage();
    if (page !== undefined) session.currentPage = page + 1;
  }

  function onSelChange(): void {
    if (imeRange) return;
    const sel = readDomSelection(document);
    // A caret the browser left between paragraphs (on a page box, between
    // pages) has no document position, and typing there would be lost: put
    // it back where the model has it.
    if (!sel && caretOutsideParagraphs()) restoreCaret();
    if (sel) {
      // A selection never spans two stories (body and a footnote, say):
      // keep the end that moved.
      const next = sameStory(sel.anchor, sel.focus) ? sel : caretAt(sel.focus);
      // Leaving pending formatting untyped removes its run, which shifts the
      // run indices the DOM still carries: re-render after it.
      const released = imeRange || composing ? undefined : releasePendingFormat(model, next);
      model.setSelection(released ?? next);
      if (released) rerender();
      if (sel.focus.story && editingHeaderFooter) headerFooterPage = focusPage();
      if (canvas) decorateTableSelection(canvas, model);
    }
    updateCurrentPage();
    session.tick++;
  }

  function caretOutsideParagraphs(): boolean {
    const domSel = window.getSelection();
    const node = domSel?.focusNode;
    if (!domSel?.isCollapsed || !node || !canvas?.contains(node)) return false;
    const el = node instanceof HTMLElement ? node : node.parentElement;
    return !el?.closest('.wk-p');
  }

  /**
   * Enter in an empty list item steps out of the list as in Word: a nested
   * item moves up a level, a top-level one stops being a list item.
   */
  function leaveEmptyListItem(): boolean {
    const sel = model.selection;
    const para = sel && orderSelection(sel).collapsed ? paragraphAt(model.doc, sel.focus) : undefined;
    const list = para && getParagraphNumbering(para);
    if (!para || !list || !isEmptyParagraph(para)) return false;
    if (list.ilvl > 0) exec(commands.setListLevelCommand, { ilvl: list.ilvl - 1 });
    else exec(commands.removeListCommand, undefined);
    return true;
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
    if (!domSel || !domSel.isCollapsed) return false;
    const pos = model.selection?.focus;
    const para = focusParagraph();
    if (!pos || !para) return false;
    const runs = para.querySelectorAll<HTMLElement>('.wk-run');
    const host = domSel.focusNode instanceof HTMLElement ? domSel.focusNode : domSel.focusNode?.parentElement;
    const last = runs[runs.length - 1];
    const isLast = runs.length === 0 || (!!host && !!last && last.contains(host));
    const len = (last?.textContent ?? '').replace(/​/g, '').length;
    return isLast && (pos.offset ?? 0) >= len;
  }

  /** Whether the paragraph a position names is rendered (in its own story). */
  function paragraphExists(pos: DocPosition): boolean {
    return !!canvas && copiesOf(canvas, pos, `.wk-p${paragraphSelector(pos)}`).length > 0;
  }

  function onKeydown(e: KeyboardEvent): void {
    // While an IME is composing (Japanese/Chinese input), Enter confirms the
    // conversion and Backspace edits the candidate — they must not reach the
    // paragraph commands. Safari reports the composing keydown only through
    // the legacy keyCode 229.
    if (e.isComposing || e.keyCode === 229) return;
    // A key after a range composition must act on the applied result.
    flushImeFinalize();
    // selectionchange fires asynchronously: a key pressed right after a click
    // or an arrow key would otherwise act on the previous caret.
    onSelChange();
    if (CARET_KEYS.has(e.key)) endTyping();
    if (jumpCaret(e)) return;
    if (handleTableKey(e, model, exec)) {
      rerender();
      session.tick++;
      return;
    }
    const mod = e.ctrlKey || e.metaKey;

    if (e.key === 'Escape' && editingHeaderFooter) {
      e.preventDefault();
      session.closeHeaderFooter();
      return;
    }

    // Undo / redo.
    if (mod && e.key.toLowerCase() === 'z') {
      e.preventDefault();
      history(e.shiftKey ? 'redo' : 'undo');
      return;
    }
    if (mod && e.key.toLowerCase() === 'y') {
      e.preventDefault();
      history('redo');
      return;
    }

    // Enter splits the paragraph at the caret (Shift+Enter = soft line break).
    // Tab types a tab character; at the start of a list item it moves the
    // item down a level (Shift+Tab up), as in Word.
    if (e.key === 'Tab' && !mod && !e.altKey) {
      e.preventDefault();
      const focus = model.selection?.focus;
      const para = focus && paragraphAt(model.doc, focus);
      if (para && getParagraphNumbering(para) && caretAtParagraphStart()) {
        exec(commands.indentStepCommand, { direction: e.shiftKey ? 'decrease' : 'increase' });
      } else if (!e.shiftKey) {
        exec(commands.insertTextCommand, { text: '\t' });
      }
      return;
    }

    if (e.key === 'Enter') {
      e.preventDefault();
      if (e.shiftKey) exec(commands.insertLineBreakCommand, { kind: 'line' });
      else if (!leaveEmptyListItem()) exec(commands.splitParagraphCommand, undefined);
      return;
    }

    // Backspace at paragraph start: a list item first loses its number (as in
    // Word), anything else merges into the previous paragraph. Never native:
    // the browser would join DOM nodes (across cells, into a table) on its own.
    if (e.key === 'Backspace' && caretAtParagraphStart()) {
      e.preventDefault();
      const para = model.selection && paragraphAt(model.doc, model.selection.focus);
      if (para && getParagraphNumbering(para)) exec(commands.removeListCommand, undefined);
      else if (commands.mergeBackCommand.isEnabled?.(model) ?? false) exec(commands.mergeBackCommand, undefined);
      return;
    }

    // Delete at paragraph end pulls the next paragraph up into this one.
    if (e.key === 'Delete' && caretAtParagraphEnd()) {
      const here = model.selection?.focus;
      if (here?.cell) {
        // Inside a cell only the next paragraph *of the same cell* may be
        // pulled up; at the cell's last paragraph Delete does nothing (as in
        // Word) instead of letting the browser merge DOM nodes on its own.
        e.preventDefault();
        const next = { ...here, para: (here.para ?? 0) + 1, inline: 0, offset: 0 };
        if (paragraphExists(next)) {
          model.setSelection({ anchor: next, focus: next });
          if (!exec(commands.mergeBackCommand, undefined)) model.setSelection({ anchor: here, focus: here });
        }
        return;
      }
      if (here) {
        // Only a top-level paragraph can be pulled up (not a table / raw block).
        const next = { ...(here.story ? { story: here.story } : {}), block: here.block + 1 };
        if (paragraphExists(next)) {
          e.preventDefault();
          // Merge the next paragraph into this one, keeping the caret at the join.
          model.setSelection({ anchor: next, focus: next });
          if (!exec(commands.mergeBackCommand, undefined)) model.setSelection({ anchor: here, focus: here });
          return;
        }
      }
      // Before a table (or at the end of the document) there is nothing to
      // pull up, as in Word.
      e.preventDefault();
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

  // Word's caret jumps (⌘ on the Mac, Ctrl elsewhere). Left to the browser,
  // ⌘↑ / ⌘↓ put the caret outside every paragraph, where typing is lost.
  const JUMPS: Readonly<Record<string, ParagraphJump>> = {
    ArrowUp: 'previous',
    ArrowDown: 'next',
    Home: 'first',
    End: 'last',
  };

  function jumpCaret(e: KeyboardEvent): boolean {
    // Only the platform's own modifier: on the Mac, Ctrl+A is the line-start
    // key the browser already handles.
    const mod = (/Mac|iPhone|iPad/.test(navigator.platform) ? e.metaKey : e.ctrlKey) && !e.altKey;
    const domSel = window.getSelection();
    const from = focusParagraph();
    if (!mod || !domSel?.focusNode || !from || !canvas) return false;
    // The browser's Select All takes the whole page box, which maps to no
    // document position.
    if (e.key.toLowerCase() === 'a' && !e.shiftKey) {
      const all = storyRange(canvas, from);
      e.preventDefault();
      if (all) domSel.setBaseAndExtent(all.start.node, all.start.offset, all.end.node, all.end.offset);
      onSelChange();
      return true;
    }
    const jump = JUMPS[e.key];
    if (!jump) return false;
    const to = paragraphJumpPoint(canvas, from, { node: domSel.focusNode, offset: domSel.focusOffset }, jump);
    e.preventDefault();
    if (!to) return true;
    if (e.shiftKey) domSel.extend(to.node, to.offset);
    else domSel.collapse(to.node, to.offset);
    (to.node instanceof HTMLElement ? to.node : to.node.parentElement)?.scrollIntoView({ block: 'nearest' });
    onSelChange();
    return true;
  }

  /**
   * Whether native typing at the caret lands in a rendered run the model has,
   * the only text syncFromDom reads back. An empty paragraph saved by Word
   * (`<w:p/>`) has no run, so its typing must go through a command.
   */
  function caretInModelRun(): boolean {
    const node = window.getSelection()?.focusNode;
    const el = node instanceof HTMLElement ? node : node?.parentElement;
    const span = el?.closest<HTMLElement>('.wk-run[data-wk-block]');
    const pos = span ? positionFromDom(span, 0) : null;
    return !!pos && !!runAtPath(model.doc, pos);
  }

  function hasRangeSelection(): boolean {
    const sel = model.selection;
    return !!sel && !orderSelection(sel).collapsed;
  }

  function history(step: 'undo' | 'redo'): void {
    // Typed text is already in the AST (onInput syncs every keystroke), so
    // undo/redo can act on the model directly.
    endTyping();
    if (step === 'undo') model.undo();
    else model.redo();
    rerender();
    session.edited();
  }

  // The edits the browser may make itself: plain typing, IME composition, and
  // a one-character delete (when it stays inside one run, see nativeDeleteFits).
  const NATIVE_INPUTS: ReadonlySet<string> = new Set([
    'insertText',
    'insertCompositionText',
    'deleteCompositionText',
    'insertFromComposition',
    'deleteContentBackward',
    'deleteContentForward',
  ]);

  /**
   * Every other input — Enter from a virtual keyboard or voice input, a word
   * or line delete (⌥⌫, ⌘⌫), a spelling suggestion, a text drop, Undo from
   * the browser's menu — becomes the matching command: the browser's own
   * handling would restructure the DOM across runs and paragraphs in ways the
   * per-run text sync cannot map back.
   */
  function applyInput(e: InputEvent): void {
    onSelChange();
    switch (e.inputType) {
      case 'insertParagraph':
        if (!leaveEmptyListItem()) exec(commands.splitParagraphCommand, undefined);
        return;
      case 'insertLineBreak':
        exec(commands.insertLineBreakCommand, { kind: 'line' });
        return;
      case 'historyUndo':
        history('undo');
        return;
      case 'historyRedo':
        history('redo');
        return;
      case 'insertFromDrop':
        if (e.dataTransfer && selectTargetRange(e)) insertTransfer(e.dataTransfer);
        return;
      case 'insertReplacementText':
      case 'insertFromYank': {
        const text = e.dataTransfer?.getData('text/plain') ?? e.data;
        if (text && selectTargetRange(e)) exec(commands.insertTextCommand, { text });
        return;
      }
    }
    if (e.inputType.startsWith('delete') && selectTargetRange(e) && hasRangeSelection()) {
      exec(commands.deleteSelectionCommand, undefined);
    }
  }

  /** Put the model selection on the range an input event is about to change. */
  function selectTargetRange(e: InputEvent): boolean {
    const [range] = e.getTargetRanges();
    if (!range) return true;
    const anchor = positionFromDom(range.startContainer, range.startOffset);
    const focus = positionFromDom(range.endContainer, range.endOffset);
    if (!anchor || !focus) return false;
    model.setSelection({ anchor, focus });
    return true;
  }

  /**
   * Whether the browser may make a one-character delete itself: it removes
   * text from one text node of a run the model has, and leaves text there
   * (a run emptied by the browser loses its span, and with it the sync).
   */
  function nativeDeleteFits(e: InputEvent): boolean {
    const [range] = e.getTargetRanges();
    const node = range?.startContainer;
    if (!range || !(node instanceof Text) || range.endContainer !== node || !caretInModelRun()) return false;
    const span = node.parentElement?.closest('.wk-run');
    const left = (span?.textContent ?? '').replace(/\u200b/g, '').length - (range.endOffset - range.startOffset);
    return left > 0;
  }

  /**
   * Typing / deleting over a *range* is routed to commands, like every input
   * outside NATIVE_INPUTS. Collapsed-caret typing stays native (reconciled by
   * onInput) so the caret and IME behave normally.
   */
  function onBeforeInput(e: InputEvent): void {
    if (!NATIVE_INPUTS.has(e.inputType)) {
      e.preventDefault();
      // Enter confirming a composition is the IME's, not a paragraph break.
      if (!composing && !imeRange && !e.isComposing) applyInput(e);
      return;
    }
    // While tracking changes every edit is recorded by a command, never synced natively.
    if (imeRange || e.isComposing) return;
    // Mouse selections reach here without a keydown to sync them.
    onSelChange();
    // A selection the browser placed outside every paragraph (on a page box)
    // has no document position: editing it natively would wreck the pages.
    if (!readDomSelection(document)) {
      e.preventDefault();
      return;
    }
    const nativeOk = e.inputType.startsWith('delete') ? nativeDeleteFits(e) : caretInModelRun();
    if (!hasRangeSelection() && !isTrackingRevisions(model) && nativeOk) return;
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
      if (isTrackingRevisions(model) && !hasRangeSelection()) {
        exec(commands.trackedDeleteCommand, { direction: e.inputType.endsWith('Forward') ? 1 : -1 });
      } else if (hasRangeSelection() || (selectTargetRange(e) && hasRangeSelection())) {
        exec(commands.deleteSelectionCommand, undefined);
      }
    }
  }

  function onCut(e: ClipboardEvent): void {
    // A menu Cut or Paste comes without a keydown to sync the selection.
    onSelChange();
    if (!hasRangeSelection() || !e.clipboardData) return;
    e.preventDefault();
    e.clipboardData.setData('text/plain', window.getSelection()?.toString() ?? '');
    exec(commands.deleteSelectionCommand, undefined);
  }

  function onPaste(e: ClipboardEvent): void {
    flushImeFinalize();
    onSelChange();
    if (!e.clipboardData) return;
    e.preventDefault();
    insertTransfer(e.clipboardData);
  }

  /** Paste or drop: formatted when the data has HTML, else its plain text. */
  function insertTransfer(data: DataTransfer): void {
    const html = data.getData('text/html');
    const blocks = html ? parseClipboardHtml(html) : [];
    if (blocks.length > 0) exec(commands.insertFragmentCommand, { blocks });
    else {
      const text = data.getData('text/plain');
      if (text) exec(commands.insertTextCommand, { text });
    }
  }

  /** Header or footer zone of the page under a point, by Word's margins. */
  function marginZoneAt(e: MouseEvent): { kind: 'header' | 'footer'; page: number } | undefined {
    const box = (e.target as Element).closest<HTMLElement>('.wk-pagebox');
    if (!box) return undefined;
    const rect = box.getBoundingClientRect();
    const zoom = rect.height / box.offsetHeight || 1;
    const y = (e.clientY - rect.top) / zoom;
    const css = getComputedStyle(box);
    const top = parseFloat(css.getPropertyValue('--m-top'));
    const bottom = parseFloat(css.getPropertyValue('--m-bottom'));
    const page = Number(box.dataset.page);
    if (y < top) return { kind: 'header', page };
    if (y > box.offsetHeight - bottom) return { kind: 'footer', page };
    return undefined;
  }

  /**
   * A single click in the header or footer margin while editing the body puts
   * the caret on the nearest body line, as in Word. The header is not
   * editable then, so the browser would leave the caret where it was.
   */
  function onMouseDown(e: MouseEvent): void {
    endTyping();
    if (session.viewMode !== 'print' || editingHeaderFooter || e.button !== 0 || e.shiftKey) return;
    // Body content placed in the margin (a picture) takes its own click.
    if ((e.target as Element).closest('.wk-body')) return;
    const zone = marginZoneAt(e);
    const body = (e.target as Element).closest('.wk-pagebox')?.querySelector<HTMLElement>('.wk-body');
    const lines = body ? [...body.querySelectorAll<HTMLElement>('.wk-p')] : [];
    const line = zone?.kind === 'header' ? lines[0] : lines.at(-1);
    if (!zone || !line) return;
    const rect = line.getBoundingClientRect();
    const y = Math.min(Math.max(e.clientY, rect.top + 1), rect.bottom - 1);
    const x = Math.min(Math.max(e.clientX, rect.left + 1), rect.right - 1);
    const hit = document.caretPositionFromPoint(x, y);
    // Off screen (a long page scrolled to its foot) the line has no hit
    // point: take its end nearest the click instead.
    const at =
      hit && line.contains(hit.offsetNode)
        ? { node: hit.offsetNode, offset: hit.offset }
        : edgePoint(line, zone.kind === 'footer');
    e.preventDefault();
    canvas?.focus({ preventScroll: true });
    window.getSelection()?.collapse(at.node, at.offset);
    onSelChange();
  }

  /** Double-click in a header/footer area opens it; in the body, closes it (Word). */
  function onDblClick(e: MouseEvent): void {
    if (session.viewMode !== 'print') return;
    const zone = marginZoneAt(e);
    if (!editingHeaderFooter && zone) {
      e.preventDefault();
      headerFooterPage = zone.page;
      session.editHeaderFooter(zone.kind, zone.page);
      return;
    }
    if (editingHeaderFooter && !zone && (e.target as Element).closest('.wk-body')) {
      e.preventDefault();
      session.closeHeaderFooter();
    }
  }

  /** Clicking a note reference goes to the note; clicking the note's own mark goes back. */
  function onClick(e: MouseEvent): void {
    const mark = (e.target as Element).closest<HTMLElement>('.wk-noteref');
    if (!mark || !canvas) return;
    const key = mark.dataset.wkNote;
    if (key) {
      const story = parseStoryKey(key);
      if (!story) return;
      e.preventDefault();
      model.setSelection(caretAt({ story, block: 0, inline: 0, offset: 0 }));
      restoreCaret();
      window.getSelection()?.focusNode?.parentElement?.scrollIntoView({ block: 'nearest' });
      session.tick++;
      return;
    }
    const own = mark.closest<HTMLElement>('[data-wk-story]')?.dataset.wkStory;
    const reference = own ? canvas.querySelector<HTMLElement>(`.wk-noteref[data-wk-note="${own}"]`) : null;
    if (!reference) return;
    e.preventDefault();
    const range = document.createRange();
    range.setStartAfter(reference);
    range.collapse(true);
    const sel = window.getSelection();
    sel?.removeAllRanges();
    sel?.addRange(range);
    reference.scrollIntoView({ block: 'center' });
  }

  $effect(() => {
    document.addEventListener('selectionchange', onSelChange);
    return () => {
      document.removeEventListener('selectionchange', onSelChange);
      clearTimeout(relayoutTimer);
    };
  });
</script>

<div
  class="wk-page"
  class:side={session.pageMovement === 'sideToSide'}
  data-view={session.viewMode}
  data-table-tool={tableToolAttr()}
  style="zoom: {session.zoom}; {tableCanvasStyle()}"
>
  <div
    bind:this={canvas}
    class="wk-canvas"
    class:hf-editing={editingHeaderFooter}
    class:hide-body={editingHeaderFooter && !session.showDocumentText}
    data-view={session.viewMode}
    contenteditable={!session.readOnly}
    spellcheck={session.spellcheck}
    role="textbox"
    tabindex="0"
    aria-multiline="true"
    aria-label={t('canvas.label')}
    oninput={onInput}
    onkeydown={onKeydown}
    onmousedown={onMouseDown}
    onpointerdown={(e) => tablePointerDown(e, model, exec)}
    onpointerup={(e) => tablePointerUp(e, model, exec)}
    onblur={endTyping}
    onpaste={onPaste}
    onbeforeinput={onBeforeInput}
    oncut={onCut}
    ondblclick={onDblClick}
    onclick={onClick}
    oncompositionstart={onCompositionStart}
    oncompositionend={onCompositionEnd}
  ></div>
  <!-- Off-screen measuring area: blocks are rendered here, measured, then moved into pages. -->
  <div bind:this={galley} class="wk-galley wk-canvas" aria-hidden="true"></div>
</div>
<ObjectSelection {canvas} />
<ShapeOverlay {canvas} />

<style>
  .wk-page {
    position: relative;
    display: flex;
    /* `safe`: a side-to-side row wider than the surface starts at its left
       edge instead of overflowing off-screen on both sides. */
    justify-content: safe center;
  }

  .wk-canvas {
    outline: none;
    color: #000;
    /* Only for the caret in an empty document; every paragraph and run carries
       its style-resolved font and size. */
    font-family: Calibri, Carlito, 'Segoe UI', system-ui, sans-serif;
    font-size: 11pt;
    line-height: 1.2;
  }

  /* Measuring area: laid out like the page but never seen or hit. */
  .wk-galley {
    position: absolute;
    top: 0;
    left: -100000px;
    visibility: hidden;
    pointer-events: none;
  }
  .wk-galley :global(.wk-galley-col) {
    display: flow-root;
  }

  /* --- Print Layout: one box per page --------------------------------- */
  .wk-canvas :global(.wk-pages) {
    display: flex;
    flex-direction: column;
    align-items: center;
    gap: 16px;
    padding: 16px 0;
  }
  .side .wk-canvas :global(.wk-pages) {
    flex-direction: row;
    align-items: flex-start;
    padding: 16px;
  }

  .wk-canvas :global(.wk-pagebox) {
    /* The positioning context for anchored objects (see --page-* / --m-*);
       isolated so behind-text pictures stay above the page background. */
    position: relative;
    isolation: isolate;
    display: grid;
    flex: none;
    box-sizing: border-box;
    background-color: #fff;
    border: 1px solid #c6c6c6;
    box-shadow: 0 1px 6px rgba(0, 0, 0, 0.12);
    /* Offscreen pages skip rendering work; a 200-page document stays responsive. */
    content-visibility: auto;
    contain-intrinsic-size: auto var(--page-w) auto var(--page-h);
  }
  /* Header, body and footer share one cell, placed by margins. */
  .wk-canvas :global(.wk-pagebox > .wk-hf),
  .wk-canvas :global(.wk-pagebox > .wk-body) {
    grid-area: 1 / 1;
    box-sizing: border-box;
    min-width: 0;
  }
  .wk-canvas :global(.wk-header) {
    align-self: start;
  }
  .wk-canvas :global(.wk-footer) {
    align-self: end;
    display: flex;
    flex-direction: column;
    justify-content: flex-end;
  }
  .wk-canvas :global(.wk-body) {
    align-self: start;
    display: grid;
    z-index: 1;
  }
  .wk-canvas :global(.wk-regions),
  .wk-canvas :global(.wk-notes) {
    grid-area: 1 / 1;
    min-width: 0;
  }
  .wk-canvas :global(.wk-regions) {
    display: flow-root;
  }
  .wk-canvas :global(.wk-region) {
    display: flex;
  }
  .wk-canvas :global(.wk-col) {
    flex: none;
    display: flow-root;
  }
  .wk-canvas :global(.wk-colgap) {
    flex: none;
  }
  .wk-canvas :global(.wk-colgap.sep) {
    background: linear-gradient(#000, #000) center / 1px 100% no-repeat;
  }
  /* 縦書き: lines run top to bottom and follow each other right to left. */
  .wk-canvas :global(.wk-vertical),
  .wk-galley :global(.wk-vertical) {
    writing-mode: vertical-rl;
  }
  .wk-canvas :global(.wk-vertical .wk-colgap.sep) {
    background-size: 100% 1px;
  }
  .wk-canvas :global(.wk-blk) {
    display: flow-root;
  }
  .wk-canvas :global(.wk-slice) {
    overflow: hidden;
    display: flow-root;
  }
  .wk-canvas :global(.wk-notes) {
    align-self: end;
    overflow: hidden;
  }
  .wk-canvas :global(.wk-note-sep) {
    width: 33%;
    margin: 6px 0 7px;
    border-top: 1px solid #000;
    user-select: none;
  }
  .wk-canvas :global(.wk-noteref) {
    font-size: 0.65em;
    vertical-align: super;
    line-height: 0;
    cursor: pointer;
    user-select: none;
  }

  /* Decorations drawn on the page, never part of the text. */
  .wk-canvas :global(.wk-crop) {
    position: absolute;
    width: 13.5pt;
    height: 13.5pt;
    box-sizing: border-box;
    border: 1px none #a6a6a6;
    pointer-events: none;
  }
  .wk-canvas :global(.wk-pgborder) {
    position: absolute;
    pointer-events: none;
    box-sizing: border-box;
  }
  .wk-canvas :global(.wk-pgborder.back) {
    z-index: 0;
  }
  .wk-canvas :global(.wk-pgborder.front) {
    z-index: 2;
  }
  .wk-canvas :global(.wk-grid) {
    position: absolute;
    pointer-events: none;
    background-image: repeating-linear-gradient(
      to bottom,
      transparent 0,
      transparent calc(var(--pitch) - 1px),
      rgba(0, 0, 0, 0.12) calc(var(--pitch) - 1px),
      rgba(0, 0, 0, 0.12) var(--pitch)
    );
  }
  .wk-canvas :global(.wk-linenum) {
    position: absolute;
    text-align: right;
    font-size: 10pt;
    color: #000;
    user-select: none;
    pointer-events: none;
  }

  /* Header & footer editing: the body dims and a labelled boundary appears. */
  .hf-editing :global(.wk-body) {
    opacity: 0.45;
  }
  .hide-body :global(.wk-body) {
    visibility: hidden;
  }
  .hf-editing {
    --hf-label-h: 18px;
  }
  .hf-editing :global(.wk-header) {
    border-bottom: 1px dashed #5b8fd6;
  }
  .hf-editing :global(.wk-footer) {
    border-top: 1px dashed #5b8fd6;
  }
  /* The labels stay in flow, cancelled out by a negative margin, so they hang
     past the zone's edge without making the zone a containing block: anchored
     header shapes (watermarks) are positioned against the page. */
  .hf-editing :global(.wk-header::after),
  .hf-editing :global(.wk-footer::before) {
    content: attr(data-label);
    display: block;
    box-sizing: border-box;
    width: max-content;
    height: var(--hf-label-h);
    padding: 1px 6px;
    background: #dfe8f6;
    border: 1px solid #5b8fd6;
    color: #1f3f74;
    font: 11px -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
    user-select: none;
    pointer-events: none;
  }
  .hf-editing :global(.wk-header::after) {
    margin-bottom: calc(-1 * var(--hf-label-h));
  }
  .hf-editing :global(.wk-footer::before) {
    margin-top: calc(-1 * var(--hf-label-h));
  }

  /* --- Other views: one continuous flow ------------------------------- */
  .wk-canvas :global(.wk-flow) {
    box-sizing: border-box;
    width: min(100%, 816px);
    min-height: 100%;
    margin: 16px auto;
    padding: 24px 48px;
    background: #fff;
  }

  /* --- Rendered content ------------------------------------------------ */
  .wk-canvas :global(.wk-p) {
    margin: 0;
    min-height: 1.2em;
    /* Runs render tabs and <w:br/> as \t / \n; keep them visible. */
    white-space: pre-wrap;
  }

  /* layoutTabStops sizes each tab to its stop; the leader is text, so it
     takes the run's font like Word's does, clipped to the tab's width. */
  .wk-canvas :global(.wk-tab) {
    display: inline-block;
    clip-path: inset(0);
    text-indent: 0;
  }

  .wk-canvas :global(.wk-tab[data-leader])::before {
    white-space: pre;
  }

  .wk-canvas :global(.wk-tab[data-leader='dot'])::before {
    content: '................................................................................................................................................................................................................................................................................................................................................................................................................';
  }

  .wk-canvas :global(.wk-tab[data-leader='hyphen'])::before {
    content: '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------';
  }

  .wk-canvas :global(.wk-tab[data-leader='underscore'])::before,
  .wk-canvas :global(.wk-tab[data-leader='heavy'])::before {
    content: '________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________';
  }

  .wk-canvas :global(.wk-tab[data-leader='heavy'])::before {
    font-weight: 700;
  }

  .wk-canvas :global(.wk-tab[data-leader='middleDot'])::before {
    content: '················································································································································································································································································································································································································';
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
