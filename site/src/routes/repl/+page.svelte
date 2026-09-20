<script lang="ts">
  import { onMount } from 'svelte';
  import type { ValidationIssue } from '@office-kit/docx';
  import { javascript } from '@codemirror/lang-javascript';
  import { syntaxTree } from '@codemirror/language';
  import { EditorState } from '@codemirror/state';
  import { oneDark } from '@codemirror/theme-one-dark';
  import { EditorView, basicSetup } from 'codemirror';
  import DocumentPreview from '$lib/components/DocumentPreview.svelte';
  import ValidationReport from '$lib/components/ValidationReport.svelte';
  import { downloadDocx } from '$lib/download';
  import starterSource from '$lib/examples/repl-default.ts?raw';

  // The starter is a type-checked module; the editor shows it from this
  // comment on, without the imports the type checker needs and the REPL does not.
  const STARTER_MARKER = '// Every @office-kit/docx function is in scope.';
  const DEFAULT_CODE = starterSource.slice(starterSource.indexOf(STARTER_MARKER));

  const DOWNLOAD_NAME = 'office-kit-docx-repl.docx';
  const RUN_DEBOUNCE_MS = 250;

  let code = $state(DEFAULT_CODE);
  let error = $state('');
  let bytes = $state<Uint8Array | null>(null);
  let issues = $state<ValidationIssue[]>([]);
  let busy = $state(false);
  // The first run waits for mount so prerendering never evaluates the code.
  let mounted = $state(false);

  let editorContainer = $state<HTMLDivElement | undefined>();
  let view: EditorView | undefined;

  onMount(() => {
    if (editorContainer) {
      view = new EditorView({
        state: EditorState.create({
          doc: code,
          extensions: [
            basicSetup,
            // Plain JavaScript, because that is what `new Function` runs: the
            // parser then flags a stray type annotation like any other typo.
            javascript(),
            // One Dark supplies the syntax colours; the surfaces are ours so
            // the editor matches every other code panel on the site. Ours come
            // first because CodeMirror gives earlier extensions precedence.
            EditorView.theme(
              {
                '&': { height: '100%', fontSize: '13px', backgroundColor: 'var(--night)' },
                '.cm-gutters': { backgroundColor: 'var(--night)', border: 'none' },
                '.cm-activeLine, .cm-activeLineGutter': { backgroundColor: 'var(--night-2)' },
                '.cm-scroller': { fontFamily: 'var(--mono)', overflow: 'auto' },
              },
              { dark: true },
            ),
            oneDark,
            EditorView.updateListener.of((update) => {
              if (update.docChanged) code = update.state.doc.toString();
            }),
          ],
        }),
        parent: editorContainer,
      });
    }
    mounted = true;
    return () => view?.destroy();
  });

  let runTimer: ReturnType<typeof setTimeout> | undefined;
  $effect(() => {
    // Reading `code` subscribes the effect, so every edit schedules a run.
    const source = code;
    if (!mounted) return;
    clearTimeout(runTimer);
    runTimer = setTimeout(() => void run(source), RUN_DEBOUNCE_MS);
  });

  // `new Function` wraps the source in a two-line header, and we add two more
  // lines before the user's code, so a stack frame's line is off by four.
  const WRAPPER_LINES = 4;
  // V8 names the frame `<anonymous>`, SpiderMonkey `Function`.
  const USER_FRAME = /(?:<anonymous>|Function):(\d+):\d+/;

  // A SyntaxError thrown by `new Function` carries no position, so the line
  // comes from the editor's own parse of the same text instead.
  function syntaxErrorLine(): number {
    if (!view) return 0;
    let position = -1;
    syntaxTree(view.state).iterate({
      enter: (node) => {
        if (position < 0 && node.type.isError) position = node.from;
        return position < 0;
      },
    });
    return position < 0 ? 0 : view.state.doc.lineAt(position).number;
  }

  // A minified stack trace is noise to someone writing a few lines of code.
  // Show the message and the line of their code it came from.
  function describeError(err: unknown): string {
    if (!(err instanceof Error)) return String(err);
    const frame = USER_FRAME.exec(err.stack ?? '');
    const line =
      err instanceof SyntaxError ? syntaxErrorLine() : frame ? Number(frame[1]) - WRAPPER_LINES : 0;
    return line > 0 ? `Line ${line}: ${err.message}` : err.message;
  }

  // Edits arrive faster than a run can finish; only the newest may report.
  let generation = 0;

  async function run(source: string): Promise<void> {
    const mine = ++generation;
    busy = true;
    try {
      const kit = await import('@office-kit/docx');
      const doc = kit.createDocx({ paragraphs: [] });
      // Every export becomes a parameter, so the code calls `appendParagraph(…)`
      // the way the docs write it, without an import line or a namespace.
      const names = Object.keys(kit);
      // Wrapped in an async function so `await` is allowed at the top level.
      const fn = new Function(
        ...names,
        'doc',
        `'use strict';\nreturn (async () => {\n${source}\n})();`,
      );
      await fn(...Object.values(kit), doc);

      const saved = kit.toUint8Array(doc);
      // Validate what was written, not the in-memory model it came from.
      const found = kit.validate(kit.openDocx(saved));
      if (mine !== generation) return;
      bytes = saved;
      issues = found;
      error = '';
    } catch (err) {
      // Code is broken for most keystrokes while someone types. Keep the last
      // document that built on screen (dimmed) instead of blanking the preview.
      if (mine === generation) error = describeError(err);
    } finally {
      if (mine === generation) busy = false;
    }
  }

  function copyCode(): void {
    void navigator.clipboard.writeText(code);
  }

  // The editor owns the text; pushing the default through a transaction keeps
  // undo history intact and updates `code` through the change listener.
  function resetCode(): void {
    view?.dispatch({ changes: { from: 0, to: view.state.doc.length, insert: DEFAULT_CODE } });
  }

  function reportPreviewError(err: unknown): void {
    error = `The preview could not draw this document: ${err instanceof Error ? err.message : String(err)}`;
  }
</script>

<svelte:head>
  <title>REPL · @office-kit/docx</title>
</svelte:head>

<section class="content">
  <header class="intro">
    <h1>REPL</h1>
    <p class="lede">
      Write code and the document redraws as you type. Every public function is already in scope,
      and <code>doc</code> is a new, empty document from <code>createDocx</code>. The preview is
      drawn from the bytes <code>toUint8Array</code> wrote, and the download is those same bytes.
      Nothing is uploaded: the code, the library, and the preview all run in this tab.
    </p>
  </header>

  <div class="repl-grid">
    <div class="pane editor-pane">
      <div class="pane-head">
        <h2>Code</h2>
        <div class="pane-actions">
          <button type="button" onclick={resetCode}>Reset</button>
          <button type="button" onclick={copyCode}>Copy</button>
        </div>
      </div>
      <div class="editor" bind:this={editorContainer}></div>
      {#if error}
        <pre class="error" role="alert">{error}</pre>
      {/if}
    </div>

    <div class="pane preview-pane">
      <div class="pane-head">
        <h2>Preview</h2>
        <div class="pane-actions">
          <span class="busy" aria-live="polite">{busy ? 'Building…' : ''}</span>
          <button
            type="button"
            class="strong"
            onclick={() => bytes && downloadDocx(bytes, DOWNLOAD_NAME)}
            disabled={!bytes}
          >
            Download .docx
          </button>
        </div>
      </div>
      <div class="output">
        <DocumentPreview {bytes} stale={error !== ''} onerror={reportPreviewError} />
      </div>
      {#if bytes}
        <div class="validation" class:stale={error !== ''}>
          <ValidationReport {issues} />
        </div>
      {/if}
    </div>
  </div>
</section>

<style>
  .content {
    max-width: 1440px;
    margin: 0 auto;
    padding: 2.25rem var(--gutter) 3rem;
  }

  .intro h1 {
    margin-bottom: 0.6rem;
  }

  .lede {
    max-width: 78ch;
    margin: 0 0 1.5rem;
    color: var(--ink-2);
  }

  .repl-grid {
    display: grid;
    grid-template-columns: minmax(0, 1fr) minmax(0, 1fr);
    gap: 1rem;
    align-items: start;
  }

  .pane {
    display: flex;
    flex-direction: column;
    min-width: 0;
    height: calc(100vh - var(--header-h) - 2rem);
    min-height: 480px;
    border: 1px solid var(--line-strong);
    border-radius: var(--radius);
    overflow: hidden;
  }

  .editor-pane {
    position: sticky;
    top: calc(var(--header-h) + 1rem);
    background: var(--night);
    border-color: var(--night-line);
  }

  .pane-head {
    flex: none;
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 0.75rem;
    min-height: 48px;
    padding: 0 0.6rem 0 1rem;
    border-bottom: 1px solid var(--line);
    background: var(--wash);
  }

  .editor-pane .pane-head {
    background: var(--night);
    border-bottom-color: var(--night-line);
    color: var(--night-ink);
  }

  .pane-head h2 {
    margin: 0;
    font-family: var(--sans);
    font-size: 0.9rem;
    font-weight: 600;
    letter-spacing: 0;
    color: inherit;
  }

  .pane-actions {
    display: flex;
    align-items: center;
    gap: 0.4rem;
  }

  .pane-actions button {
    height: 32px;
    padding: 0 0.75rem;
    border: 1px solid var(--line-strong);
    border-radius: var(--radius-sm);
    background: var(--paper);
    color: var(--ink);
    font-family: var(--sans);
    font-size: 0.85rem;
    font-weight: 550;
    cursor: pointer;
  }

  .pane-actions button:hover:not(:disabled) {
    border-color: var(--ink-3);
  }

  .pane-actions button:disabled {
    opacity: 0.5;
    cursor: not-allowed;
  }

  .editor-pane .pane-actions button {
    background: var(--night-2);
    border-color: var(--night-line);
    color: var(--night-ink);
  }

  .pane-actions button.strong {
    background: var(--ink);
    border-color: var(--ink);
    color: var(--paper);
  }

  .pane-actions button.strong:hover:not(:disabled) {
    background: var(--accent);
    border-color: var(--accent);
    color: var(--on-accent);
  }

  .busy {
    color: var(--ink-3);
    font-size: 0.82rem;
  }

  .editor {
    flex: 1;
    min-height: 0;
    overflow: hidden;
  }

  .editor :global(.cm-editor) {
    height: 100%;
    font-size: 13px;
  }

  .editor :global(.cm-scroller) {
    font-family: var(--mono);
    line-height: 1.6;
  }

  /* An error is red in both themes because it sits on the night surface. */
  .error {
    flex: none;
    max-height: 35%;
    margin: 0;
    border: none;
    border-top: 1px solid #7a2a1a;
    border-radius: 0;
    background: #2a1410;
    color: #ffb4a1;
    font-size: 0.8rem;
    white-space: pre-wrap;
    overflow: auto;
  }

  .output {
    flex: 1;
    min-height: 0;
  }

  .validation {
    flex: none;
    max-height: 30%;
    padding: 0.7rem 1rem;
    border-top: 1px solid var(--line);
    background: var(--paper);
    font-size: 0.9rem;
    overflow-y: auto;
  }

  .validation.stale {
    opacity: 0.45;
  }

  @media (max-width: 900px) {
    .repl-grid {
      grid-template-columns: 1fr;
    }

    .pane {
      height: auto;
      min-height: 0;
    }

    .editor-pane {
      position: static;
    }

    .editor {
      height: 46vh;
      min-height: 280px;
      flex: none;
    }

    .validation {
      max-height: none;
    }
  }
</style>
