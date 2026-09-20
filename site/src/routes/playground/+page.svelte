<script lang="ts">
  import { onMount } from 'svelte';
  import type { ValidationIssue } from '@office-kit/docx';
  import DocumentPreview from '$lib/components/DocumentPreview.svelte';
  import ValidationReport from '$lib/components/ValidationReport.svelte';
  import { downloadDocx } from '$lib/download';

  type Stats = ReturnType<typeof import('@office-kit/docx').statistics>;

  const DOCX_ACCEPT =
    '.docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document';
  const SAMPLE_NAME = 'office-kit-demo.docx';

  let status = $state('Loading the library…');
  let busy = $state(false);
  let dropping = $state(false);
  let fileName = $state('');
  let stats = $state<Stats | null>(null);
  let docTitle = $state<string | undefined>();
  let docAuthor = $state<string | undefined>();
  let issues = $state<ValidationIssue[]>([]);
  // The re-saved bytes, which are also what the preview below is showing.
  let savedBytes = $state<Uint8Array | null>(null);

  // Opens the file with the library, saves it again, and previews the saved
  // bytes rather than the upload: what is on screen is the round trip.
  async function inspect(bytes: Uint8Array, name: string): Promise<void> {
    busy = true;
    fileName = name;
    status = `Opening ${name}…`;
    try {
      const core = await import('@office-kit/docx');
      const doc = core.openDocx(bytes);
      stats = core.statistics(doc);
      docTitle = core.title(doc);
      docAuthor = core.author(doc);
      issues = core.validate(doc);
      savedBytes = core.toUint8Array(doc);
      status = `Opened ${name} (${bytes.byteLength.toLocaleString()} bytes) and saved it again (${savedBytes.byteLength.toLocaleString()} bytes). The preview shows the saved copy.`;
    } catch (err) {
      fail(err);
    } finally {
      busy = false;
    }
  }

  function fail(err: unknown): void {
    stats = null;
    issues = [];
    savedBytes = null;
    status = `This file could not be opened: ${err instanceof Error ? err.message : String(err)}`;
  }

  // Gives a visitor with no .docx at hand something real to inspect: the
  // document from the landing page, built in this tab.
  async function loadSample(): Promise<void> {
    const [{ toUint8Array }, { buildHeroDocument }] = await Promise.all([
      import('@office-kit/docx'),
      import('$lib/examples/hero-document'),
    ]);
    await inspect(toUint8Array(buildHeroDocument()), SAMPLE_NAME);
  }

  async function onFileChosen(file: File): Promise<void> {
    await inspect(new Uint8Array(await file.arrayBuffer()), file.name);
  }

  function onDrop(event: DragEvent): void {
    event.preventDefault();
    dropping = false;
    const file = event.dataTransfer?.files[0];
    if (file) void onFileChosen(file);
  }

  function downloadSaved(): void {
    if (savedBytes) downloadDocx(savedBytes, fileName.replace(/\.docx$/i, '') + '.roundtrip.docx');
  }

  const cells = $derived(
    stats
      ? [
          { label: 'Paragraphs', value: stats.paragraphs },
          { label: 'Headings', value: stats.headings },
          { label: 'Tables', value: stats.tables },
          { label: 'Images', value: stats.images },
          { label: 'Words', value: stats.words.toLocaleString() },
          { label: 'Comments', value: stats.comments },
          { label: 'Footnotes and endnotes', value: stats.footnotes + stats.endnotes },
          { label: 'Title', value: docTitle || 'Not set' },
          { label: 'Author', value: docAuthor || 'Not set' },
        ]
      : [],
  );

  onMount(() => void loadSample());
</script>

<svelte:head>
  <title>Playground · @office-kit/docx</title>
</svelte:head>

<section class="content">
  <h1>Open a .docx in your browser</h1>
  <p class="lede">
    Drop a file and this page opens it with <code>@office-kit/docx</code>, counts what is inside,
    validates the package, saves it again, and renders the saved copy with
    <code>@office-kit/docx-preview</code>. Nothing is uploaded: the whole pipeline runs in this tab.
  </p>

  <div
    class="drop"
    class:dropping
    role="group"
    aria-label="Choose a .docx file"
    ondragover={(e) => {
      e.preventDefault();
      dropping = true;
    }}
    ondragleave={() => (dropping = false)}
    ondrop={onDrop}
  >
    <p class="drop-text">{fileName || 'Drop a .docx file here'}</p>
    <div class="drop-actions">
      <label class="btn primary drop-pick">
        <input
          type="file"
          accept={DOCX_ACCEPT}
          onchange={(e) => {
            const file = e.currentTarget.files?.[0];
            if (file) void onFileChosen(file);
          }}
        />
        Choose a file
      </label>
      <button type="button" class="btn" onclick={loadSample} disabled={busy}>
        Load the sample document
      </button>
      {#if savedBytes}
        <button type="button" class="btn" onclick={downloadSaved}>Download the re-saved file</button>
      {/if}
    </div>
  </div>

  <p class="caveat">
    The preview is drawn by the open-source docx-preview renderer, which approximates Word’s layout:
    pagination, fonts, and floating objects can differ. Word and LibreOffice remain the exact
    renderers.
  </p>

  <p class="status" class:busy aria-live="polite">{status}</p>

  {#if cells.length > 0}
    <div class="meta-clip">
      <dl class="meta">
        {#each cells as cell (cell.label)}
          <div class="cell">
            <dt>{cell.label}</dt>
            <dd>{cell.value}</dd>
          </div>
        {/each}
      </dl>
    </div>
  {/if}

  {#if stats}
    <h2>Validation</h2>
    <ValidationReport {issues} />
  {/if}

  <h2>Preview</h2>
  <div class="preview-frame">
    <DocumentPreview bytes={savedBytes} maxHeight="80vh" onerror={fail} />
  </div>
</section>

<style>
  .content {
    max-width: 1000px;
    margin: 0 auto;
    padding: 2.75rem var(--gutter) 5rem;
  }

  .lede {
    max-width: 66ch;
    color: var(--ink-2);
    font-size: 1.08rem;
  }

  .lede code {
    white-space: nowrap;
  }

  .drop {
    display: flex;
    flex-direction: column;
    align-items: center;
    gap: 1.1rem;
    margin: 2rem 0 0;
    padding: 2.25rem 1.25rem;
    border: 1.5px dashed var(--line-strong);
    border-radius: 12px;
    background: var(--wash);
    text-align: center;
    transition:
      border-color 120ms ease,
      background 120ms ease;
  }

  .drop.dropping {
    border-color: var(--accent);
    background: var(--accent-wash);
  }

  .drop-text {
    margin: 0;
    font-family: var(--display);
    font-size: 1.25rem;
    font-weight: 600;
    letter-spacing: -0.015em;
    overflow-wrap: anywhere;
  }

  .drop-actions {
    display: flex;
    flex-wrap: wrap;
    justify-content: center;
    gap: 0.6rem;
  }

  .drop-pick input {
    position: absolute;
    width: 1px;
    height: 1px;
    opacity: 0;
  }

  .drop-pick:focus-within {
    outline: 2px solid var(--accent);
    outline-offset: 2px;
  }

  .btn:disabled {
    opacity: 0.6;
    cursor: progress;
  }

  .caveat {
    max-width: 72ch;
    margin: 1rem 0 0;
    color: var(--ink-3);
    font-size: 0.88rem;
  }

  .status {
    min-height: 1.6em;
    margin: 1.5rem 0 0;
    color: var(--ink-2);
    font-size: 0.95rem;
  }

  .status.busy {
    color: var(--accent-ink);
  }

  h2 {
    margin: 3rem 0 1rem;
    padding-bottom: 0.6rem;
    border-bottom: 1px solid var(--line);
    font-size: 1.4rem;
  }

  /* Each cell draws its own right and bottom rule and the grid is pulled 1px
   * past the clipping box, so the outer edge never doubles up and a short last
   * row leaves plain paper rather than a filled gap. */
  .meta-clip {
    margin-top: 1rem;
    border: 1px solid var(--line);
    border-radius: var(--radius);
    overflow: hidden;
  }

  .meta {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
    margin: 0 -1px -1px 0;
  }

  .cell {
    display: flex;
    flex-direction: column;
    gap: 0.15rem;
    padding: 0.8rem 1rem;
    border-right: 1px solid var(--line);
    border-bottom: 1px solid var(--line);
  }

  .cell dt {
    color: var(--ink-3);
    font-size: 0.8rem;
  }

  .cell dd {
    margin: 0;
    font-weight: 600;
    font-size: 0.97rem;
    overflow-wrap: anywhere;
  }

  .preview-frame {
    border: 1px solid var(--line-strong);
    border-radius: var(--radius);
    overflow: hidden;
  }

  @media (max-width: 560px) {
    .drop-actions .btn {
      flex: 1 1 100%;
    }
  }
</style>
