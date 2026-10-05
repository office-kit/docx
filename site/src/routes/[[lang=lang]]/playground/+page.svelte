<script lang="ts">
  import { onMount } from 'svelte';
  import type { ValidationIssue } from '@office-kit/docx';
  import DocumentPreview from '$lib/components/DocumentPreview.svelte';
  import RichText from '$lib/components/RichText.svelte';
  import ValidationReport from '$lib/components/ValidationReport.svelte';
  import { downloadDocx } from '$lib/download';
  import { localized } from '$lib/i18n';
  import tools from '$lib/i18n/messages/tools';

  type Stats = ReturnType<typeof import('@office-kit/docx').statistics>;

  const DOCX_ACCEPT =
    '.docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document';
  const SAMPLE_NAME = 'office-kit-demo.docx';

  const m = $derived(localized(tools).playground);

  let status = $state(localized(tools).playground.loading);
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
    status = m.opening(name);
    try {
      const core = await import('@office-kit/docx');
      const doc = core.openDocx(bytes);
      stats = core.statistics(doc);
      docTitle = core.title(doc);
      docAuthor = core.author(doc);
      issues = core.validate(doc);
      savedBytes = core.toUint8Array(doc);
      status = m.opened(
        name,
        bytes.byteLength.toLocaleString(),
        savedBytes.byteLength.toLocaleString(),
      );
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
    status = m.failed(err instanceof Error ? err.message : String(err));
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
          { label: m.stats.paragraphs, value: stats.paragraphs },
          { label: m.stats.headings, value: stats.headings },
          { label: m.stats.tables, value: stats.tables },
          { label: m.stats.images, value: stats.images },
          { label: m.stats.words, value: stats.words.toLocaleString() },
          { label: m.stats.comments, value: stats.comments },
          { label: m.stats.notes, value: stats.footnotes + stats.endnotes },
          { label: m.stats.title, value: docTitle || m.stats.notSet },
          { label: m.stats.author, value: docAuthor || m.stats.notSet },
        ]
      : [],
  );

  onMount(() => void loadSample());
</script>

<svelte:head>
  <title>{m.title} · @office-kit/docx</title>
</svelte:head>

<section class="content">
  <h1>{m.heading}</h1>
  <p class="lede"><RichText text={m.lede} /></p>

  <div
    class="drop"
    class:dropping
    role="group"
    aria-label={m.chooseLabel}
    ondragover={(e) => {
      e.preventDefault();
      dropping = true;
    }}
    ondragleave={() => (dropping = false)}
    ondrop={onDrop}
  >
    <p class="drop-text">{fileName || m.dropHere}</p>
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
        {m.choose}
      </label>
      <button type="button" class="btn" onclick={loadSample} disabled={busy}>
        {m.sample}
      </button>
      {#if savedBytes}
        <button type="button" class="btn" onclick={downloadSaved}>{m.downloadSaved}</button>
      {/if}
    </div>
  </div>

  <p class="caveat">{m.caveat}</p>

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
    <h2>{m.validation}</h2>
    <ValidationReport {issues} />
  {/if}

  <h2>{m.preview}</h2>
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

  .lede :global(code) {
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
