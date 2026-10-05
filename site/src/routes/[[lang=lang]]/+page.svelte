<script lang="ts">
  import { base } from '$app/paths';
  import { FamilyGrid, InstallCommand, getProduct } from '@office-kit/site-kit';
  import RichText from '$lib/components/RichText.svelte';
  import { downloadDocx } from '$lib/download';
  import { href, localized } from '$lib/i18n';
  import home from '$lib/i18n/messages/home';
  import type { PageRun } from '$lib/server/page-view';
  import type { PageProps } from './$types';

  const { data }: PageProps = $props();

  const currentProduct = getProduct('docx');

  const m = $derived(localized(home));

  const DOWNLOAD_NAME = 'office-kit-demo.docx';

  let download = $state<{ state: 'idle' | 'working' | 'done' | 'failed'; note: string }>({
    state: 'idle',
    note: '',
  });

  // Runs the exact function shown in the code panel, in the visitor's browser.
  // The library loads on click so the landing page itself ships none of it.
  async function downloadDocument(): Promise<void> {
    download = { state: 'working', note: '' };
    try {
      const [{ toUint8Array }, { buildHeroDocument }] = await Promise.all([
        import('@office-kit/docx'),
        import('$lib/examples/hero-document'),
      ]);
      const bytes = toUint8Array(buildHeroDocument());
      downloadDocx(bytes, DOWNLOAD_NAME);
      download = {
        state: 'done',
        note: m.stage.saved(DOWNLOAD_NAME, (bytes.byteLength / 1024).toFixed(1)),
      };
    } catch (err) {
      console.error(err);
      download = {
        state: 'failed',
        note: m.stage.failed,
      };
    }
  }

</script>

<svelte:head>
  <title>{m.title}</title>
</svelte:head>

{#snippet runs(list: PageRun[])}
  {#each list as run, i (i)}
    <span class:bold={run.bold} class:italic={run.italic}>{run.text}</span>
  {/each}
{/snippet}

<section class="band hero">
  <div class="frame hero-inner">
    <h1>{m.hero.heading}</h1>
    <p class="lede">{m.hero.lede}</p>
    <div class="cta">
      <a href={href('/docs/getting-started')} class="btn primary">{m.hero.start}</a>
      <a href={href('/playground')} class="btn">{m.hero.playground}</a>
      <InstallCommand pkg={currentProduct.pkg} />
    </div>
  </div>
</section>

<section class="band stage" aria-labelledby="stage-title">
  <div class="frame stage-inner">
    <h2 id="stage-title" class="visually-hidden">{m.stage.heading}</h2>
    <div class="stage-grid">
      <figure class="code-pane">
        <figcaption>hero-document.ts</figcaption>
        <!-- A scrollable region must be focusable, or keyboard users cannot scroll it. -->
        <!-- svelte-ignore a11y_no_noninteractive_tabindex -->
        <div class="code-scroll" tabindex="0" role="region" aria-label={m.stage.codeLabel}>
          {@html data.heroCode}
        </div>
      </figure>
      <figure class="sheet-pane">
        <div class="sheet" role="group" aria-label={m.stage.sheetLabel}>
          {#each data.page as block, i (i)}
            {#if block.kind === 'heading'}
              <p
                class="sheet-heading"
                class:bold={block.bold}
                role="heading"
                aria-level={block.level + 2}
                style:--size-pt={block.sizePt}
                style:color={block.color}
              >
                {block.text}
              </p>
            {:else if block.kind === 'paragraph'}
              <p>{@render runs(block.runs)}</p>
            {:else if block.kind === 'list'}
              <svelte:element this={block.ordered ? 'ol' : 'ul'}>
                {#each block.items as item, k (k)}
                  <li>{@render runs(item)}</li>
                {/each}
              </svelte:element>
            {:else}
              <table style:--cell-border={block.borderColor}>
                <tbody>
                  {#each block.rows as row, r (r)}
                    <tr>
                      {#each row as cell, c (c)}
                        <td style:background={cell.fill}>{cell.text}</td>
                      {/each}
                    </tr>
                  {/each}
                </tbody>
              </table>
            {/if}
          {/each}
        </div>
      </figure>
    </div>
    <div class="stage-foot">
      <p><RichText text={m.stage.note} /></p>
      <div class="stage-action">
        <button
          type="button"
          class="btn stage-btn"
          onclick={downloadDocument}
          disabled={download.state === 'working'}
        >
          {download.state === 'working' ? m.stage.building : m.stage.download}
        </button>
        <p class="stage-note" aria-live="polite">{download.note}</p>
      </div>
    </div>
  </div>
</section>

<section class="band">
  <div class="frame">
    <ul class="paths">
      {#each m.paths as p (p.code)}
        <li>
          <code class="path-code">{p.code}</code>
          <h2>{p.title}</h2>
          <p>{p.body}</p>
          <a href={href(p.href)}>{p.link}</a>
        </li>
      {/each}
    </ul>
  </div>
</section>

<section class="band template">
  <div class="frame template-inner">
    <div class="template-text">
      <h2>{m.template.heading}</h2>
      {#each m.template.body as text, i (i)}
        <p><RichText {text} /></p>
      {/each}
    </div>
    <figure class="template-code">
      <figcaption>template-fill.ts</figcaption>
      {@html data.templateCode}
    </figure>
  </div>
</section>

<section class="band proof" aria-labelledby="proof-title">
  <div class="frame proof-inner">
    <h2 id="proof-title">{m.proof.heading}</h2>
    <p class="proof-lede">{m.proof.lede}</p>
    <dl>
      {#each m.proof.items as item (item.claim)}
        <div class="proof-row">
          <dt>{item.claim}</dt>
          <dd>{item.how}</dd>
        </div>
      {/each}
    </dl>
  </div>
</section>

<section class="band" aria-labelledby="compare-title">
  <div class="frame compare">
    <div class="section-head">
      <h2 id="compare-title">{m.compare.heading}</h2>
      <p><RichText text={m.compare.lede} /></p>
    </div>
    <div class="table-scroll">
      <table>
        <thead>
          <tr>
            <th scope="col"><span class="visually-hidden">{m.compare.capability}</span></th>
            <th scope="col">@office-kit/docx</th>
            <th scope="col">docx (npm)</th>
          </tr>
        </thead>
        <tbody>
          {#each m.compare.rows as row (row.topic)}
            <tr>
              <th scope="row">{row.topic}</th>
              <td data-tone={row.ours.tone}>{row.ours.text}</td>
              <td data-tone={row.theirs.tone}>{row.theirs.text}</td>
            </tr>
          {/each}
        </tbody>
      </table>
    </div>
    <p class="compare-advice">{m.compare.advice}</p>
  </div>
</section>

<section class="band" aria-labelledby="caps-title">
  <div class="frame caps">
    <div class="section-head">
      <h2 id="caps-title">{m.caps.heading}</h2>
      <p><RichText text={m.caps.lede} /></p>
    </div>
    <dl class="caps-grid">
      {#each m.caps.items as c (c.area)}
        <div>
          <dt>{c.area}</dt>
          <dd>{c.items}</dd>
        </div>
      {/each}
    </dl>
    <div class="not-yet">
      <h3>{m.caps.notYet}</h3>
      <ul>
        {#each m.caps.notYetItems as item (item)}
          <li>{item}</li>
        {/each}
      </ul>
    </div>
  </div>
</section>

<section class="band agents">
  <div class="frame agents-inner">
    <div>
      <h2>{m.agents.heading}</h2>
      <p>{m.agents.body}</p>
    </div>
    <ul>
      <li>
        <a href="{base}/llms.txt"><code>/llms.txt</code></a>
        <span>{m.agents.index}</span>
      </li>
      <li>
        <a href="{base}/llms-full.txt"><code>/llms-full.txt</code></a>
        <span>{m.agents.full}</span>
      </li>
    </ul>
  </div>
</section>

<section class="band" aria-labelledby="family-title">
  <div class="frame">
    <div class="section-head family-head">
      <h2 id="family-title">{m.family.heading}</h2>
      <p>{m.family.body}</p>
    </div>
    <FamilyGrid product="docx" />
  </div>
</section>

<style>
  /* Hero. */
  .hero-inner {
    padding: clamp(3rem, 6vw, 4.75rem) var(--gutter) clamp(2.75rem, 5vw, 3.75rem);
    text-align: center;
  }

  h1 {
    max-width: 21ch;
    margin: 0 auto;
    font-size: clamp(2.35rem, 6vw, 4.25rem);
    font-weight: 650;
    line-height: 1.02;
    letter-spacing: -0.04em;
  }

  .lede {
    max-width: 62ch;
    margin: 1.75rem auto 0;
    color: var(--ink-2);
    font-size: clamp(1.05rem, 1.6vw, 1.2rem);
    line-height: 1.55;
    text-wrap: pretty;
  }

  .cta {
    display: flex;
    flex-wrap: wrap;
    justify-content: center;
    gap: 0.6rem;
    margin-top: 2.25rem;
  }

  /* Stage: the product accent as a field, ruled like a sheet of lined paper. */
  .stage {
    background-color: var(--accent);
    background-image: linear-gradient(rgb(255 255 255 / 0.13) 1px, transparent 1px);
    background-size: 100% 28px;
  }

  .stage-inner {
    border-inline-color: rgb(255 255 255 / 0.22);
    padding: clamp(1.5rem, 4vw, 3.5rem) var(--gutter) 0;
  }

  .stage-grid {
    display: grid;
    grid-template-columns: minmax(0, 1fr) minmax(0, 1fr);
    gap: clamp(1rem, 2vw, 1.75rem);
    align-items: stretch;
  }

  figure {
    margin: 0;
  }

  .code-pane {
    display: flex;
    flex-direction: column;
    min-height: 0;
    border-radius: var(--radius);
    background: var(--night);
    box-shadow: var(--shadow-pop);
    overflow: hidden;
  }

  .code-pane figcaption,
  .template-code figcaption {
    flex: none;
    padding: 0.6rem 1rem;
    border-bottom: 1px solid var(--night-line);
    color: var(--night-ink-2);
    font-family: var(--mono);
    font-size: 0.8rem;
  }

  /* The pane's height comes from the page beside it; the code scrolls inside. */
  .code-scroll {
    position: relative;
    flex: 1;
    min-height: 0;
    overflow: auto;
  }

  .code-scroll :global(pre) {
    position: absolute;
    inset: 0 auto auto 0;
    min-width: 100%;
    margin: 0;
    border: none;
    border-radius: 0;
    overflow: visible;
    font-size: 0.8rem;
  }

  .sheet-pane {
    container-type: inline-size;
  }

  /* The sheet is the document's paper, not this site's: white in both themes,
   * and set in the serif Word falls back to when a file names no font. One
   * document point is `--pt` CSS pixels, so the sizes stay in the proportions
   * the file specifies while the sheet scales with its column. */
  .sheet {
    --pt: clamp(1.2px, 0.27cqw, 1.5px);
    /* ECMA-376 §17.3.2.38: a run with no w:sz is 10 points. */
    --default-size-pt: 10;
    height: 100%;
    padding: clamp(1.5rem, 9cqw, 3.5rem) clamp(1.25rem, 9cqw, 3.5rem);
    border-radius: 3px;
    background: #fff;
    box-shadow: var(--shadow-pop);
    color: #1b1b1b;
    font-family: 'Times New Roman', Times, 'Liberation Serif', serif;
    font-size: calc(var(--default-size-pt) * var(--pt));
    line-height: 1.35;
  }

  .sheet p,
  .sheet ul,
  .sheet ol {
    margin: 0 0 0.55em;
  }

  .sheet .sheet-heading {
    margin: 0.7em 0 0.3em;
    font-size: calc(var(--size-pt, var(--default-size-pt)) * var(--pt));
    line-height: 1.15;
  }

  .sheet .sheet-heading:first-child {
    margin-top: 0;
  }

  .sheet .bold {
    font-weight: 700;
  }

  .sheet .italic {
    font-style: italic;
  }

  .sheet ul,
  .sheet ol {
    padding-left: 2.2em;
  }

  .sheet li {
    margin: 0 0 0.2em;
  }

  .sheet table {
    width: 100%;
    margin: 0.4em 0 0;
    border-collapse: collapse;
    table-layout: fixed;
    font-size: inherit;
  }

  .sheet td {
    padding: 0.3em 0.5em;
    border: 1px solid var(--cell-border, transparent);
    color: inherit;
    font-size: inherit;
  }

  .stage-foot {
    display: flex;
    align-items: flex-start;
    justify-content: space-between;
    gap: 1.5rem 3rem;
    padding: clamp(1.25rem, 3vw, 2rem) 0 clamp(1.5rem, 3vw, 2.5rem);
    color: var(--on-accent);
  }

  .stage-foot p {
    max-width: 62ch;
    margin: 0;
    font-size: 1rem;
    line-height: 1.55;
  }

  .stage-foot :global(code) {
    white-space: nowrap;
    background: rgb(255 255 255 / 0.16);
    border-color: rgb(255 255 255 / 0.3);
    color: inherit;
  }

  .stage-action {
    flex: none;
    display: flex;
    flex-direction: column;
    align-items: flex-end;
    gap: 0.5rem;
    max-width: 22rem;
  }

  .stage-btn {
    background: #fff;
    border-color: #fff;
    color: #15171c;
  }

  .stage-btn:hover {
    background: #15171c;
    border-color: #15171c;
    color: #fff;
  }

  .stage-btn:disabled {
    cursor: progress;
    opacity: 0.8;
  }

  .stage-foot .stage-note {
    font-size: 0.88rem;
    text-align: right;
  }

  .stage :global(:focus-visible) {
    outline-color: #fff;
  }

  /* Three ways in. */
  .paths {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    list-style: none;
    margin: 0;
    padding: 0;
  }

  .paths li {
    display: flex;
    flex-direction: column;
    margin: 0;
    padding: 2.5rem var(--gutter) 2.25rem;
    border-right: 1px solid var(--line);
  }

  .paths li:last-child {
    border-right: none;
  }

  .path-code {
    align-self: flex-start;
    background: var(--accent-wash);
    border-color: transparent;
    color: var(--accent-ink);
    font-size: 0.82rem;
  }

  .paths h2 {
    margin: 1.1rem 0 0.6rem;
    font-size: 1.35rem;
  }

  .paths p {
    margin: 0 0 1.25rem;
    color: var(--ink-2);
    font-size: 0.97rem;
  }

  .paths a {
    margin-top: auto;
    font-weight: 550;
  }

  /* Template section. */
  .template-inner {
    display: grid;
    grid-template-columns: minmax(0, 1fr) minmax(0, 1fr);
    gap: clamp(2rem, 5vw, 4.5rem);
    align-items: center;
    padding: clamp(3rem, 7vw, 5.5rem) var(--gutter);
  }

  .template-text h2 {
    margin: 0 0 1.25rem;
    font-size: clamp(1.7rem, 3.4vw, 2.5rem);
    letter-spacing: -0.03em;
    line-height: 1.08;
  }

  .template-text p {
    color: var(--ink-2);
    max-width: 52ch;
  }

  .template-code {
    border-radius: var(--radius);
    background: var(--night);
    border: 1px solid var(--night-line);
    overflow: hidden;
  }

  .template-code :global(pre) {
    margin: 0;
    border: none;
    border-radius: 0;
  }

  /* Proof: the one dark band on the page. */
  .proof {
    background: var(--night);
    color: var(--night-ink);
    border-bottom-color: var(--night-line);
  }

  .proof-inner {
    border-inline-color: var(--night-line);
    padding: clamp(3rem, 7vw, 5.5rem) var(--gutter);
  }

  .proof h2 {
    margin: 0;
    color: #fff;
    font-size: clamp(1.9rem, 4.2vw, 3.1rem);
    letter-spacing: -0.035em;
    line-height: 1.05;
  }

  .proof-lede {
    max-width: 58ch;
    margin: 1.1rem 0 2.75rem;
    color: var(--night-ink-2);
    font-size: 1.08rem;
  }

  .proof dl {
    margin: 0;
    border-top: 1px solid var(--night-line);
  }

  .proof-row {
    display: grid;
    grid-template-columns: minmax(0, 5fr) minmax(0, 6fr);
    gap: 0.5rem 3rem;
    padding: 1.5rem 0;
    border-bottom: 1px solid var(--night-line);
  }

  .proof dt {
    font-family: var(--display);
    font-size: 1.25rem;
    font-weight: 550;
    line-height: 1.25;
    letter-spacing: -0.015em;
    color: #fff;
    text-wrap: balance;
  }

  .proof dd {
    margin: 0;
    color: var(--night-ink-2);
    font-size: 0.98rem;
  }

  /* Section heads shared by the lower sections. */
  .section-head {
    max-width: 66ch;
    padding: clamp(3rem, 7vw, 5rem) var(--gutter) 0;
  }

  .section-head h2 {
    margin: 0 0 1rem;
    font-size: clamp(1.7rem, 3.4vw, 2.5rem);
    letter-spacing: -0.03em;
    line-height: 1.08;
  }

  .section-head p {
    margin: 0;
    color: var(--ink-2);
    font-size: 1.04rem;
  }

  /* Comparison. */
  .compare {
    padding-bottom: clamp(3rem, 6vw, 4.5rem);
  }

  .compare .table-scroll {
    margin-top: 2.25rem;
    padding: 0 var(--gutter);
  }

  .compare table {
    min-width: 640px;
    margin: 0;
    table-layout: fixed;
  }

  .compare thead th {
    width: 37%;
    padding-bottom: 0.8rem;
    font-family: var(--mono);
    font-size: 0.85rem;
    font-weight: 550;
    color: var(--ink);
  }

  .compare thead th:first-child {
    width: 26%;
  }

  .compare thead th:nth-child(2) {
    color: var(--accent-ink);
  }

  .compare tbody th {
    padding-right: 1.5rem;
    border-bottom-color: var(--line);
    font-weight: 600;
    color: var(--ink);
    white-space: normal;
  }

  .compare td {
    padding-right: 1.5rem;
    color: var(--ink-2);
  }

  /* A mark in front of the text carries yes / no, so colour is never the only cue. */
  .compare td[data-tone]::before {
    display: inline-block;
    width: 1.25em;
    font-weight: 700;
  }

  .compare td[data-tone='yes']::before {
    content: '✓';
    color: #168a4f;
  }

  .compare td[data-tone='no']::before {
    content: '✕';
    color: var(--ink-3);
  }

  .compare td[data-tone='part']::before {
    content: '◐';
    color: var(--ink-3);
  }

  .compare-advice {
    max-width: 70ch;
    margin: 2rem 0 0;
    padding: 0 var(--gutter);
    color: var(--ink);
    font-size: 1.02rem;
  }

  /* Capabilities. */
  .caps-grid {
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    margin: 2.5rem 0 0;
    border-top: 1px solid var(--line);
  }

  .caps-grid > div {
    padding: 1.5rem var(--gutter) 1.6rem;
    border-right: 1px solid var(--line);
    border-bottom: 1px solid var(--line);
  }

  .caps-grid > div:nth-child(4n) {
    border-right: none;
  }

  .caps-grid dt {
    font-family: var(--display);
    font-size: 1.12rem;
    font-weight: 600;
    letter-spacing: -0.01em;
  }

  .caps-grid dd {
    margin: 0.45rem 0 0;
    color: var(--ink-2);
    font-size: 0.93rem;
    line-height: 1.5;
  }

  .not-yet {
    display: grid;
    grid-template-columns: minmax(0, 1fr) minmax(0, 3fr);
    gap: 0.75rem 2rem;
    padding: 1.75rem var(--gutter) 2.25rem;
    background: var(--wash);
  }

  .not-yet h3 {
    margin: 0;
    font-size: 1.12rem;
  }

  .not-yet ul {
    margin: 0;
    padding: 0;
    list-style: none;
    columns: 2;
    column-gap: 2.5rem;
    color: var(--ink-2);
    font-size: 0.93rem;
  }

  .not-yet li {
    margin: 0 0 0.45rem;
    break-inside: avoid;
  }

  /* Agents. */
  .agents-inner {
    display: grid;
    grid-template-columns: minmax(0, 5fr) minmax(0, 6fr);
    gap: clamp(2rem, 5vw, 4.5rem);
    align-items: center;
    padding: clamp(3rem, 7vw, 5rem) var(--gutter);
  }

  .agents h2 {
    margin: 0 0 1rem;
    font-size: clamp(1.7rem, 3.4vw, 2.5rem);
    letter-spacing: -0.03em;
    line-height: 1.08;
  }

  .agents p {
    color: var(--ink-2);
    max-width: 50ch;
  }

  .agents ul {
    list-style: none;
    margin: 0;
    padding: 0;
    border-top: 1px solid var(--line);
  }

  .agents li {
    display: grid;
    grid-template-columns: 9rem minmax(0, 1fr);
    gap: 0.25rem 1rem;
    align-items: baseline;
    margin: 0;
    padding: 0.95rem 0;
    border-bottom: 1px solid var(--line);
  }

  .agents li span {
    color: var(--ink-2);
    font-size: 0.95rem;
  }

  .family-head {
    padding-bottom: 2.5rem;
    border-bottom: 1px solid var(--line);
    max-width: none;
  }

  .family-head p {
    max-width: 66ch;
  }

  @media (max-width: 960px) {
    .stage-grid {
      grid-template-columns: 1fr;
    }

    /* Page first on narrow screens: it is the payoff, the code is the proof. */
    .sheet-pane {
      order: -1;
    }

    .code-scroll {
      height: 300px;
      flex: none;
    }

    .paths {
      grid-template-columns: 1fr;
    }

    .paths li {
      border-right: none;
      border-bottom: 1px solid var(--line);
      padding-block: 2rem;
    }

    .paths li:last-child {
      border-bottom: none;
    }

    .template-inner,
    .agents-inner {
      grid-template-columns: 1fr;
    }

    .caps-grid {
      grid-template-columns: repeat(2, 1fr);
    }

    .caps-grid > div:nth-child(2n) {
      border-right: none;
    }
  }

  @media (max-width: 720px) {
    .stage-foot {
      flex-direction: column;
    }

    .stage-action {
      align-items: flex-start;
      max-width: none;
    }

    .stage-foot .stage-note {
      text-align: left;
    }

    .proof-row {
      grid-template-columns: 1fr;
    }

    .not-yet {
      grid-template-columns: 1fr;
    }

    .not-yet ul {
      columns: 1;
    }
  }

  @media (max-width: 520px) {
    .cta .btn {
      flex: 1 1 100%;
    }

    .caps-grid {
      grid-template-columns: 1fr;
    }

    .caps-grid > div {
      border-right: none;
    }

    .agents li {
      grid-template-columns: 1fr;
    }
  }
</style>
