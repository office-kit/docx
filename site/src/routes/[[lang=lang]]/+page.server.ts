import { openDocx, toUint8Array } from "@office-kit/docx";
import { buildHeroDocument } from "$lib/examples/hero-document";
import heroDocumentSource from "$lib/examples/hero-document.ts?raw";
import templateFillSource from "$lib/examples/template-fill.ts?raw";
import { highlight } from "$lib/server/highlight";
import { readPage } from "$lib/server/page-view";
import type { PageServerLoad } from "./$types";

// Each example opens with a comment addressed to maintainers; visitors only
// need the code, so a snippet starts where the code does.
const from = (source: string, marker: string): string => source.slice(source.indexOf(marker));

export const load: PageServerLoad = async () => {
  // The page is read back from the saved bytes, not from the in-memory model,
  // so what the visitor sees is what is in the .docx. Doing it at prerender
  // time puts the page in the static HTML: it needs no client JS and cannot
  // drift from the code beside it.
  const saved = toUint8Array(buildHeroDocument());
  const page = readPage(openDocx(saved));

  const [heroCode, templateCode] = await Promise.all([
    highlight(from(heroDocumentSource, "export function"), "ts"),
    highlight(from(templateFillSource, "import "), "ts"),
  ]);

  return { page, heroCode, templateCode };
};
