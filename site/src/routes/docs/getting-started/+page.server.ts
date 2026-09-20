import { examples } from "$lib/examples";
import { highlight } from "$lib/server/highlight";
import type { PageServerLoad } from "./$types";

export const load: PageServerLoad = async () => {
  const [fromScratch, templateFill] = await Promise.all([
    highlight(examples.fromScratch.source, "ts"),
    highlight(examples.templateFill.source, "ts"),
  ]);
  return {
    fromScratch: { path: examples.fromScratch.path, html: fromScratch },
    templateFill: { path: examples.templateFill.path, html: templateFill },
  };
};
