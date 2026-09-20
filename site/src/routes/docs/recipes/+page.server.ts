import { examples, type Example, type ExampleKey } from "$lib/examples";
import { highlight } from "$lib/server/highlight";
import type { PageServerLoad } from "./$types";

const KEYS: ExampleKey[] = [
  "recipeMailMerge",
  "recipeStyledBase",
  "recipeTrackedChanges",
  "previewEmbed",
];

export const load: PageServerLoad = async () => {
  const recipes = await Promise.all(
    KEYS.map(async (key) => {
      const ex: Example = examples[key];
      return {
        key,
        title: ex.title,
        description: ex.description,
        path: ex.path,
        seeAlso: ex.seeAlso,
        html: await highlight(ex.source, "ts"),
      };
    }),
  );
  return { recipes };
};
