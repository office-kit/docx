/**
 * Character scale (`w:w`, §17.3.2.43; Word's 文字の拡大/縮小) on the canvas.
 * The renderer draws a scaled run with a horizontal transform, which leaves
 * its layout width unscaled; this sets the margin that makes the run advance
 * by its scaled width. Browser-only; run it before tabs are laid out.
 */

const FULL_SCALE = 100;

export function layoutScaledText(root: ParentNode): void {
  const runs = Array.from(root.querySelectorAll<HTMLElement>("[data-wk-scale]"));
  for (const run of runs) run.style.marginRight = "";
  // Read every width before writing any margin: one layout pass, not one per run.
  const widths = runs.map((run) => run.offsetWidth);
  runs.forEach((run, i) => {
    const scale = Number(run.dataset.wkScale) / FULL_SCALE;
    run.style.marginRight = `${(widths[i] ?? 0) * (scale - 1)}px`;
  });
}
