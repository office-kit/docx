<script lang="ts">
  /** One Table Styles gallery thumbnail: a 5×5 table drawn with the style's resolved formatting. */
  import type { TablePreviewCell } from '@office-kit/docx-editor';

  type Props = { name: string; cells: readonly (readonly TablePreviewCell[])[] | undefined; selected: boolean; onclick: () => void };
  const { name, cells, selected, onclick }: Props = $props();

  function css(c: TablePreviewCell): string {
    return [
      c.background ? `background:${c.background}` : '',
      `border-top:${c.top}`,
      `border-right:${c.right}`,
      `border-bottom:${c.bottom}`,
      `border-left:${c.left}`,
      `--ink:${c.text}`,
      c.bold ? '--ink-h:2px' : '',
    ]
      .filter(Boolean)
      .join(';');
  }
</script>

<button class="tbl-tile" class:selected {onclick} title={name} aria-label={name} role="menuitemradio" aria-checked={selected}>
  <span class="tbl-mini">
    {#each cells ?? [] as row, r (r)}
      {#each row as cell, c (c)}<i style={css(cell)}><b></b></i>{/each}
    {/each}
  </span>
</button>

<style>
  .tbl-tile {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 62px;
    height: 46px;
    padding: 3px;
    border: 1px solid transparent;
    border-radius: 3px;
    background: #fff;
    cursor: pointer;
  }
  .tbl-tile:hover {
    border-color: #c8c8c8;
  }
  .tbl-tile.selected {
    border-color: #a9bcd9;
    background: #ebf0f6;
    box-shadow: 0 0 0 1px #a9bcd9;
  }
  .tbl-mini {
    display: grid;
    grid-template-columns: repeat(5, 10px);
    grid-auto-rows: 7px;
    border-collapse: collapse;
  }
  .tbl-mini i {
    display: flex;
    align-items: center;
    justify-content: center;
    box-sizing: border-box;
    border-width: 0;
  }
  /* A stroke standing for the cell's text, in its color and weight. */
  .tbl-mini b {
    width: 5px;
    height: var(--ink-h, 1px);
    background: var(--ink, #000);
    opacity: 0.7;
  }
</style>
