<script lang="ts" module>
  // Line icons drawn on a 20×20 grid in the style of Word's ribbon (thin
  // outlines, a colored accent only where Word has one). Drawn as SVG rather
  // than font glyphs so every platform shows the same shape.
  export const ICONS = {
    undo: 'M7 5 3.5 8.5 7 12M4 8.5h8a4.5 4.5 0 0 1 0 9H9',
    redo: 'M13 5l3.5 3.5L13 12M16 8.5H8a4.5 4.5 0 0 0 0 9h3',
    save: 'M4 3.5h10l2.5 2.5v10.5H3.5V3.5zM6.5 3.5v4h6v-4M6 16.5v-5h8v5',
    newDoc: 'M5 2.5h7l3.5 3.5v11.5H5zM12 2.5V6h3.5',
    open: 'M2.5 5.5v10h13l2-7H6l-2 7M2.5 5.5V4h5l1.5 1.5h6v3',
    search: 'M8.5 3a5.5 5.5 0 1 1 0 11 5.5 5.5 0 0 1 0-11zM12.5 12.5 17 17',
    comment: 'M3 4h14v9.5H9l-4 3v-3H3z',
    share: 'M11 3.5h5.5V9M16.5 3.5 9.5 10.5M14 12v4.5H3.5V6H8',
    paste: 'M7 3.5H4.5v13H9M7 3.5V2.5h6v1M13 3.5h2.5v4M7 3.5v2h6v-2M10 8.5h7.5v9H10z',
    cut: 'M6 13.5a2.25 2.25 0 1 0 0 .01M14 13.5a2.25 2.25 0 1 0 0 .01M7.5 12 14 3M12.5 12 6 3',
    copy: 'M7 6.5h9v11H7zM4 13.5V3h9',
    alignLeft: 'M3 4.5h14M3 8.5h9M3 12.5h14M3 16.5h9',
    alignCenter: 'M3 4.5h14M5.5 8.5h9M3 12.5h14M5.5 16.5h9',
    alignRight: 'M3 4.5h14M8 8.5h9M3 12.5h14M8 16.5h9',
    alignJustify: 'M3 4.5h14M3 8.5h14M3 12.5h14M3 16.5h14',
    bullets: 'M8 5h9.5M8 10h9.5M8 15h9.5M3.5 5h.5M3.5 10h.5M3.5 15h.5',
    numbering: 'M8 5h9.5M8 10h9.5M8 15h9.5M3 3.5l1-.5v4M3 8.5h2l-2 3h2M3 13.5h2l-1 1 1 1.5H3',
    indentMore: 'M9 4.5h8M9 8.5h8M9 12.5h8M3 16.5h14M3 6.5l3 2.5-3 2.5',
    indentLess: 'M9 4.5h8M9 8.5h8M9 12.5h8M3 16.5h14M6 6.5l-3 2.5 3 2.5',
    lineSpacing: 'M9 4.5h8.5M9 8.5h8.5M9 12.5h8.5M9 16.5h8.5M4.5 3v14M2.5 5l2-2 2 2M2.5 15l2 2 2-2',
    pilcrow: 'M10 3.5v13M13.5 3.5v13M15.5 3.5H8.5a3.5 3.5 0 0 0 0 7H10',
    table: 'M3 3.5h14v13H3zM3 8h14M3 12.5h14M8 3.5v13M12.5 3.5v13',
    picture: 'M3 4h14v12H3zM3 13l4-4 3.5 3.5L13 10l4 4M13.5 7.5a.5.5 0 1 0 0 .01',
    link: 'M8.5 11.5l3-3M7 9 5 11a2.5 2.5 0 0 0 3.5 3.5l2-2M13 11l2-2a2.5 2.5 0 0 0-3.5-3.5l-2 2',
    pageBreak: 'M5 2.5v5h10v-5M5 17.5v-5h10v5M2.5 10h2M7 10h2M11 10h2M15.5 10h2',
    sectionBreak: 'M5 2.5v5h10v-5M5 17.5v-5h10v5M2.5 10h15',
    pageNumber: 'M5 2.5h7l3.5 3.5v11.5H5zM8.5 14h3M10 12v4',
    orientation: 'M3 6h9v11H3zM9 3h8v6',
    size: 'M5 2.5h10v15H5zM8 6h4',
    toc: 'M3 4h5M3 8h14M3 12h14M3 16h14M14 4h3',
    footnote: 'M3 4h14M3 8h14M3 12h8M3 16.5h6M14 14.5v-3l-1 .5',
    endnote: 'M3 4h14M3 8h14M3 12h14M3 16.5h6M14 16.5h3',
    bookmark: 'M5.5 2.5h9v15L10 13.5l-4.5 4z',
    accept: 'M3 10.5 7.5 15 17 5',
    reject: 'M4.5 4.5l11 11M15.5 4.5l-11 11',
    zoomIn: 'M8.5 3a5.5 5.5 0 1 1 0 11 5.5 5.5 0 0 1 0-11zM12.5 12.5 17 17M6 8.5h5M8.5 6v5',
    zoomOut: 'M8.5 3a5.5 5.5 0 1 1 0 11 5.5 5.5 0 0 1 0-11zM12.5 12.5 17 17M6 8.5h5',
    onePage: 'M5.5 2.5h9v15h-9z',
    xml: 'M7 5 2.5 10 7 15M13 5l4.5 5-4.5 5M11.5 3.5l-3 13',
    language: 'M3 4h8M7 2.5V4M4.5 4c.5 3 2.5 5.5 5.5 6.5M9.5 4c-.5 3-2.5 5.5-5.5 6.5M10.5 17.5l3-8 3 8M11.5 15h4',
    chevronDown: 'M5 8l5 5 5-5',
    chevronRight: 'M8 5l5 5-5 5',
    close: 'M5 5l10 10M15 5 5 15',
  } as const;

  export type IconName = keyof typeof ICONS;
</script>

<script lang="ts">
  type Props = { name: IconName; size?: number };
  const { name, size = 20 }: Props = $props();
</script>

<svg
  viewBox="0 0 20 20"
  width={size}
  height={size}
  fill="none"
  stroke="currentColor"
  stroke-width="1.2"
  stroke-linecap="round"
  stroke-linejoin="round"
  aria-hidden="true"><path d={ICONS[name]} /></svg>
