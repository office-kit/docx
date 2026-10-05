/** Qualified element → the command id that creates or edits it (see ../ledger.ts). */
export default {
  // Shapes, text boxes, WordArt and ink are VML inside w:pict (commands/shapes.ts).
  "w:pict": "shape.insert",
  "v:shape": "shape.insert",
  "v:rect": "shape.change",
  "v:roundrect": "shape.change",
  "v:oval": "shape.change",
  "v:group": "shape.group",
  "v:fill": "shape.fill",
  "v:stroke": "shape.outline",
  "v:shadow": "shape.shadow",
  "v:textbox": "shape.textLayout",
  "w:txbxContent": "shape.text",
  "v:textpath": "shape.wordArt",
  "v:path": "shape.insertWordArt",
  "wvml:wrap": "shape.wrap",
} satisfies Record<string, string>;
