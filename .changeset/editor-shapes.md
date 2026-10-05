---
"@office-kit/docx": minor
"@office-kit/docx-editor": minor
---

Shapes, text boxes, WordArt, ink and SmartArt.

- `@office-kit/docx`: build and edit VML shapes (`addShape` with about 120 Word presets, `addWordArt`, ink strokes, `addDrawingCanvas`, groups), with fill (solid, gradient, pattern, picture), outline, shadow, layout, wrapping, z-order, alt text, linked text boxes and vertical text. Insert and edit SmartArt graphics (`addSmartArt`, `setSmartArtNodes`, layouts, colours, styles, size, `removeSmartArt`) as ECMA-376 DrawingML diagram parts. Existing VML from Word documents is read, including shape types.
- `@office-kit/docx-editor`: renders VML shapes, text boxes, WordArt, ink and SmartArt on the canvas (`renderPictHtml` is exported for watermarks), with commands for every Shape Format, SmartArt Design and Draw tab action, editable text boxes, and Ink to Shape recognition (`recognizeInkShape`).
