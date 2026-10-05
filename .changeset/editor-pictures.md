---
"@office-kit/docx": minor
"@office-kit/docx-editor": minor
---

Pictures, charts and object arrangement.

`@office-kit/docx` adds a picture API (`src/api/picture.ts`):

- `readDrawing` reads a `<w:drawing>`: wrap, anchor position, crop, shape, border, effects, color adjustments and image.
- Arrange: `setDrawingWrap`, `setDrawingPosition`, `setDrawingAnchorOptions`, `arrangeDrawing` (Bring Forward / Send Backward and their variants), `setDrawingTransform`, `setDrawingName`, `setDrawingHidden`, `setDrawingAspectLock`, `setDrawingHyperlink` and `removeDrawing`.
- Picture formatting: `setPictureCrop`, `setPictureGeometry`, `setPictureOutline`, `setPictureEffects`, `setPictureColorAdjustments`, `changePicture` and `resetPicture`.
- Charts: `addChartRun`, `readChart` and `setChart` write a real chart part with an embedded workbook.
- `imagePixelSize` and `imageNaturalSizeEmu` read image dimensions.

Everything is written as ECMA-376 markup, with no Office extension namespaces.

Fixed: `setImageSizeEmu` no longer writes into an `a:ext` that sits inside an extension list.

`@office-kit/docx-editor` adds:

- Rendering for floating and inline pictures and charts: wrapping, z-order, rotation, crop, effects and blip adjustments.
- `layoutFloatingObjects` and `floatFrameStart` for placing floating objects in the browser.
- Picture, arrange and chart commands, including `drawing.layout` for the Layout Options dialog and `chart.insert` / `chart.edit`.
