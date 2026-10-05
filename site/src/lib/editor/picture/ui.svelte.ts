/**
 * Picture editing modes shared by the ribbon (which toggles them) and the
 * canvas overlay (which draws and handles them). One picture is edited at a
 * time, so one set of flags is enough.
 */
export const pictureModes = $state({
  /** Crop: the selection handles become crop handles. */
  cropping: false,
  /** Color ▸ Set Transparent Color: the next click on the picture picks the color. */
  pickingTransparent: false,
  /** Wrap Text ▸ Edit Wrap Points: the wrap polygon's vertices are draggable. */
  editingWrapPoints: false,
});

/** Leave every mode (on Escape, or when the selection moves elsewhere). */
export function resetPictureModes(): void {
  pictureModes.cropping = false;
  pictureModes.pickingTransparent = false;
  pictureModes.editingWrapPoints = false;
}
