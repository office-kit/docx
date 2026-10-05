/**
 * Gallery previews (Picture Styles, Corrections, Color) show the selected
 * picture with each option applied, like Word's live galleries.
 */

import { pictureSvg } from "@office-kit/docx-editor";
import type { PictureImage, PictureInfo } from "@office-kit/docx";

// One object URL per image part: dozens of previews share it instead of each
// embedding the image as a data URL. Changing a picture writes a new part, so
// the part name identifies the bytes.
const hrefs = new Map<string, string>();

function previewHref(image: PictureImage): string {
  const cached = hrefs.get(image.partName);
  if (cached) return cached;
  const url = URL.createObjectURL(
    new Blob([image.data.slice().buffer], { type: image.contentType }),
  );
  hrefs.set(image.partName, url);
  return url;
}

let previewCount = 0;

/** The picture drawn `w` × `h` with `patch` applied (trusted SVG markup). */
export function picturePreview(
  picture: PictureInfo,
  patch: Partial<PictureInfo>,
  w: number,
  h: number,
): string {
  const image = picture.image;
  if (!image) return "";
  previewCount += 1;
  return pictureSvg(
    {
      width: w,
      height: h,
      href: previewHref(image),
      picture: { ...picture, ...patch },
      rotation: 0,
      flipH: false,
      flipV: false,
    },
    `wkpv${previewCount}-`,
  );
}
