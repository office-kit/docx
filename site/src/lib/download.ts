const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

/** Hands `bytes` to the visitor as a file download, without a round trip to any server. */
export function downloadDocx(bytes: Uint8Array, fileName: string): void {
  // `slice()` yields an ArrayBuffer-backed copy, which is what BlobPart accepts.
  const url = URL.createObjectURL(new Blob([bytes.slice()], { type: DOCX_MIME }));
  const a = document.createElement("a");
  a.href = url;
  a.download = fileName;
  a.click();
  URL.revokeObjectURL(url);
}
