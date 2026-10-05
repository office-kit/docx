/** Qualified element → the command id that creates or edits it (see ../ledger.ts). */
export default {
  // Track Changes: typing, deleting and formatting record these while on.
  "w:trackRevisions": "review.trackChanges",
  "w:ins": "review.trackChanges",
  "w:del": "review.trackedDelete",
  "w:delText": "review.trackedDelete",
  "w:delInstrText": "review.trackedDelete",
  "w:rPrChange": "review.trackChanges",
  "w:pPrChange": "review.trackChanges",
  // Comments.
  "w:comment": "review.addComment",
  "w:comments": "review.addComment",
  "w:commentRangeStart": "review.addComment",
  "w:commentRangeEnd": "review.addComment",
  "w:commentReference": "review.addComment",
  // Proofing language.
  "w:lang": "review.language",
  "w:noProof": "review.language",
  // Protection.
  "w:documentProtection": "review.protect",
  "w:writeProtection": "review.writeProtection",
  "w:permStart": "review.addEditableRange",
  "w:permEnd": "review.addEditableRange",
  "w:locked": "review.lockStyles",
  // View.
  "w:view": "view.setView",
  "w:zoom": "view.setZoom",
  "w:outlineLvl": "outline.setLevel",
  // Document properties.
  "ep:HyperlinkBase": "docprops.setApp",
} satisfies Record<string, string>;
