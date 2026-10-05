/** Qualified element → the command id that creates or edits it (see ../ledger.ts). */
export default {
  // Phonetic Guide.
  "w:ruby": "text.ruby",
  "w:rubyPr": "text.ruby",
  "w:rubyAlign": "text.ruby",
  "w:hps": "text.ruby",
  "w:hpsRaise": "text.ruby",
  "w:hpsBaseText": "text.ruby",
  "w:lid": "text.ruby",
  "w:rt": "text.ruby",
  "w:rubyBase": "text.ruby",
  // Character Border, Asian Layout.
  "w:bdr": "text.bdr",
  "w:eastAsianLayout": "text.eastAsianLayout",
  "w:fitText": "text.fitText",
  // Paragraph borders between paragraphs, tab stops.
  "w:between": "paragraph.bdr",
  "w:bar": "paragraph.bdr",
  "w:tabs": "paragraph.tabs",
  // List libraries, Restart / Continue Numbering.
  "w:abstractNumId": "list.applyPreset",
  "w:multiLevelType": "list.applyPreset",
  "w:lvlOverride": "list.restart",
  "w:startOverride": "list.restart",
} satisfies Record<string, string>;
