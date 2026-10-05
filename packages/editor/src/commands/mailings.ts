/**
 * Mailings commands: envelopes, labels and mail merge. The recipient list
 * itself is session state (Word keeps it in the external data source), so
 * commands that need it take it as a parameter. Merging and creating a label
 * document produce a *new* document rather than editing this one — the UI
 * calls `mergeToNewDocument` / `createLabelDocument` directly for those.
 */

import {
  type AddressBlockOptions,
  addEnvelope,
  attachRecipientList,
  type EnvelopeOptions,
  type GreetingLineOptions,
  insertAddressBlock,
  insertGreetingLine,
  insertMergeField,
  insertMergeRule,
  type MailMergeDocumentType,
  type MergeRule,
  type PreviewOptions,
  previewMailMerge,
  type RecipientList,
  removeEnvelope,
  setMailMergeDocumentType,
  setMailMergeFieldMap,
  setRecipientInclusion,
  updateLabels,
  type WmlParagraph,
} from "@office-kit/docx";
import { absoluteOffset } from "../char-offset.js";
import { paragraphAt } from "../doc-access.js";
import type { EditorModel } from "../model.js";
import { orderSelection } from "../selection.js";
import type { Command } from "./types.js";

/** The caret paragraph and character offset (the end of a selected range). */
function caretPoint(model: EditorModel): { paragraph: WmlParagraph; offset: number } {
  const sel = model.selection;
  const at = sel ? orderSelection(sel).end : undefined;
  const paragraph = at ? paragraphAt(model.doc, at) : undefined;
  if (!at || !paragraph) throw new Error("Place the insertion point in text first.");
  return { paragraph, offset: absoluteOffset(paragraph, at) };
}

const hasCaret = (model: EditorModel): boolean => {
  const focus = model.selection?.focus;
  return !!focus && !!paragraphAt(model.doc, focus);
};

/** Start Mail Merge ▸ Letters / E-mail Messages / Envelopes / Labels / Directory / Normal. */
export const startMailMergeCommand: Command<{ type: MailMergeDocumentType | "normal" }> = {
  id: "mailings.start",
  group: "mailings",
  label: "Start Mail Merge",
  run(model, { type }) {
    setMailMergeDocumentType(model.doc, type);
  },
};

/** Select Recipients ▸ Use an Existing List / Type a New List. */
export const selectRecipientsCommand: Command<{
  path: string;
  list: RecipientList;
  fieldMap?: Readonly<Record<string, string>>;
}> = {
  id: "mailings.selectRecipients",
  group: "mailings",
  label: "Select Recipients",
  run(model, { path, list, fieldMap }) {
    attachRecipientList(model.doc, path, list, fieldMap);
  },
};

/** Match Fields ▸ OK. */
export const matchFieldsCommand: Command<{
  columns: readonly string[];
  map: Readonly<Record<string, string>>;
}> = {
  id: "mailings.matchFields",
  group: "mailings",
  label: "Match Fields",
  run(model, { columns, map }) {
    setMailMergeFieldMap(model.doc, columns, map);
  },
};

/** Edit Recipient List ▸ OK: the include check boxes. */
export const recipientInclusionCommand: Command<{
  list: RecipientList;
  included: readonly boolean[];
}> = {
  id: "mailings.recipientInclusion",
  group: "mailings",
  label: "Edit Recipient List",
  run(model, { list, included }) {
    setRecipientInclusion(model.doc, list, included);
  },
};

export const insertMergeFieldCommand: Command<{ name: string }> = {
  id: "mailings.mergeField",
  group: "mailings",
  label: "Insert Merge Field",
  run(model, { name }) {
    const { paragraph, offset } = caretPoint(model);
    insertMergeField(model.doc, paragraph, offset, name);
  },
  isEnabled: hasCaret,
};

export const insertAddressBlockCommand: Command<AddressBlockOptions> = {
  id: "mailings.addressBlock",
  group: "mailings",
  label: "Address Block",
  run(model, options) {
    const { paragraph, offset } = caretPoint(model);
    insertAddressBlock(model.doc, paragraph, offset, options);
  },
  isEnabled: hasCaret,
};

export const insertGreetingLineCommand: Command<GreetingLineOptions> = {
  id: "mailings.greetingLine",
  group: "mailings",
  label: "Greeting Line",
  run(model, options) {
    const { paragraph, offset } = caretPoint(model);
    insertGreetingLine(model.doc, paragraph, offset, options);
  },
  isEnabled: hasCaret,
};

/** Rules ▸ Ask… / Fill-in… / If…Then…Else… / Merge Record # / … */
export const insertRuleCommand: Command<{ rule: MergeRule }> = {
  id: "mailings.rule",
  group: "mailings",
  label: "Rules",
  run(model, { rule }) {
    const { paragraph, offset } = caretPoint(model);
    insertMergeRule(model.doc, paragraph, offset, rule);
  },
  isEnabled: hasCaret,
};

/** Preview Results on (a record index) / off (`index` undefined), and record navigation. */
export const previewResultsCommand: Command<{
  list: RecipientList | undefined;
  index: number | undefined;
  options?: PreviewOptions;
}> = {
  id: "mailings.preview",
  group: "mailings",
  label: "Preview Results",
  run(model, { list, index, options }) {
    previewMailMerge(model.doc, list, index, options);
  },
};

/** Envelopes ▸ Add to Document. */
export const addEnvelopeCommand: Command<EnvelopeOptions> = {
  id: "mailings.envelope",
  group: "mailings",
  label: "Envelopes",
  run(model, options) {
    addEnvelope(model.doc, options);
  },
};

export const removeEnvelopeCommand: Command<void, boolean> = {
  id: "mailings.removeEnvelope",
  group: "mailings",
  label: "Remove Envelope",
  run: (model) => removeEnvelope(model.doc),
};

/** Update Labels: copy the first label's layout to every other label. */
export const updateLabelsCommand: Command<void, number> = {
  id: "mailings.updateLabels",
  group: "mailings",
  label: "Update Labels",
  run: (model) => updateLabels(model.doc),
};

export const mailingsCommands = [
  startMailMergeCommand,
  selectRecipientsCommand,
  matchFieldsCommand,
  recipientInclusionCommand,
  insertMergeFieldCommand,
  insertAddressBlockCommand,
  insertGreetingLineCommand,
  insertRuleCommand,
  previewResultsCommand,
  addEnvelopeCommand,
  removeEnvelopeCommand,
  updateLabelsCommand,
];
