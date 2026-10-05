import type { MergeRule } from "@office-kit/docx";

/** Which Rules ▸ item the shared rule dialog is editing. */
class RuleDialogState {
  kind = $state<MergeRule["kind"]>("ask");
}

export const ruleDialog = new RuleDialogState();
