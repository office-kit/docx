<script lang="ts">
  /**
   * The Insert-tab dialogs, mounted once so any tab can open them by id
   * (`session.openDialog(DIALOG.crossReference)` …), plus Word's field and
   * equation gestures on the canvas: F9 updates fields, Option/Alt+F9
   * toggles field codes, double-clicking an equation edits it.
   */
  import { commands, equationAtElement } from '@office-kit/docx-editor';
  import { getSession } from '../session.svelte';
  import BookmarkDialog from './BookmarkDialog.svelte';
  import CommentDialog from './CommentDialog.svelte';
  import CrossReferenceDialog from './CrossReferenceDialog.svelte';
  import DateTimeDialog from './DateTimeDialog.svelte';
  import DropCapDialog from './DropCapDialog.svelte';
  import EquationDialog from './EquationDialog.svelte';
  import FieldDialog from './FieldDialog.svelte';
  import LinkDialog from './LinkDialog.svelte';
  import PageNumberFormatDialog from './PageNumberFormatDialog.svelte';
  import SignatureLineDialog from './SignatureLineDialog.svelte';
  import SymbolDialog from './SymbolDialog.svelte';
  import { DIALOG, fieldContext, insertUi } from './state.svelte';

  const session = getSession();

  function onkeydown(e: KeyboardEvent): void {
    if (e.key !== 'F9' || session.dialog) return;
    e.preventDefault();
    if (e.altKey) session.showFieldCodes = !session.showFieldCodes;
    else session.apply(commands.updateFieldsCommand, fieldContext(session));
  }

  function ondblclick(e: MouseEvent): void {
    const model = session.model;
    const el = e.target instanceof Element ? e.target.closest('.wk-math') : null;
    if (!model || !el) return;
    const ref = equationAtElement(model, el);
    if (!ref) return;
    insertUi.equation = ref;
    session.openDialog(DIALOG.equation);
  }
</script>

<svelte:window {onkeydown} {ondblclick} />

<LinkDialog />
<BookmarkDialog />
<CrossReferenceDialog />
<CommentDialog />
<FieldDialog />
<DateTimeDialog />
<SymbolDialog />
<EquationDialog />
<DropCapDialog />
<PageNumberFormatDialog />
<SignatureLineDialog />
