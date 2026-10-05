<script lang="ts">
  /** Every References-tab dialog, mounted with the tab. */
  import CaptionDialog from './CaptionDialog.svelte';
  import IndexDialogs from './IndexDialogs.svelte';
  import NoteDialog from './NoteDialog.svelte';
  import SourceDialogs from './SourceDialogs.svelte';
  import TocDialogs from './TocDialogs.svelte';
</script>

<TocDialogs />
<NoteDialog />
<SourceDialogs />
<CaptionDialog />
<IndexDialogs />
