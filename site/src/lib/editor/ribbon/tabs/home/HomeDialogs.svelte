<script lang="ts">
  /**
   * Every Home dialog, mounted once with the page so keyboard shortcuts
   * (⌘D, ⌘⌥M, ⌘⌥⇧S …) open them from any tab.
   */
  import { onMount } from 'svelte';
  import { getSession } from '../../../session.svelte';
  import BordersDialog from './BordersDialog.svelte';
  import FindDialog from './FindDialog.svelte';
  import FontDialog from './FontDialog.svelte';
  import ParagraphDialog from './ParagraphDialog.svelte';
  import StyleDialog from './StyleDialog.svelte';
  import TabsDialog from './TabsDialog.svelte';
  import TextDialogs from './TextDialogs.svelte';
  import { handleHomeShortcut } from './shortcuts';

  const session = getSession();

  onMount(() => {
    // Runs after the canvas' own key handling (⌘B/I/U, undo …), which marks
    // the keys it took with preventDefault.
    const onKey = (e: KeyboardEvent): void => {
      if (!e.defaultPrevented && handleHomeShortcut(e, session)) e.preventDefault();
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  });
</script>

<FontDialog />
<ParagraphDialog />
<TabsDialog />
<BordersDialog />
<StyleDialog />
<FindDialog />
<TextDialogs />
