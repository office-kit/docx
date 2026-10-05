<script lang="ts">
  /** The raw-XML inspector: reaches every element the ribbon does not surface. */
  import RawXmlInspector from '../RawXmlInspector.svelte';
  import { getSession } from '../session.svelte';
  import { t } from '../i18n/index.svelte';

  const session = getSession();
</script>

<div class="pane-head">{t('xml.title')}</div>
{#if session.model}
  <RawXmlInspector model={session.model} onchange={() => { session.changed(); session.status = 'Raw XML edited.'; }} />
{/if}
