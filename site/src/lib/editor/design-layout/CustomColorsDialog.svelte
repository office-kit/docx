<script lang="ts">
  /** Design ▸ Colors ▸ Customize Colors… (Word's "Create New Theme Colors"). */
  import { commands, resolveTheme } from '@office-kit/docx-editor';
  import { THEME_COLOR_SLOTS, type ThemeColorScheme } from '@office-kit/docx';
  import Dialog from '../Dialog.svelte';
  import { getSession } from '../session.svelte';
  import { t, type MessageKey } from '../i18n/index.svelte';

  const ID = 'design.customColors';
  const session = getSession();
  const LABELS: Readonly<Record<(typeof THEME_COLOR_SLOTS)[number], MessageKey>> = {
    dk1: 'dsn.cc.dk1',
    lt1: 'dsn.cc.lt1',
    dk2: 'dsn.cc.dk2',
    lt2: 'dsn.cc.lt2',
    accent1: 'dsn.cc.accent1',
    accent2: 'dsn.cc.accent2',
    accent3: 'dsn.cc.accent3',
    accent4: 'dsn.cc.accent4',
    accent5: 'dsn.cc.accent5',
    accent6: 'dsn.cc.accent6',
    hlink: 'dsn.cc.hlink',
    folHlink: 'dsn.cc.folHlink',
  };

  let draft = $state<ThemeColorScheme | null>(null);
  let original: ThemeColorScheme | null = null;
  let open = false;

  // Start from the document's colours each time the dialog opens, as Word does.
  $effect(() => {
    const showing = session.dialog === ID;
    if (showing && !open && session.model) {
      original = { ...resolveTheme(session.model.doc).colors, name: t('dsn.cc.defaultName') };
      draft = { ...original };
    }
    open = showing;
  });

  function set(slot: (typeof THEME_COLOR_SLOTS)[number], value: string): void {
    if (draft) draft = { ...draft, [slot]: value.slice(1).toUpperCase() };
  }

  function ok(): boolean | void {
    if (!draft || draft.name.trim() === '') return false;
    session.apply(commands.themeColorsCommand, { colors: { ...draft, name: draft.name.trim() } });
  }
</script>

<Dialog id={ID} title={t('dsn.cc.title')} onok={ok}>
  {#if draft}
    <div class="cc-layout">
      <fieldset>
        <legend>{t('dsn.cc.themeColors')}</legend>
        <div class="form-grid">
          {#each THEME_COLOR_SLOTS as slot (slot)}
            <label for="cc-{slot}">{t(LABELS[slot])}</label>
            <input id="cc-{slot}" type="color" value="#{draft[slot]}" oninput={(e) => set(slot, (e.currentTarget as HTMLInputElement).value)} />
          {/each}
        </div>
      </fieldset>
      <fieldset>
        <legend>{t('dsn.cc.sample')}</legend>
        <div class="cc-sample" style="background:#{draft.lt1};color:#{draft.dk1}">
          <div class="cc-sample-band" style="background:#{draft.dk2};color:#{draft.lt1}">{t('dsn.sample.text')}</div>
          <div class="cc-sample-bars">
            {#each ['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'] as const as k (k)}<i style="background:#{draft[k]}"></i>{/each}
          </div>
          <span style="color:#{draft.hlink};text-decoration:underline">{t('dsn.cc.hlink')}</span>
          <span style="color:#{draft.folHlink};text-decoration:underline">{t('dsn.cc.folHlink')}</span>
        </div>
      </fieldset>
    </div>
    <div class="field">
      <label for="cc-name">{t('dsn.name')}</label>
      <input id="cc-name" type="text" bind:value={draft.name} />
      <button type="button" class="push" onclick={() => (draft = original && { ...original })}>{t('dsn.reset')}</button>
    </div>
  {/if}
</Dialog>
