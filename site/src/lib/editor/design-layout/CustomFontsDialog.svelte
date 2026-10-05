<script lang="ts">
  /** Design ▸ Fonts ▸ Customize Fonts… (Word's "Create New Theme Fonts"). */
  import { commands, resolveTheme } from '@office-kit/docx-editor';
  import { THEME_FONT_SCHEMES } from '@office-kit/docx';
  import Dialog from '../Dialog.svelte';
  import { getSession } from '../session.svelte';
  import { t } from '../i18n/index.svelte';

  const ID = 'design.customFonts';
  const session = getSession();
  // The fonts Word lists first, plus every font a built-in scheme uses.
  const FONT_CHOICES = [...new Set(THEME_FONT_SCHEMES.flatMap((s) => [s.major, s.minor]))].toSorted();

  let open = $state(false);
  let major = $state('');
  let minor = $state('');
  let name = $state('');

  $effect(() => {
    const showing = session.dialog === ID;
    if (showing && !open && session.model) {
      const fonts = resolveTheme(session.model.doc).fonts;
      major = fonts.major;
      minor = fonts.minor;
      name = t('dsn.cf.defaultName');
    }
    open = showing;
  });

  function ok(): boolean | void {
    if (!major.trim() || !minor.trim() || !name.trim()) return false;
    session.apply(commands.themeFontsCommand, { fonts: { name: name.trim(), major: major.trim(), minor: minor.trim() } });
  }
</script>

<Dialog id={ID} title={t('dsn.cf.title')} onok={ok}>
  <datalist id="cf-fonts">{#each FONT_CHOICES as f (f)}<option value={f}></option>{/each}</datalist>
  <div class="form-grid">
    <label for="cf-major">{t('dsn.cf.heading')}</label>
    <input id="cf-major" type="text" list="cf-fonts" bind:value={major} />
    <label for="cf-minor">{t('dsn.cf.body')}</label>
    <input id="cf-minor" type="text" list="cf-fonts" bind:value={minor} />
  </div>
  <fieldset>
    <legend>{t('dsn.cc.sample')}</legend>
    <div class="cf-sample">
      <div style="font-family:'{major}',sans-serif;font-size:18px">{t('dsn.cf.headingSample')}</div>
      <div style="font-family:'{minor}',sans-serif">{t('dsn.cf.bodySample')}</div>
    </div>
  </fieldset>
  <div class="field">
    <label for="cf-name">{t('dsn.name')}</label>
    <input id="cf-name" type="text" bind:value={name} />
  </div>
</Dialog>
