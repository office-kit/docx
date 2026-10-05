<script lang="ts">
  /**
   * Insert ▸ Field: every ECMA-376 field by Word's categories. Picking options
   * (argument, `\@` / `\#` / `\*` formats, field switches, "Preserve
   * formatting during updates" → `\* MERGEFORMAT`) writes the field code,
   * which stays editable for anything the options do not cover. The field is
   * inserted with its computed result.
   */
  import { formatDatePicture, listStyles, quoteFieldArgument } from '@office-kit/docx';
  import { commands, documentBookmarks } from '@office-kit/docx-editor';
  import Dialog from '../Dialog.svelte';
  import { getSession } from '../session.svelte';
  import { locale, t } from '../i18n/index.svelte';
  import { DATE_FORMATS, DATE_LANGUAGE } from './date-formats';
  import { FIELD_CATEGORIES, FIELDS, NUMBER_FORMATS, NUMERIC_PICTURES, TEXT_FORMATS, type FieldCategory, type FieldDef } from './field-catalog';
  import { DIALOG, fieldContext } from './state.svelte';

  const session = getSession();
  const CATEGORY_KEYS = {
    dateTime: 'ins.field.c.dateTime',
    automation: 'ins.field.c.automation',
    info: 'ins.field.c.info',
    formulas: 'ins.field.c.formulas',
    index: 'ins.field.c.index',
    links: 'ins.field.c.links',
    mailMerge: 'ins.field.c.mailMerge',
    numbering: 'ins.field.c.numbering',
    user: 'ins.field.c.user',
    forms: 'ins.field.c.forms',
  } as const satisfies Record<FieldCategory, string>;
  const ARG_LABEL_KEYS = {
    identifier: 'ins.field.identifier',
    formula: 'ins.field.formula',
    text: 'ins.field.text',
    code: 'ins.field.code',
  } as const;
  const DEFAULT_FIELD = 'DATE';

  let category = $state<FieldCategory | 'all'>('all');
  let name = $state(DEFAULT_FIELD);
  let arg = $state('');
  let picture = $state('');
  let format = $state('');
  let switches = $state<string[]>([]);
  let preserve = $state(true);
  let code = $state('');
  // Form field options.
  let formText = $state('');
  let formChecked = $state(false);
  let formEntries = $state('');

  const listed = $derived(category === 'all' ? FIELDS.toSorted((a, b) => a.name.localeCompare(b.name)) : FIELDS.filter((f) => f.category === category));
  const def = $derived<FieldDef | undefined>(FIELDS.find((f) => f.name === name));
  const doc = $derived(session.dialog === DIALOG.field && session.version >= 0 ? session.model?.doc : undefined);
  const bookmarkNames = $derived(doc ? documentBookmarks(doc).map((b) => b.name) : []);
  const styleIds = $derived(doc ? listStyles(doc).map((s) => s.styleId) : []);
  const datePictures = $derived(DATE_FORMATS[locale()]);
  const isForm = $derived(def?.arg?.kind === 'form');

  $effect(() => {
    if (session.dialog !== DIALOG.field) return;
    category = 'all';
    choose(DEFAULT_FIELD);
  });

  function choose(next: string): void {
    name = next;
    arg = '';
    picture = '';
    format = '';
    switches = [];
    preserve = FIELDS.find((f) => f.name === next)?.arg?.kind !== 'form';
  }

  const built = $derived.by(() => {
    if (!def) return '';
    const parts = [def.name];
    const value = arg.trim();
    if (value) parts.push(def.arg?.kind === 'text' && def.arg.raw ? value : quoteFieldArgument(value));
    parts.push(...switches);
    if (picture) parts.push(def.result === 'date' ? `\\@ ${quoteFieldArgument(picture)}` : `\\# ${quoteFieldArgument(picture)}`);
    if (format) parts.push(`\\* ${format}`);
    if (preserve) parts.push('\\* MERGEFORMAT');
    return parts.join(' ');
  });

  // Options rewrite the code; edits typed into the code box last until the
  // next option change, as in Word's Field Codes view.
  $effect(() => {
    code = built;
  });

  function toggleSwitch(s: string, on: boolean): void {
    switches = on ? [...switches, s] : switches.filter((x) => x !== s);
  }

  function ok(): boolean {
    if (!def) return false;
    if (def.arg?.kind === 'form') {
      const nameOpt = arg.trim() ? { name: arg.trim() } : {};
      const options =
        def.arg.form === 'text'
          ? { kind: 'text' as const, ...nameOpt, ...(formText ? { defaultText: formText } : {}) }
          : def.arg.form === 'checkBox'
            ? { kind: 'checkBox' as const, ...nameOpt, checked: formChecked }
            : { kind: 'dropDown' as const, ...nameOpt, entries: formEntries.split('\n').map((e) => e.trim()).filter(Boolean) };
      session.apply(commands.insertFormFieldCommand, options);
    } else {
      session.apply(commands.insertFieldAtCaretCommand, {
        instruction: code,
        context: fieldContext(session),
        ...(def.result === 'date' ? { lang: DATE_LANGUAGE[locale()] } : {}),
      });
    }
    return session.status === '';
  }
</script>

<Dialog id={DIALOG.field} title={t('ins.field')} onok={ok} okDisabled={!def || (!isForm && !code.trim())}>
  <div class="dialog-row">
    <div class="dialog-col">
      <label class="stack">{t('ins.field.categories')}
        <select value={category} onchange={(e) => { const v = e.currentTarget.value; category = FIELD_CATEGORIES.find((c) => c === v) ?? 'all'; }}>
          <option value="all">{t('ins.field.c.all')}</option>
          {#each FIELD_CATEGORIES as c (c)}<option value={c}>{t(CATEGORY_KEYS[c])}</option>{/each}
        </select>
      </label>
      <span>{t('ins.field.names')}</span>
      <div class="listbox" role="listbox" aria-label={t('ins.field.names')} style="height: 260px">
        {#each listed as f (f.name)}
          <button type="button" role="option" aria-selected={f.name === name} class:selected={f.name === name} onclick={() => choose(f.name)}>{f.name}</button>
        {/each}
      </div>
    </div>
    <div class="dialog-col">
      {#if def}
        <p class="syntax">{def.syntax}</p>
        {#if def.arg?.kind === 'bookmark'}
          <label class="stack">{t('ins.field.bookmark')}
            <select bind:value={arg}>
              <option value="">{t('ins.field.none')}</option>
              {#each bookmarkNames as b (b)}<option value={b}>{b}</option>{/each}
            </select>
          </label>
        {:else if def.arg?.kind === 'property'}
          <label class="stack">{t('ins.field.property')}
            <select bind:value={arg}>
              <option value="">{t('ins.field.none')}</option>
              {#each def.arg.choices as p (p)}<option value={p}>{p}</option>{/each}
            </select>
          </label>
        {:else if def.arg?.kind === 'style'}
          <label class="stack">{t('ins.field.style')}
            <select bind:value={arg}>
              <option value="">{t('ins.field.none')}</option>
              {#each styleIds as s (s)}<option value={s}>{s}</option>{/each}
            </select>
          </label>
        {:else if def.arg?.kind === 'text'}
          <label class="stack">{t(ARG_LABEL_KEYS[def.arg.label])}<input type="text" bind:value={arg} /></label>
        {:else if def.arg?.kind === 'form'}
          <label class="stack">{t('ins.bm.name')}<input type="text" bind:value={arg} maxlength="20" /></label>
          {#if def.arg.form === 'text'}
            <label class="stack">{t('ins.field.text')}<input type="text" bind:value={formText} /></label>
          {:else if def.arg.form === 'checkBox'}
            <label><input type="checkbox" bind:checked={formChecked} /> {t('ins.field.checked')}</label>
          {:else}
            <label class="stack">{t('ins.field.entries')}<textarea rows="4" bind:value={formEntries}></textarea></label>
          {/if}
        {/if}
        {#if def.result === 'date'}
          <span>{t('ins.field.dateFormat')}</span>
          <div class="listbox" role="listbox" aria-label={t('ins.field.dateFormat')} style="height: 120px">
            {#each datePictures as p (p)}
              <button type="button" role="option" aria-selected={p === picture} class:selected={p === picture} onclick={() => (picture = p)}>{formatDatePicture(new Date(), p, DATE_LANGUAGE[locale()])}</button>
            {/each}
          </div>
        {:else if def.result === 'number' || def.result === 'text'}
          <label class="stack">{t('ins.field.format')}
            <select bind:value={format}>
              <option value="">{t('ins.field.none')}</option>
              {#each def.result === 'number' ? NUMBER_FORMATS : TEXT_FORMATS as f (f)}<option value={f}>{f}</option>{/each}
            </select>
          </label>
          {#if def.result === 'number'}
            <label class="stack">{t('ins.field.options')} (\#)
              <select bind:value={picture}>
                <option value="">{t('ins.field.none')}</option>
                {#each NUMERIC_PICTURES as p (p)}<option value={p}>{p}</option>{/each}
              </select>
            </label>
          {/if}
        {/if}
        {#if def.switches?.length}
          <fieldset>
            <legend>{t('ins.field.options')}</legend>
            {#each def.switches as s (s)}
              <label style="margin-right: 8px"><input type="checkbox" checked={switches.includes(s)} onchange={(e) => toggleSwitch(s, e.currentTarget.checked)} /> <code>{s}</code></label>
            {/each}
          </fieldset>
        {/if}
        {#if !isForm}
          <label><input type="checkbox" bind:checked={preserve} /> {t('ins.field.preserve')}</label>
          <label class="stack">{t('ins.field.codes')}<input type="text" bind:value={code} spellcheck="false" /></label>
        {/if}
      {/if}
    </div>
  </div>
  {#if session.status}<p class="error-note">{session.status}</p>{/if}
</Dialog>
