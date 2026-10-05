<script lang="ts">
  /** Insert Address Block, Insert Greeting Line, and the Rules dialogs. */
  import {
    ADDRESS_FIELDS,
    ADDRESS_NAME_FORMATS,
    GREETING_NAME_FORMATS,
    type MergeComparison,
    type MergeRule,
  } from '@office-kit/docx';
  import { commands } from '@office-kit/docx-editor';
  import Dialog from '../Dialog.svelte';
  import { getSession } from '../session.svelte';
  import { t } from '../i18n/index.svelte';
  import { merge } from './merge.svelte';
  import { ruleDialog } from './rule.svelte';

  const session = getSession();
  type NameFormat = keyof typeof ADDRESS_NAME_FORMATS;
  type GreetingFormat = keyof typeof GREETING_NAME_FORMATS;
  const NAME_FORMATS = Object.keys(ADDRESS_NAME_FORMATS).filter((k): k is NameFormat => Object.hasOwn(ADDRESS_NAME_FORMATS, k));
  const GREETING_FORMATS = Object.keys(GREETING_NAME_FORMATS).filter((k): k is GreetingFormat => Object.hasOwn(GREETING_NAME_FORMATS, k));
  // Word's own choices in the Greeting Line drop-downs.
  const SALUTATIONS = ['Dear', 'To', ''] as const;
  const PUNCTUATION = [',', ':', ''] as const;
  const INVALID_GREETINGS = ['Dear Sir or Madam,', 'To Whom It May Concern:', ''] as const;

  // --- Address Block -------------------------------------------------------------------
  let nameFormat = $state<NameFormat>('Mr. Joshua Randall Jr.');
  let company = $state(true);
  let postal = $state(true);
  let country = $state<'never' | 'always' | 'ifDifferent'>('ifDifferent');
  let excludedCountry = $state('United States');
  let formatByCountry = $state(true);

  function insertAddressBlock(): void {
    session.apply(commands.insertAddressBlockCommand, {
      nameFormat: ADDRESS_NAME_FORMATS[nameFormat],
      company,
      postalAddress: postal,
      country,
      excludedCountry,
      formatByCountry,
    });
    if (merge.preview) merge.show(session);
  }

  // --- Greeting Line -------------------------------------------------------------------
  let salutation = $state<string>('Dear');
  let greetingFormat = $state<GreetingFormat>('Mr. Randall');
  let punctuation = $state<string>(',');
  let invalidGreeting = $state<string>('Dear Sir or Madam,');

  function insertGreetingLine(): void {
    session.apply(commands.insertGreetingLineCommand, {
      salutation,
      nameFormat: GREETING_NAME_FORMATS[greetingFormat],
      punctuation,
      invalidNameGreeting: invalidGreeting,
    });
    if (merge.preview) merge.show(session);
  }

  // --- Rules ---------------------------------------------------------------------------
  const COMPARISONS: readonly MergeComparison[] = ['=', '<>', '<', '<=', '>', '>=', 'isBlank', 'isNotBlank'];
  const COMPARISON_KEYS = {
    '=': 'mail.cmp.equals',
    '<>': 'mail.cmp.notEquals',
    '<': 'mail.cmp.less',
    '<=': 'mail.cmp.lessOrEqual',
    '>': 'mail.cmp.greater',
    '>=': 'mail.cmp.greaterOrEqual',
    isBlank: 'mail.cmp.isBlank',
    isNotBlank: 'mail.cmp.isNotBlank',
  } as const satisfies Record<MergeComparison, string>;
  const fieldNames = $derived(merge.list?.columns ?? ADDRESS_FIELDS);
  let bookmark = $state('');
  let prompt = $state('');
  let defaultText = $state('');
  let askOnce = $state(false);
  let field = $state('');
  let comparison = $state<MergeComparison>('=');
  let value = $state('');
  let trueText = $state('');
  let falseText = $state('');
  const kind = $derived(ruleDialog.kind);
  const needsField = $derived(kind === 'if' || kind === 'nextRecordIf' || kind === 'skipRecordIf');
  const blankTest = $derived(comparison === 'isBlank' || comparison === 'isNotBlank');

  function buildRule(): MergeRule {
    const column = field || fieldNames[0] || '';
    const compared = { field: column, comparison, ...(blankTest ? {} : { value }) };
    const def = defaultText ? { defaultText } : {};
    switch (kind) {
      case 'ask': return { kind, bookmark, prompt, askOnce, ...def };
      case 'fillIn': return { kind, prompt, askOnce, ...def };
      case 'if': return { kind, ...compared, trueText, falseText };
      case 'nextRecordIf': return { kind, ...compared };
      case 'skipRecordIf': return { kind, ...compared };
      case 'setBookmark': return { kind, bookmark, value };
      case 'mergeRecord':
      case 'mergeSequence':
      case 'nextRecord': return { kind };
    }
  }

  function insertRule(): boolean {
    session.apply(commands.insertRuleCommand, { rule: buildRule() });
    if (session.status) return false;
    if (merge.preview) merge.show(session);
    return true;
  }
</script>

<Dialog id="mailings.addressBlock" title={t('mail.addressBlockTitle')} onok={insertAddressBlock}>
  <label class="field">{t('mail.recipientName')}
    <select bind:value={nameFormat}>
      {#each NAME_FORMATS as f (f)}<option value={f}>{f}</option>{/each}
    </select>
  </label>
  <label><input type="checkbox" bind:checked={company} /> {t('mail.companyName')}</label>
  <label><input type="checkbox" bind:checked={postal} /> {t('mail.postalAddress')}</label>
  <fieldset disabled={!postal}>
    <label><input type="radio" bind:group={country} value="never" /> {t('mail.countryNever')}</label>
    <label><input type="radio" bind:group={country} value="always" /> {t('mail.countryAlways')}</label>
    <label><input type="radio" bind:group={country} value="ifDifferent" /> {t('mail.countryIfDifferent')}</label>
    <input bind:value={excludedCountry} disabled={country !== 'ifDifferent'} aria-label={t('mail.countryIfDifferent')} />
    <label><input type="checkbox" bind:checked={formatByCountry} /> {t('mail.formatByCountry')}</label>
  </fieldset>
</Dialog>

<Dialog id="mailings.greetingLine" title={t('mail.greetingLineTitle')} onok={insertGreetingLine}>
  <div class="row">
    <select bind:value={salutation} aria-label={t('mail.salutation')}>
      {#each SALUTATIONS as s (s)}<option value={s}>{s || t('mail.none')}</option>{/each}
    </select>
    <select bind:value={greetingFormat} aria-label={t('mail.recipientName')}>
      {#each GREETING_FORMATS as f (f)}<option value={f}>{f}</option>{/each}
    </select>
    <select bind:value={punctuation} aria-label={t('mail.punctuation')}>
      {#each PUNCTUATION as p (p)}<option value={p}>{p || t('mail.none')}</option>{/each}
    </select>
  </div>
  <label class="field">{t('mail.invalidNames')}
    <select bind:value={invalidGreeting}>
      {#each INVALID_GREETINGS as g (g)}<option value={g}>{g || t('mail.none')}</option>{/each}
    </select>
  </label>
</Dialog>

<Dialog id="mailings.rule" title={t(`mail.rule.${kind}`)} onok={insertRule}>
  {#if kind === 'ask' || kind === 'setBookmark'}
    <label class="field">{t('mail.bookmark')} <input bind:value={bookmark} /></label>
  {/if}
  {#if kind === 'ask' || kind === 'fillIn'}
    <label class="field">{t('mail.prompt')} <input bind:value={prompt} size="36" /></label>
    <label class="field">{t('mail.defaultText')} <input bind:value={defaultText} size="36" /></label>
    <label><input type="checkbox" bind:checked={askOnce} /> {t('mail.askOnce')}</label>
  {/if}
  {#if needsField}
    <div class="row">
      <label class="field">{t('mail.fieldName')}
        <select bind:value={field}>
          {#each fieldNames as f (f)}<option value={f}>{f}</option>{/each}
        </select>
      </label>
      <label class="field">{t('mail.comparison')}
        <select bind:value={comparison}>
          {#each COMPARISONS as c (c)}<option value={c}>{t(COMPARISON_KEYS[c])}</option>{/each}
        </select>
      </label>
      <label class="field">{t('mail.compareTo')} <input bind:value={value} disabled={blankTest} /></label>
    </div>
  {/if}
  {#if kind === 'setBookmark'}
    <label class="field">{t('mail.value')} <input bind:value={value} size="36" /></label>
  {/if}
  {#if kind === 'if'}
    <label class="field">{t('mail.trueText')}<textarea bind:value={trueText} rows="2" cols="40"></textarea></label>
    <label class="field">{t('mail.falseText')}<textarea bind:value={falseText} rows="2" cols="40"></textarea></label>
  {/if}
</Dialog>
