<script lang="ts">
  /** Format ▸ Tabs: the selected paragraphs' custom tab stops. */
  import { getParagraphTabs, type TabAlignment, type TabLeader, type TabStop } from '@office-kit/docx';
  import { commands, paragraphAt } from '@office-kit/docx-editor';
  import Dialog from '../../../Dialog.svelte';
  import { getSession } from '../../../session.svelte';
  import { locale, t, type MessageKey } from '../../../i18n/index.svelte';
  import { fromTwips, toTwips, unitFor } from './units';

  const session = getSession();
  const ID = 'home.tabs';
  const ALIGNMENTS: ReadonlyArray<{ value: Exclude<TabAlignment, 'clear'>; key: MessageKey }> = [
    { value: 'left', key: 'home.tabs.left' },
    { value: 'center', key: 'home.tabs.center' },
    { value: 'right', key: 'home.tabs.right' },
    { value: 'decimal', key: 'home.tabs.decimal' },
    { value: 'bar', key: 'home.tabs.bar' },
  ];
  const LEADERS: ReadonlyArray<{ value: TabLeader; label: string }> = [
    { value: 'none', label: '—' },
    { value: 'dot', label: '.......' },
    { value: 'hyphen', label: '-------' },
    { value: 'underscore', label: '______' },
    { value: 'middleDot', label: '·······' },
    { value: 'heavy', label: '▬▬▬' },
  ];

  let stops = $state<TabStop[]>([]);
  let position = $state(0);
  let alignment = $state<Exclude<TabAlignment, 'clear'>>('left');
  let leader = $state<TabLeader>('none');
  const unit = $derived(unitFor(locale()));

  $effect(() => {
    if (session.dialog !== ID) return;
    const model = session.model;
    const focus = model?.selection?.focus;
    const para = model && focus ? paragraphAt(model.doc, focus) : undefined;
    stops = para ? getParagraphTabs(para) : [];
  });

  function set(): void {
    const at = toTwips(position, unit);
    stops = [...stops.filter((s) => s.position !== at), { position: at, alignment, leader }].toSorted((a, b) => a.position - b.position);
  }
</script>

<Dialog id={ID} title={t('home.tabs.dialog')} onok={() => { session.apply(commands.setTabsCommand, { tabs: stops }); }}>
  <div class="tab-stops">
    <div class="stop-list" role="listbox" aria-label={t('home.tabs.stops')}>
      {#each stops as stop (stop.position)}
        <button type="button" class="mi" role="option" aria-selected="false" onclick={() => { position = fromTwips(stop.position, unit); alignment = stop.alignment === 'clear' ? 'left' : stop.alignment; leader = stop.leader ?? 'none'; }}>
          {fromTwips(stop.position, unit)} {unit} · {stop.alignment}
        </button>
      {/each}
    </div>
    <div class="form-grid">
      <label class="field">{t('home.tabs.position')}<input type="number" min="0" step="0.1" bind:value={position} />{unit}</label>
      <label class="field">{t('home.para.alignment')}
        <select bind:value={alignment}>{#each ALIGNMENTS as a (a.value)}<option value={a.value}>{t(a.key)}</option>{/each}</select>
      </label>
      <label class="field">{t('home.tabs.leader')}
        <select bind:value={leader}>{#each LEADERS as l (l.value)}<option value={l.value}>{l.label}</option>{/each}</select>
      </label>
    </div>
  </div>
  <div class="field">
    <button type="button" class="push" onclick={set}>{t('home.tabs.set')}</button>
    <button type="button" class="push" onclick={() => (stops = stops.filter((s) => s.position !== toTwips(position, unit)))}>{t('home.tabs.clear')}</button>
    <button type="button" class="push" onclick={() => (stops = [])}>{t('home.tabs.clearAll')}</button>
  </div>
</Dialog>
