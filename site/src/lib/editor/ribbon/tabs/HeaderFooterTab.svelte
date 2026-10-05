<script lang="ts">
  /**
   * Word's contextual Header & Footer tab, shown while a header or footer is
   * open for editing: the Header / Footer / Page Number menus, navigation
   * between headers, footers and sections, Link to Previous, the
   * first-page / odd-even / show-text options, the distances, and Close.
   */
  import { commands, documentSections } from '@office-kit/docx-editor';
  import Button from '../Button.svelte';
  import SplitButton from '../SplitButton.svelte';
  import Group from '../Group.svelte';
  import { getSession, type HeaderFooterTarget } from '../../session.svelte';
  import { locale, t } from '../../i18n/index.svelte';

  const session = getSession();

  // Word for Mac shows inches in English and centimetres elsewhere.
  const TWIPS_PER_INCH = 1440;
  const TWIPS_PER_CM = 1440 / 2.54;
  const unit = $derived(locale() === 'en' ? { twips: TWIPS_PER_INCH, suffix: '"', step: 0.1 } : { twips: TWIPS_PER_CM, suffix: ' cm', step: 0.1 });

  /** The header/footer being edited (the canvas sets it when one is opened). */
  const target = $derived.by((): HeaderFooterTarget => {
    const story = session.story;
    const current = session.headerFooter;
    if (current && story && current.kind === story.kind) return current;
    const kind = story?.kind === 'footer' ? 'footer' : 'header';
    const page = session.pages[session.currentPage - 1];
    return { section: page?.section ?? 0, kind, type: page?.kind ?? 'default' };
  });

  const section = $derived(session.tick >= 0 && session.model ? documentSections(session.model.document)[target.section] : undefined);
  const sectionCount = $derived(session.tick >= 0 && session.model ? documentSections(session.model.document).length : 1);
  const linked = $derived(session.tick >= 0 && !!session.model && commands.headerFooterLinked(session.model, target));
  const firstPage = $derived(session.tick >= 0 && !!session.model && commands.hasDifferentFirstPage(session.model, target.section));

  function show(value: number | undefined): string {
    return value === undefined ? '' : `${Math.round((value / unit.twips) * 100) / 100}`;
  }

  function setDistance(kind: 'header' | 'footer', text: string): void {
    const value = Number.parseFloat(text);
    if (!Number.isFinite(value)) return;
    const twips = Math.round(value * unit.twips);
    session.apply(kind === 'header' ? commands.headerDistanceCommand : commands.footerDistanceCommand, { section: target.section, twips });
  }

  /** Open the same kind in another section (its first page), as Word's Previous / Next do. */
  function goToSection(delta: number): void {
    const next = target.section + delta;
    const page = session.pages.findIndex((p) => p.section === next);
    if (page < 0) return;
    session.editHeaderFooter(target.kind, page);
  }

  function goTo(kind: 'header' | 'footer'): void {
    session.editHeaderFooter(kind, session.currentPage - 1);
  }
</script>

<Group label={t('group.headerFooter')}>
  <SplitButton id="hf.header" size="large" icon="header" tip={t('hf.header')}>
    {#snippet menu()}
      <button class="mi" onclick={() => goTo('header')}>{t('hf.goToHeader')}</button>
    {/snippet}
  </SplitButton>
  <SplitButton id="hf.footer" size="large" icon="footer" tip={t('hf.footer')}>
    {#snippet menu()}
      <button class="mi" onclick={() => goTo('footer')}>{t('hf.goToFooter')}</button>
    {/snippet}
  </SplitButton>
  <SplitButton id="hf.pageNumber" size="large" icon="pageNumber" tip={t('hf.pageNumber')}>
    {#snippet menu()}
      <button class="mi" onclick={() => session.apply(commands.addPageNumberFooterCommand, {})}>{t('hf.pageNumber')}</button>
    {/snippet}
  </SplitButton>
</Group>
<Group label={t('hf.goToHeader')}>
  <Button size="large" icon="goToHeader" tip={t('hf.goToHeader')} disabled={target.kind === 'header'} onclick={() => goTo('header')} />
  <Button size="large" icon="goToFooter" tip={t('hf.goToFooter')} disabled={target.kind === 'footer'} onclick={() => goTo('footer')} />
  <div class="col">
    <Button size="mid" icon="previousSection" label={t('hf.previous')} tip={t('hf.previousTip')} disabled={target.section === 0} onclick={() => goToSection(-1)} />
    <Button size="mid" icon="nextSection" label={t('hf.next')} tip={t('hf.nextTip')} disabled={target.section >= sectionCount - 1} onclick={() => goToSection(1)} />
    <Button
      size="mid"
      icon="linkToPrevious"
      tip={t('hf.linkToPrevious')}
      on={linked}
      disabled={target.section === 0}
      onclick={() => session.apply(commands.linkToPreviousCommand, { ...target, linked: !linked })}
    />
  </div>
</Group>
<Group label={t('hf.differentFirst')}>
  <label class="check-item"
    ><input type="checkbox" checked={firstPage} onchange={(e) => session.apply(commands.differentFirstPageCommand, { section: target.section, on: e.currentTarget.checked })} />{t('hf.differentFirst')}</label
  >
  <label class="check-item"
    ><input type="checkbox" checked={session.active(commands.differentOddEvenCommand)} onchange={(e) => session.apply(commands.differentOddEvenCommand, { on: e.currentTarget.checked })} />{t('hf.differentOddEven')}</label
  >
  <label class="check-item"><input type="checkbox" bind:checked={session.showDocumentText} />{t('hf.showDocumentText')}</label>
</Group>
<Group label={t('hf.headerFromTop')}>
  <div class="rows">
    <label class="field"
      >{t('hf.headerFromTop')}<input type="number" min="0" step={unit.step} value={show(section?.margins.header)} onchange={(e) => setDistance('header', e.currentTarget.value)} />{unit.suffix}</label
    >
    <label class="field"
      >{t('hf.footerFromBottom')}<input type="number" min="0" step={unit.step} value={show(section?.margins.footer)} onchange={(e) => setDistance('footer', e.currentTarget.value)} />{unit.suffix}</label
    >
  </div>
</Group>
<Group label={t('hf.close')}>
  <Button size="large" icon="closeHeaderFooter" tip={t('hf.close')} onclick={() => session.closeHeaderFooter()} />
</Group>
