<script lang="ts">
  /**
   * Word's Insert tab, group by group in Word for Mac's order: Pages, Tables,
   * Illustrations, Links, Comments, Header & Footer, Text, Symbols. (Media is
   * not offered: online video is a Word extension outside ECMA-376.)
   */
  import { buildEquation } from '@office-kit/docx';
  import { commands, renderMath } from '@office-kit/docx-editor';
  import RibbonIcon from '../../RibbonIcon.svelte';
  import Button from '../Button.svelte';
  import Group from '../Group.svelte';
  import SplitButton from '../SplitButton.svelte';
  import TableButton from '../TableButton.svelte';
  import IllustrationButtons from '../IllustrationButtons.svelte';
  import TextBoxButtons from '../TextBoxButtons.svelte';
  import FooterButton from '../../insert/FooterButton.svelte';
  import HeaderButton from '../../insert/HeaderButton.svelte';
  import PageNumberButton from '../../insert/PageNumberButton.svelte';
  import { BUILT_IN_EQUATIONS } from '../../insert/equations';
  import { DIALOG, fieldContext, insertUi } from '../../insert/state.svelte';
  import { getSession } from '../../session.svelte';
  import { t } from '../../i18n/index.svelte';

  const session = getSession();
  const hasCaret = $derived(session.tick >= 0 && !!session.model?.selection);
  const COVER_PAGES = commands.COVER_PAGE_PRESETS;
  // Document Property quick parts: DOCPROPERTY fields Word's menu offers.
  const DOC_PROPERTIES = ['Author', 'Title', 'Subject', 'Keywords', 'Comments', 'Category', 'Company', 'Manager'] as const;

  // Thumbnails are the library's own MathML for each built-in equation.
  const equationPreviews = BUILT_IN_EQUATIONS.map((e) => ({ key: e.key, linear: e.linear, html: renderMath(buildEquation(e.linear)) }));

  function newEquation(): void {
    insertUi.equation = null;
    session.openDialog(DIALOG.equation);
  }

  async function textFromFile(e: Event): Promise<void> {
    const input = e.currentTarget;
    if (!(input instanceof HTMLInputElement)) return;
    const file = input.files?.[0];
    input.value = '';
    if (!file) return;
    session.openMenu = null;
    session.apply(commands.insertTextFromFileCommand, { bytes: new Uint8Array(await file.arrayBuffer()) });
  }
</script>

<Group label={t('group.pages')}>
  <SplitButton id="insert.coverPage" size="large" icon="coverPage" tip={t('ins.coverPage')}>
    {#snippet menu()}
      <div class="menu-head">{t('ins.coverPage.builtIn')}</div>
      <div class="gallery-menu">
        {#each COVER_PAGES as preset (preset)}
          <button class="tile" role="menuitem" onclick={() => session.apply(commands.insertCoverPageCommand, { preset })}>
            <span class="preview-page" aria-hidden="true"><span class="preview-mark preset-{preset}"></span></span>
            <span class="tile-name">{t(`ins.preset.${preset}`)}</span>
          </button>
        {/each}
      </div>
      <hr />
      <button class="mi" role="menuitem" onclick={() => session.apply(commands.removeCoverPageCommand, undefined)}>{t('ins.coverPage.remove')}</button>
    {/snippet}
  </SplitButton>
  <Button size="large" icon="blankPage" tip={t('ins.blankPage')} onclick={() => session.apply(commands.insertBlankPageCommand, undefined)} disabled={!hasCaret} />
  <Button size="large" icon="pageBreak" tip={t('ins.pageBreak')} onclick={() => session.apply(commands.insertPageBreakCommand, undefined)} disabled={!hasCaret} />
</Group>
<Group label={t('group.tables')}>
  <TableButton />
</Group>
<Group label={t('group.illustrations')}>
  <IllustrationButtons />
</Group>
<Group label={t('group.links')}>
  <Button size="large" icon="link" tip={t('ins.link')} onclick={() => session.openDialog(DIALOG.link)} disabled={!hasCaret} />
  <Button size="large" icon="bookmark" tip={t('ref.bookmark')} onclick={() => session.openDialog(DIALOG.bookmark)} />
  <Button size="large" icon="crossReference" tip={t('ins.crossReference')} onclick={() => session.openDialog(DIALOG.crossReference)} disabled={!hasCaret} />
</Group>
<Group label={t('group.comments')}>
  <Button size="large" icon="newComment" tip={t('ins.comment')} onclick={() => session.openDialog(DIALOG.comment)} disabled={!hasCaret} />
</Group>
<Group label={t('group.headerFooter')}>
  <HeaderButton />
  <FooterButton />
  <PageNumberButton />
</Group>
<Group label={t('ins.group.text')}>
  <TextBoxButtons />
  <SplitButton id="insert.dropCap" size="large" icon="dropCap" tip={t('ins.dropCap')} disabled={!hasCaret}>
    {#snippet menu()}
      <button class="mi" role="menuitem" onclick={() => session.apply(commands.dropCapCommand, { position: 'none' })}>{t('ins.dropCap.none')}</button>
      <button class="mi" role="menuitem" onclick={() => session.apply(commands.dropCapCommand, { position: 'drop' })}>{t('ins.dropCap.dropped')}</button>
      <button class="mi" role="menuitem" onclick={() => session.apply(commands.dropCapCommand, { position: 'margin' })}>{t('ins.dropCap.inMargin')}</button>
      <hr />
      <button class="mi" role="menuitem" onclick={() => session.openDialog(DIALOG.dropCap)}>{t('ins.dropCap.options')}</button>
    {/snippet}
  </SplitButton>
  <div class="col">
    <SplitButton id="insert.field" size="mid" icon="field" tip={t('ins.field')} onclick={() => session.openDialog(DIALOG.field)} disabled={!hasCaret}>
      {#snippet menu()}
        <button class="mi" role="menuitem" onclick={() => session.openDialog(DIALOG.field)}>{t('ins.field.insert')}</button>
        <button class="mi" role="menuitem" onclick={() => session.apply(commands.updateFieldsCommand, fieldContext(session))}>{t('ins.field.update')}</button>
        <button class="mi check" class:checked={session.showFieldCodes} role="menuitemcheckbox" aria-checked={session.showFieldCodes} onclick={() => { session.showFieldCodes = !session.showFieldCodes; session.openMenu = null; }}>{t('ins.field.toggle')}</button>
        <div class="menu-head">{t('ins.field.docProperty')}</div>
        {#each DOC_PROPERTIES as p (p)}
          <button class="mi" role="menuitem" onclick={() => session.apply(commands.insertFieldAtCaretCommand, { instruction: `DOCPROPERTY ${p} \\* MERGEFORMAT`, context: fieldContext(session) })}>{p}</button>
        {/each}
      {/snippet}
    </SplitButton>
    <Button size="mid" icon="dateTime" tip={t('ins.dateTime')} onclick={() => session.openDialog(DIALOG.dateTime)} disabled={!hasCaret} />
    <SplitButton id="insert.object" size="mid" icon="object" tip={t('ins.object')} disabled={!hasCaret}>
      {#snippet menu()}
        <label class="mi">{t('ins.object.textFromFile')}<input type="file" accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document" onchange={textFromFile} hidden /></label>
        <button class="mi" role="menuitem" onclick={() => session.openDialog(DIALOG.signatureLine)}>{t('ins.sig')}</button>
      {/snippet}
    </SplitButton>
  </div>
</Group>
<Group label={t('ins.group.symbols')}>
  <SplitButton id="insert.equation" size="large" icon="equation" tip={t('ins.equation')} onclick={newEquation} disabled={!hasCaret}>
    {#snippet menu()}
      <div class="menu-head">{t('ins.coverPage.builtIn')}</div>
      {#each equationPreviews as e (e.key)}
        <button class="mi" role="menuitem" style="flex-direction: column; align-items: flex-start" onclick={() => session.apply(commands.insertEquationCommand, { linear: e.linear })}>
          <span class="tile-name">{t(e.key)}</span>
          <!-- Library-rendered MathML (text escaped by renderMath). -->
          <span class="math-preview" style="border: none; padding: 0">{@html e.html}</span>
        </button>
      {/each}
      <hr />
      <button class="mi" role="menuitem" onclick={newEquation}><RibbonIcon name="equation" size={16} />{t('ins.eq.insertNew')}</button>
    {/snippet}
  </SplitButton>
  <Button size="large" icon="symbol" tip={t('ins.symbol')} onclick={() => session.openDialog(DIALOG.symbol)} disabled={!hasCaret} />
</Group>
