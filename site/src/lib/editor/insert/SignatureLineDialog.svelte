<script lang="ts">
  /**
   * Signature Setup: the signer details Word stores on `o:signatureline`, and
   * the picture of the line (an "X", a rule, the signer's name and title)
   * drawn here on a canvas, since a signature line shows as that picture
   * until it is signed.
   */
  import { commands } from '@office-kit/docx-editor';
  import Dialog from '../Dialog.svelte';
  import { getSession } from '../session.svelte';
  import { t } from '../i18n/index.svelte';
  import { DIALOG } from './state.svelte';

  const session = getSession();
  // Word's default signature line picture: 2 × 1 in, drawn at 2× for sharpness.
  const WIDTH_PT = 144;
  const HEIGHT_PT = 72;
  const SCALE = 2;

  let signer = $state('');
  let signerTitle = $state('');
  let email = $state('');
  let instructions = $state('');
  let allowComments = $state(false);
  let showDate = $state(true);

  $effect(() => {
    if (session.dialog !== DIALOG.signatureLine) return;
    signer = '';
    signerTitle = '';
    email = '';
    instructions = '';
    allowComments = false;
    showDate = true;
  });

  async function picture(): Promise<Uint8Array> {
    const canvas = document.createElement('canvas');
    canvas.width = WIDTH_PT * SCALE;
    canvas.height = HEIGHT_PT * SCALE;
    const ctx = canvas.getContext('2d');
    if (!ctx) throw new Error('This browser cannot draw the signature line picture.');
    ctx.scale(SCALE, SCALE);
    ctx.fillStyle = '#fff';
    ctx.fillRect(0, 0, WIDTH_PT, HEIGHT_PT);
    ctx.strokeStyle = '#000';
    ctx.fillStyle = '#000';
    ctx.lineWidth = 1;
    // The "X" and the rule sit at about 55% of the height, as Word draws them.
    const ruleY = HEIGHT_PT * 0.55;
    ctx.font = '20px Calibri, Arial, sans-serif';
    ctx.fillText('X', 6, ruleY - 4);
    ctx.beginPath();
    ctx.moveTo(4, ruleY);
    ctx.lineTo(WIDTH_PT - 4, ruleY);
    ctx.stroke();
    ctx.font = '9px Calibri, Arial, sans-serif';
    if (signer) ctx.fillText(signer, 6, ruleY + 12);
    if (signerTitle) ctx.fillText(signerTitle, 6, ruleY + 23);
    const blob = await new Promise<Blob | null>((resolve) => canvas.toBlob(resolve, 'image/png'));
    if (!blob) throw new Error('This browser cannot draw the signature line picture.');
    return new Uint8Array(await blob.arrayBuffer());
  }

  async function insert(): Promise<void> {
    try {
      const image = await picture();
      session.apply(commands.insertSignatureLineCommand, {
        image,
        widthPt: WIDTH_PT,
        heightPt: HEIGHT_PT,
        allowComments,
        showSignDate: showDate,
        ...(signer ? { suggestedSigner: signer } : {}),
        ...(signerTitle ? { suggestedSignerTitle: signerTitle } : {}),
        ...(email ? { suggestedSignerEmail: email } : {}),
        ...(instructions ? { instructions } : {}),
      });
    } catch (err) {
      session.status = err instanceof Error ? err.message : String(err);
      return;
    }
    if (session.status === '') session.dialog = null;
  }

  // The picture is drawn asynchronously, so OK closes the dialog itself once inserted.
  function ok(): boolean {
    void insert();
    return false;
  }
</script>

<Dialog id={DIALOG.signatureLine} title={t('ins.sig.title')} onok={ok}>
  <label class="stack">{t('ins.sig.signer')}<input type="text" bind:value={signer} /></label>
  <label class="stack">{t('ins.sig.signerTitle')}<input type="text" bind:value={signerTitle} /></label>
  <label class="stack">{t('ins.sig.email')}<input type="email" bind:value={email} /></label>
  <label class="stack">{t('ins.sig.instructions')}<textarea rows="2" bind:value={instructions}></textarea></label>
  <label><input type="checkbox" bind:checked={allowComments} /> {t('ins.sig.allowComments')}</label>
  <label><input type="checkbox" bind:checked={showDate} /> {t('ins.sig.showDate')}</label>
  {#if session.status}<p class="error-note">{session.status}</p>{/if}
</Dialog>
