<script lang="ts">
  /** Word's status bar: counts on the left, zoom on the right. */
  import { getSession, MIN_ZOOM, MAX_ZOOM } from './session.svelte';
  import { t } from './i18n/index.svelte';

  const session = getSession();
</script>

<footer class="statusbar">
  <span>{t('status.pageOf').replace('{0}', String(session.currentPage)).replace('{1}', String(session.pageCount))}</span>
  <span title={`${session.charCount} ${t('status.chars')}`}>{session.wordCount} {t('status.words')}</span>
  {#if session.status}<span class="status-msg" role="status">{session.status}</span>{/if}
  <div class="spacer"></div>
  <div class="zoom">
    <button onclick={() => session.setZoom(session.zoom - 0.1)} aria-label={t('view.zoomOut')}>−</button>
    <input
      type="range"
      min={MIN_ZOOM * 100}
      max={MAX_ZOOM * 100}
      step="10"
      value={Math.round(session.zoom * 100)}
      oninput={(e) => session.setZoom(Number((e.currentTarget as HTMLInputElement).value) / 100)}
      aria-label="Zoom"
    />
    <button onclick={() => session.setZoom(session.zoom + 0.1)} aria-label={t('view.zoomIn')}>+</button>
    <span class="zoom-val">{Math.round(session.zoom * 100)}%</span>
  </div>
</footer>

<style>
  .statusbar {
    display: flex;
    align-items: center;
    gap: 16px;
    height: 26px;
    padding: 0 14px;
    border-top: 1px solid var(--status-line);
    background: var(--status-bg);
    font-size: 13px;
    color: var(--text);
  }
  .status-msg { color: #a4262c; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
  .spacer { flex: 1; }
  .zoom { display: flex; align-items: center; gap: 6px; }
  .zoom button { width: 18px; height: 18px; border: none; background: none; cursor: pointer; font-size: 14px; line-height: 1; border-radius: 3px; color: inherit; }
  .zoom button:hover { background: var(--hover); }
  .zoom input[type='range'] { width: 130px; accent-color: #8a8a8a; }
  .zoom-val { min-width: 40px; text-align: right; }
</style>
