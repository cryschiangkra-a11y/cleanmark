<script lang="ts">
  import { t } from "../i18n/index.svelte";

  let { onfile }: { onfile: (file: File) => void } = $props();
  let input: HTMLInputElement;
  let dragging = $state(false);

  function pick(files: FileList | null | undefined) {
    const f = files?.[0];
    if (f) onfile(f);
  }
</script>

<div
  class="zone"
  class:dragging
  role="button"
  tabindex="0"
  aria-label={t("dropTitle")}
  onclick={() => input.click()}
  onkeydown={(e) => (e.key === "Enter" || e.key === " ") && (e.preventDefault(), input.click())}
  ondragover={(e) => {
    e.preventDefault();
    dragging = true;
  }}
  ondragleave={() => (dragging = false)}
  ondrop={(e) => {
    e.preventDefault();
    dragging = false;
    pick(e.dataTransfer?.files);
  }}
>
  <svg viewBox="0 0 24 24" width="36" height="36" aria-hidden="true">
    <path d="M12 16V4m0 0l-4 4m4-4l4 4M4 16v2a2 2 0 002 2h12a2 2 0 002-2v-2" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" />
  </svg>
  <p class="title">{t("dropTitle")}</p>
  <span class="choose">{t("dropButton")}</span>
  <p class="muted hint">{t("dropHint")}</p>
  <input
    bind:this={input}
    type="file"
    accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    hidden
    data-testid="file-input"
    onchange={(e) => pick(e.currentTarget.files)}
  />
</div>
<p class="privacy">
  <svg viewBox="0 0 24 24" width="16" height="16" aria-hidden="true">
    <path d="M12 3l7 3v5c0 4.5-3 8.5-7 10-4-1.5-7-5.5-7-10V6l7-3z" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linejoin="round" />
  </svg>
  {t("privacy")}
</p>

<ol class="steps">
  <li><strong>1 · {t("step1")}</strong><span class="muted">{t("step1Desc")}</span></li>
  <li><strong>2 · {t("step2")}</strong><span class="muted">{t("step2Desc")}</span></li>
  <li><strong>3 · {t("step3")}</strong><span class="muted">{t("step3Desc")}</span></li>
</ol>

<style>
  .zone {
    border: 2px dashed var(--border-strong);
    border-radius: var(--radius-lg);
    background: var(--surface);
    padding: 48px 20px;
    text-align: center;
    cursor: pointer;
    transition: border-color 0.15s, background 0.15s;
  }
  .zone:hover,
  .zone.dragging {
    border-color: var(--accent);
    background: var(--accent-soft);
  }
  .title {
    font-size: 18px;
    font-weight: 600;
    margin: 8px 0 12px;
  }
  .choose {
    display: inline-block;
    background: var(--accent);
    color: var(--accent-text);
    border-radius: var(--radius);
    padding: 8px 18px;
    font-weight: 600;
  }
  .hint {
    font-size: 13px;
    margin: 12px 0 0;
  }
  .privacy {
    display: flex;
    justify-content: center;
    align-items: center;
    gap: 6px;
    font-size: 14px;
    color: var(--accent-soft-text);
    margin: 14px 0 32px;
  }
  .steps {
    list-style: none;
    padding: 0;
    margin: 0;
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(160px, 1fr));
    gap: 12px;
  }
  .steps li {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    padding: 12px 14px;
    display: flex;
    flex-direction: column;
    font-size: 14px;
  }
</style>
