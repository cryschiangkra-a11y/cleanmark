<script lang="ts">
  import { i18n, setLang, t } from "../i18n/index.svelte";
</script>

<header>
  <div class="brand">
    <svg viewBox="0 0 32 32" width="32" height="32" aria-hidden="true">
      <rect width="32" height="32" rx="8" fill="var(--accent)" />
      <path d="M9 17l4.5 4.5L23 11" fill="none" stroke="var(--accent-text)" stroke-width="3" stroke-linecap="round" stroke-linejoin="round" />
    </svg>
    <div>
      <h1>CleanMark</h1>
      <p class="muted">{t("tagline")}</p>
    </div>
  </div>
  <button class="lang" onclick={() => setLang(i18n.lang === "en" ? "zh-Hant" : "en")}>{t("langSwitch")}</button>
</header>

<style>
  header {
    display: flex;
    justify-content: space-between;
    align-items: flex-start;
    gap: 12px;
    margin-bottom: 28px;
  }
  .brand {
    display: flex;
    gap: 12px;
    align-items: center;
  }
  h1 {
    margin: 0;
    font-size: 22px;
    line-height: 1.2;
  }
  p {
    margin: 2px 0 0;
    font-size: 14px;
  }
  .lang {
    white-space: nowrap;
    flex-shrink: 0;
    font-size: 14px;
    padding: 6px 12px;
  }
</style>
