<script lang="ts">
  import { t } from "../i18n/index.svelte";
</script>

<footer>
  <h2>{t("limitsTitle")}</h2>
  <ul>
    <li>{t("limitStatistical")}</li>
    <li>{t("limitDropped")}</li>
  </ul>
  <p class="rule">{t("usageRule")}</p>
</footer>

<style>
  footer {
    margin-top: 40px;
    padding-top: 20px;
    border-top: 1px solid var(--border);
    color: var(--text-muted);
    font-size: 14px;
  }
  h2 {
    font-size: 15px;
    color: var(--text);
    margin: 0 0 6px;
  }
  ul {
    margin: 0 0 12px;
    padding-left: 20px;
  }
  .rule {
    margin: 0;
  }
</style>
