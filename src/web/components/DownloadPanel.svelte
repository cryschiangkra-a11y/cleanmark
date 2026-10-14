<script lang="ts">
  import { t } from "../i18n/index.svelte";

  let {
    phase,
    onbuild,
    ondownload,
    onreset,
  }: {
    phase: "ready" | "building" | "done";
    onbuild: () => void;
    ondownload: () => void;
    onreset: () => void;
  } = $props();
</script>

<div class="row">
  {#if phase === "ready"}
    <button class="primary" onclick={onbuild}>{t("buildButton")}</button>
  {:else if phase === "building"}
    <span class="muted" role="status">{t("building")}</span>
  {:else}
    <span class="pill ok" role="status" data-testid="self-check-ok">✓ {t("selfCheckPassed")}</span>
    <span class="spacer"></span>
    <button class="primary" onclick={ondownload}>{t("downloadAgain")}</button>
    <button onclick={onreset}>{t("another")}</button>
  {/if}
</div>

<style>
  .row {
    display: flex;
    gap: 10px;
    align-items: center;
    flex-wrap: wrap;
  }
  .spacer {
    flex: 1;
  }
</style>
