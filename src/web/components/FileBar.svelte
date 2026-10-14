<script lang="ts">
  import { t } from "../i18n/index.svelte";

  let { name, count, onchange }: { name: string; count: number; onchange: () => void } = $props();
</script>

<div class="bar card">
  <svg viewBox="0 0 24 24" width="22" height="22" aria-hidden="true">
    <path d="M14 3H7a2 2 0 00-2 2v14a2 2 0 002 2h10a2 2 0 002-2V8l-5-5zm0 0v5h5M9 13h6M9 17h6" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linejoin="round" />
  </svg>
  <span class="name">{name}</span>
  {#if count > 0}
    <span class="pill warn" data-testid="found-count">{t("foundN", { n: count })}</span>
  {:else}
    <span class="pill ok" data-testid="found-count">{t("foundNone")}</span>
  {/if}
  <button onclick={onchange}>{t("changeFile")}</button>
</div>

<style>
  .bar {
    display: flex;
    align-items: center;
    gap: 10px;
    padding: 12px 16px;
    flex-wrap: wrap;
  }
  .name {
    flex: 1;
    min-width: 0;
    overflow-wrap: anywhere;
    font-weight: 600;
  }
  button {
    font-size: 14px;
    padding: 4px 12px;
  }
</style>
