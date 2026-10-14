<script lang="ts">
  import type { Report } from "../../core/model";
  import { t, type MessageKey } from "../i18n/index.svelte";
  import { hiddenTotal } from "../report-summary";

  let { report }: { report: Report } = $props();
  const sum = (x: Record<string, number>) => Object.values(x).reduce((a, b) => a + b, 0);
</script>

<ul class="rows">
  <li>
    <details>
      <summary>
        <span class={report.visible.length ? "icon warn" : "icon ok"} aria-hidden="true">{report.visible.length ? "!" : "✓"}</span>
        <span class="label">{t("visibleTitle")}</span>
        <span class="muted">{report.visible.length || t("none")}</span>
      </summary>
      {#if report.visible.length}
        <ul class="detail">
          {#each report.visible as v, i (i)}
            <li>{t(`v_${v.code}` as MessageKey)}{v.text ? ` “${v.text}”` : ""} <code>{v.part}</code></li>
          {/each}
        </ul>
      {/if}
    </details>
  </li>
  <li>
    <details>
      <summary>
        <span class={report.metadata.length ? "icon warn" : "icon ok"} aria-hidden="true">{report.metadata.length ? "!" : "✓"}</span>
        <span class="label">{t("metadataTitle")}</span>
        <span class="muted">{report.metadata.length || t("none")}</span>
      </summary>
      {#if report.metadata.length}
        <ul class="detail">
          {#each report.metadata as m, i (i)}
            <li><code>{m.field}</code> {m.value}</li>
          {/each}
        </ul>
      {/if}
    </details>
  </li>
  <li>
    <details>
      <summary>
        <span class={report.c2pa.length ? "icon warn" : "icon ok"} aria-hidden="true">{report.c2pa.length ? "!" : "✓"}</span>
        <span class="label">{t("c2paTitle")}</span>
        <span class="muted">{report.c2pa.length || t("none")}</span>
      </summary>
      {#if report.c2pa.length}
        <ul class="detail">
          {#each report.c2pa as part (part)}
            <li><code>{part}</code></li>
          {/each}
        </ul>
      {/if}
    </details>
  </li>
  <li>
    <details>
      <summary>
        <span class={hiddenTotal(report) ? "icon warn" : "icon ok"} aria-hidden="true">{hiddenTotal(report) ? "!" : "✓"}</span>
        <span class="label">{t("hiddenTitle")}</span>
        <span class="muted">{hiddenTotal(report) || t("none")}</span>
      </summary>
      <ul class="detail">
        {#if sum(report.hidden.removed)}<li>{t("hiddenRemoved", { n: sum(report.hidden.removed) })} <code>{Object.keys(report.hidden.removed).join(", ")}</code></li>{/if}
        {#if sum(report.hidden.replaced)}<li>{t("hiddenReplaced", { n: sum(report.hidden.replaced) })}</li>{/if}
        {#if report.hidden.kept}<li>{t("hiddenKept", { n: report.hidden.kept })}</li>{/if}
      </ul>
    </details>
  </li>
  <li class="note muted">
    <span class="icon info" aria-hidden="true">i</span>{t("statisticalNote")}
  </li>
</ul>

<style>
  .rows {
    list-style: none;
    margin: 0;
    padding: 0;
  }
  .rows > li {
    border-top: 1px solid var(--border);
    padding: 10px 0;
  }
  .rows > li:first-child {
    border-top: none;
    padding-top: 0;
  }
  summary {
    display: flex;
    align-items: center;
    gap: 10px;
    cursor: pointer;
    list-style: none;
  }
  summary::-webkit-details-marker {
    display: none;
  }
  .label {
    flex: 1;
  }
  .icon {
    display: inline-flex;
    width: 20px;
    height: 20px;
    border-radius: 50%;
    align-items: center;
    justify-content: center;
    font-size: 12px;
    font-weight: 700;
    flex-shrink: 0;
  }
  .icon.warn {
    background: var(--warn-soft);
    color: var(--warn-text);
  }
  .icon.ok {
    background: var(--ok-soft);
    color: var(--ok-text);
  }
  .icon.info {
    background: var(--surface-muted);
    color: var(--text-muted);
  }
  .note {
    display: flex;
    gap: 10px;
    font-size: 14px;
  }
  .detail {
    margin: 8px 0 0 30px;
    padding: 0;
    list-style: none;
    font-size: 14px;
    overflow-wrap: anywhere;
  }
  .detail li {
    padding: 2px 0;
  }
  code {
    font-size: 12px;
    background: var(--surface-muted);
    border-radius: 4px;
    padding: 1px 5px;
  }
</style>
