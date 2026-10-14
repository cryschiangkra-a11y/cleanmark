<script lang="ts">
  import type { Report } from "../../core/model";
  import { t } from "../i18n/index.svelte";

  let {
    report,
    removeIds = $bindable(),
    keepImages = $bindable(),
  }: { report: Report; removeIds: string[]; keepImages: number[] } = $props();

  function toggleRemove(id: string, on: boolean) {
    removeIds = on ? [...removeIds, id] : removeIds.filter((x) => x !== id);
  }
  function setKeep(index: number, keep: boolean) {
    keepImages = keep ? [...new Set([...keepImages, index])] : keepImages.filter((x) => x !== index);
  }
  const keepable = $derived(report.images.filter((i) => i.keepable).map((i) => i.index));
  const tagNames = { exif: "EXIF", xmp: "XMP", c2pa: "C2PA" } as const;
</script>

{#if report.phrases.length === 0 && report.images.length === 0}
  <p class="muted">{t("nothingToConfirm")}</p>
{/if}

{#if report.phrases.length}
  <h3>{t("phrasesTitle")}</h3>
  <ul class="list">
    {#each report.phrases as ph (ph.paragraphId)}
      <li>
        <label>
          <input
            type="checkbox"
            checked={removeIds.includes(ph.paragraphId)}
            onchange={(e) => toggleRemove(ph.paragraphId, e.currentTarget.checked)}
          />
          <span>{t("phraseRemove")}</span>
        </label>
        <q>{ph.text}</q>
      </li>
    {/each}
  </ul>
{/if}

{#if report.images.length}
  <div class="images-head">
    <h3>{t("imagesTitle")}</h3>
    <div class="bulk">
      <button onclick={() => (keepImages = [])}>{t("allPlaceholder")}</button>
      <button onclick={() => (keepImages = [...keepable])}>{t("allKeep")}</button>
    </div>
  </div>
  <ul class="list">
    {#each report.images as img (img.index)}
      <li class="img">
        <div>
          <strong>{t("imageN", { n: img.index })}</strong>
          <span class="muted context">· {img.context ? t("imageAt", { text: img.context }) : t("imageAtStart")}</span>
          {#if img.tags.length}
            <span class="pill warn">{t("imageTagged", { tags: img.tags.map((x) => tagNames[x]).join(", ") })}</span>
          {/if}
        </div>
        {#if img.keepable}
          <div class="seg" role="radiogroup" aria-label={t("imageN", { n: img.index })}>
            <label class:on={!keepImages.includes(img.index)}>
              <input type="radio" name={`img-${img.index}`} checked={!keepImages.includes(img.index)} onchange={() => setKeep(img.index, false)} />
              {t("usePlaceholder")}
            </label>
            <label class:on={keepImages.includes(img.index)}>
              <input type="radio" name={`img-${img.index}`} checked={keepImages.includes(img.index)} onchange={() => setKeep(img.index, true)} />
              {t("keepOriginal")}
            </label>
          </div>
        {:else}
          <span class="muted small">{t("notKeepable")}</span>
        {/if}
      </li>
    {/each}
  </ul>
  {#if keepImages.length}
    <p class="info">{t("keepInfo")}</p>
  {/if}
{/if}

<style>
  h3 {
    font-size: 15px;
    margin: 0 0 8px;
  }
  .list {
    list-style: none;
    margin: 0 0 16px;
    padding: 0;
  }
  .list li {
    border-top: 1px solid var(--border);
    padding: 10px 0;
    display: flex;
    flex-direction: column;
    gap: 6px;
  }
  .list li.img {
    flex-direction: row;
    flex-wrap: wrap;
    justify-content: space-between;
    align-items: center;
  }
  label {
    display: inline-flex;
    gap: 6px;
    align-items: center;
    cursor: pointer;
  }
  q {
    color: var(--text-muted);
    font-size: 14px;
    overflow-wrap: anywhere;
  }
  .images-head {
    display: flex;
    justify-content: space-between;
    align-items: center;
    gap: 8px;
    flex-wrap: wrap;
    margin-bottom: 8px;
  }
  .images-head h3 {
    margin: 0;
  }
  .bulk {
    display: flex;
    gap: 6px;
  }
  .bulk button {
    font-size: 13px;
    padding: 4px 10px;
  }
  .seg {
    display: inline-flex;
    border: 1px solid var(--border-strong);
    border-radius: var(--radius);
    overflow: hidden;
  }
  .seg label {
    padding: 4px 12px;
    font-size: 14px;
  }
  .seg label.on {
    background: var(--accent-soft);
    color: var(--accent-soft-text);
    font-weight: 600;
  }
  .seg input {
    position: absolute;
    opacity: 0;
    pointer-events: none;
  }
  .context {
    overflow-wrap: anywhere;
  }
  .small {
    font-size: 13px;
  }
  .info {
    background: var(--ok-soft);
    color: var(--ok-text);
    border-radius: var(--radius);
    padding: 8px 12px;
    font-size: 14px;
    margin: 0;
  }
</style>
