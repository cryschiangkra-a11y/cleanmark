<script lang="ts">
  import type { CmDocument } from "../../core/model";
  import { allBlocks } from "../../core/walk";
  import { listLabels } from "../list-label";
  import PreviewBlocks from "./PreviewBlocks.svelte";

  let { doc }: { doc: CmDocument } = $props();

  const labels = $derived(listLabels(allBlocks(doc), doc.lists));

  // 保留咗嘅原圖用 blob URL 顯示；換內容時釋放舊 URL
  let imageUrls = $state.raw(new Map<string, string>());
  $effect(() => {
    const urls = new Map<string, string>();
    for (const [path, bytes] of Object.entries(doc.media)) {
      urls.set(path, URL.createObjectURL(new Blob([bytes as BlobPart])));
    }
    imageUrls = urls;
    return () => urls.forEach((u) => URL.revokeObjectURL(u));
  });
</script>

<div class="frame" data-testid="preview">
  <div class="page">
    {#if doc.header.length}
      <div class="hf"><PreviewBlocks blocks={doc.header} {labels} {imageUrls} /></div>
    {/if}
    <PreviewBlocks blocks={doc.body} {labels} {imageUrls} />
    {#if doc.footer.length}
      <div class="hf footer"><PreviewBlocks blocks={doc.footer} {labels} {imageUrls} /></div>
    {/if}
  </div>
</div>

<style>
  .frame {
    max-height: 520px;
    overflow: auto;
    background: var(--surface-muted);
    border-radius: var(--radius);
    padding: 16px;
  }
  .page {
    background: #fff;
    color: #000;
    max-width: 640px;
    margin: 0 auto;
    padding: 32px 36px;
    border-radius: 4px;
    box-shadow: 0 1px 3px rgba(0, 0, 0, 0.12);
    font-family: "Times New Roman", "PMingLiU", serif;
    font-size: 11pt;
  }
  .hf {
    color: #666;
    border-bottom: 1px dashed #ddd;
    margin-bottom: 16px;
    font-size: 0.9em;
  }
  .hf.footer {
    border-bottom: none;
    border-top: 1px dashed #ddd;
    margin: 16px 0 0;
  }
  @media (max-width: 480px) {
    .frame {
      padding: 8px;
    }
    .page {
      padding: 16px;
    }
  }
</style>
