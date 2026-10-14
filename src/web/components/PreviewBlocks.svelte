<script lang="ts">
  import type { Block } from "../../core/model";
  import { t } from "../i18n/index.svelte";
  import { mathPreviewText } from "../math-text";
  import { EMU_TO_PX, paragraphCss, runCss } from "../preview-style";
  import PreviewBlocks from "./PreviewBlocks.svelte";

  let {
    blocks,
    labels,
    imageUrls,
  }: { blocks: Block[]; labels: Map<string, string>; imageUrls: Map<string, string> } = $props();
</script>

{#each blocks as b, i (i)}
  {#if b.type === "paragraph"}
    <div
      class="p"
      class:heading={b.heading > 0}
      role={b.heading > 0 ? "heading" : undefined}
      aria-level={b.heading > 0 ? b.heading : undefined}
      style={paragraphCss(b)}
    >
      {#if labels.has(b.id)}<span class="label">{labels.get(b.id)}</span>{/if}
      {#each b.inlines as inl, j (j)}
        {#if inl.type === "text"}
          {#if inl.link}
            <a href={inl.link} target="_blank" rel="noopener noreferrer" style={runCss(inl.style)}>{inl.text}</a>
          {:else}
            <span style={runCss(inl.style)}>{inl.text}</span>
          {/if}
        {:else if inl.type === "break"}
          <br />
        {:else if inl.type === "math"}
          <span class="math" class:display={inl.display}>{mathPreviewText(inl.root)}</span>
        {:else}
          <sup>[{inl.id}]</sup>
        {/if}
      {/each}
      {#if b.inlines.length === 0}&nbsp;{/if}
    </div>
  {:else if b.type === "table"}
    <table class:borders={b.borders}>
      <tbody>
        {#each b.rows as row, r (r)}
          <tr>
            {#each row as cell, c (c)}
              <td colspan={cell.colSpan} rowspan={cell.rowSpan}>
                <PreviewBlocks blocks={cell.blocks} {labels} {imageUrls} />
              </td>
            {/each}
          </tr>
        {/each}
      </tbody>
    </table>
  {:else if b.type === "image"}
    {@const w = Math.max(60, Math.round(b.widthEmu * EMU_TO_PX))}
    {@const h = Math.max(30, Math.round(b.heightEmu * EMU_TO_PX))}
    {#if b.keep && imageUrls.has(b.mediaPath)}
      <img src={imageUrls.get(b.mediaPath)} alt={t("imageN", { n: b.index })} style={`width:${w}px;max-width:100%;height:auto`} />
    {:else}
      <div class="placeholder" style={`width:${w}px;height:${h}px`}>{t("placeholderText", { n: b.index })}</div>
    {/if}
  {:else}
    <hr class="page-break" />
  {/if}
{/each}

<style>
  .p {
    white-space: pre-wrap;
    overflow-wrap: anywhere;
  }
  .label {
    display: inline-block;
    min-width: 1.6em;
    margin-right: 0.4em;
  }
  table {
    border-collapse: collapse;
    margin: 6px 0;
    max-width: 100%;
  }
  td {
    padding: 4px 6px;
    vertical-align: top;
    border: 1px dotted #ddd;
  }
  table.borders td {
    border: 1px solid #000;
  }
  .math {
    font-family: "Cambria Math", "STIX Two Math", "Times New Roman", serif;
    font-style: italic;
  }
  .math.display {
    display: block;
    text-align: center;
    margin: 4px 0;
  }
  .placeholder {
    max-width: 100%;
    border: 2px dashed #999;
    display: flex;
    align-items: center;
    justify-content: center;
    color: #777;
    font-size: 12px;
    text-align: center;
    margin: 6px 0;
  }
  .page-break {
    border: none;
    border-top: 2px dashed #ccc;
    margin: 16px 0;
  }
  a {
    color: inherit;
  }
</style>
