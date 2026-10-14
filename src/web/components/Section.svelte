<script lang="ts">
  import type { Snippet } from "svelte";

  let { step, title, children }: { step: number; title: string; children: Snippet } = $props();
</script>

<section class="card">
  <h2><span class="num" aria-hidden="true">{step}</span>{title}</h2>
  {@render children()}
</section>

<style>
  section {
    margin-top: 16px;
  }
  h2 {
    display: flex;
    align-items: center;
    gap: 10px;
    font-size: 17px;
    margin: 0 0 14px;
  }
  .num {
    display: inline-flex;
    width: 24px;
    height: 24px;
    border-radius: 50%;
    align-items: center;
    justify-content: center;
    background: var(--accent-soft);
    color: var(--accent-soft-text);
    font-size: 13px;
  }
</style>
