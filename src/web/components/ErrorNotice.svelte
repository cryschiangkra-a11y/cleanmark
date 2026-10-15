<script lang="ts">
  import { t, type MessageKey } from "../i18n/index.svelte";

  let { code, problems, onreset }: { code: string; problems: { code: string; detail: string }[]; onreset: () => void } =
    $props();
  const KNOWN = ["legacyFormat", "unsupported", "tooLarge", "encrypted", "corrupt", "unsafeArchive", "selfCheckFailed"];
  const key = $derived((KNOWN.includes(code) ? `err_${code}` : "err_unknown") as MessageKey);
</script>

<div class="error card" role="alert" data-testid="error">
  <p>{t(key)}</p>
  {#if problems.length}
    <ul>
      {#each problems as p, i (i)}
        <li><code>{p.code}</code> {p.detail}</li>
      {/each}
    </ul>
  {/if}
  <button onclick={onreset}>{t("changeFile")}</button>
</div>

<style>
  .error {
    background: var(--danger-soft);
    color: var(--danger-text);
    border-color: transparent;
    margin-top: 16px;
  }
  p {
    margin: 0 0 12px;
  }
  ul {
    margin: 0 0 12px;
    font-size: 13px;
  }
</style>
