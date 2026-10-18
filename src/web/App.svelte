<script lang="ts">
  import { untrack } from "svelte";
  import { applyChoices } from "../core/clean/apply-choices";
  import type { Report } from "../core/model";
  import ConfirmPanel from "./components/ConfirmPanel.svelte";
  import DownloadPanel from "./components/DownloadPanel.svelte";
  import DropZone from "./components/DropZone.svelte";
  import ErrorNotice from "./components/ErrorNotice.svelte";
  import FileBar from "./components/FileBar.svelte";
  import Preview from "./components/Preview.svelte";
  import ReportPanel from "./components/ReportPanel.svelte";
  import Section from "./components/Section.svelte";
  import SiteFooter from "./components/SiteFooter.svelte";
  import SiteHeader from "./components/SiteHeader.svelte";
  import { t } from "./i18n/index.svelte";
  import type { AnalyzeResult, BuildResult } from "./protocol";
  import { findingCount } from "./report-summary";
  import { CleanMarkWorker, WorkerError } from "./worker-client";

  type Phase = "idle" | "analyzing" | "ready" | "building" | "done" | "error";

  const worker = new CleanMarkWorker();
  let phase = $state<Phase>("idle");
  let fileName = $state("");
  let analysis = $state.raw<AnalyzeResult | null>(null);
  let removeIds = $state<string[]>([]);
  let keepImages = $state<number[]>([]);
  // 可疑字元：同形字同撇號預設換；特殊空格預設唔換（好多時係正常排版）
  let fixHomoglyphs = $state(true);
  let fixApostrophes = $state(true);
  let replaceSpecialSpaces = $state(false);
  const choices = $derived({ removeParagraphs: removeIds, keepImages, fixHomoglyphs, fixApostrophes, replaceSpecialSpaces });
  let built = $state.raw<BuildResult | null>(null);
  let error = $state<{ code: string; problems: { code: string; detail: string }[] } | null>(null);

  const report = $derived<Report | null>(analysis?.report ?? null);
  const previewDoc = $derived(analysis ? applyChoices(analysis.doc, choices) : null);

  function fail(e: unknown) {
    if (!(e instanceof WorkerError)) console.error(e);
    error = e instanceof WorkerError ? { code: e.code, problems: e.problems } : { code: "unknown", problems: [] };
    phase = "error";
  }

  async function onfile(file: File) {
    reset();
    fileName = file.name;
    phase = "analyzing";
    try {
      const result = await worker.analyze(file.name, await file.arrayBuffer());
      analysis = result;
      removeIds = []; // 可疑字句預設保留，由用家決定刪唔刪（例如 AI 使用聲明可能係必須保留嘅）
      // 圖片預設保留（會清走圖片入面嘅 metadata），排版先唔會變
      keepImages = result.report.images.filter((i) => i.keepable).map((i) => i.index);
      phase = "ready";
    } catch (e) {
      fail(e);
    }
  }

  // 用家改咗選擇，之前整好嘅檔案就作廢。只追蹤選擇，唔追蹤 phase
  $effect(() => {
    $state.snapshot(choices);
    untrack(() => {
      if (phase === "done") {
        built = null;
        phase = "ready";
      }
    });
  });

  async function onbuild() {
    phase = "building";
    try {
      // $state 係 Proxy，postMessage 複製唔到，要先轉做普通陣列
      built = await worker.build($state.snapshot(choices));
      phase = "done";
      download();
    } catch (e) {
      fail(e);
    }
  }

  function download() {
    if (!built) return;
    const url = URL.createObjectURL(
      new Blob([built.bytes as BlobPart], {
        type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      }),
    );
    const a = document.createElement("a");
    a.href = url;
    a.download = built.name;
    a.click();
    setTimeout(() => URL.revokeObjectURL(url), 10_000);
  }

  function reset() {
    phase = "idle";
    fileName = "";
    analysis = null;
    removeIds = [];
    keepImages = [];
    fixHomoglyphs = true;
    fixApostrophes = true;
    replaceSpecialSpaces = false;
    built = null;
    error = null;
  }
</script>

<main>
  <SiteHeader />

  {#if phase === "idle"}
    <DropZone {onfile} />
  {:else if phase === "analyzing"}
    <div class="card loading" role="status">{t("analyzing")}</div>
  {:else if phase === "error" && error}
    <ErrorNotice code={error.code} problems={error.problems} onreset={reset} />
  {/if}

  {#if report && previewDoc && (phase === "ready" || phase === "building" || phase === "done")}
    <FileBar name={fileName} count={findingCount(report)} onchange={reset} />
    <Section step={1} title={t("reportTitle")}>
      <ReportPanel {report} />
    </Section>
    <Section step={2} title={t("confirmTitle")}>
      <ConfirmPanel {report} bind:removeIds bind:keepImages bind:fixHomoglyphs bind:fixApostrophes bind:replaceSpecialSpaces />
    </Section>
    <Section step={3} title={t("previewTitle")}>
      <Preview doc={previewDoc} />
    </Section>
    <Section step={4} title={t("downloadTitle")}>
      <DownloadPanel phase={phase as "ready" | "building" | "done"} {onbuild} ondownload={download} onreset={reset} />
    </Section>
  {/if}

  <SiteFooter />
</main>

<style>
  main {
    max-width: 760px;
    margin: 0 auto;
    padding: 32px 16px 48px;
  }
  .loading {
    text-align: center;
    padding: 40px;
  }
</style>
