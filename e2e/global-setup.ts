import { mkdirSync, writeFileSync } from "node:fs";
import { buildDocx, imageP, p, pngWith, wordTextWatermark } from "../tests/fixtures/docx-builder";

// 產生 e2e 用嘅樣本檔：水印、metadata、C2PA 圖片、可疑字句、隱形字元
export default async function globalSetup(): Promise<void> {
  mkdirSync("e2e/.fixtures", { recursive: true });
  const bytes = await buildDocx({
    header: p("季度報告") + wordTextWatermark("AI Generated"),
    core: { "dc:creator": "ChatGPT" },
    app: { Application: "AI Writer" },
    media: [{ rId: "rIdA", name: "chart.png", bytes: pngWith("c2pa") }],
    body: p("第一章") + p("Clаude summary") + p("營業額\u200B上升 12%") + imageP("rIdA") + p("本文由 AI 協助生成。"),
  });
  writeFileSync("e2e/.fixtures/sample.docx", bytes);
}
