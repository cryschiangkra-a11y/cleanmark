import { readFileSync } from "node:fs";
import { expect, test } from "@playwright/test";
import { readDocx } from "../src/core/readers/docx";
import { selfCheckDocx } from "../src/core/verify/self-check-docx";
import { allImages, allParagraphs, paragraphText } from "../src/core/walk";

test.use({ locale: "zh-TW" });

test("拖入 Word → 報告 → 改選擇 → 下載乾淨檔案", async ({ page }) => {
  await page.goto("/");
  await expect(page.getByText("把 Word 檔案拖到這裡")).toBeVisible();

  await page.getByTestId("file-input").setInputFiles("e2e/.fixtures/sample.docx");
  await expect(page.getByTestId("found-count")).toHaveText("找到 6 項");

  // 預設：可疑字句保留、圖片保留（清走圖片資料）
  const preview = page.getByTestId("preview");
  await expect(preview).toContainText("本文由 AI");
  await expect(preview.locator("img")).toHaveCount(1);
  await expect(page.getByText("保留的圖片會清除 EXIF")).toBeVisible();

  await page.getByLabel("刪除這段").check();
  await expect(preview).not.toContainText("本文由 AI");
  await page.getByText("佔位框", { exact: true }).click(); // 單選掣本身係隱藏嘅，撳文字標籤
  await expect(preview).toContainText("圖片 1：請重新插入");

  const downloadPromise = page.waitForEvent("download");
  await page.getByRole("button", { name: "清除並下載" }).click();
  const download = await downloadPromise;
  expect(download.suggestedFilename()).toBe("sample-clean.docx");
  await expect(page.getByTestId("self-check-ok")).toBeVisible();

  const bytes = new Uint8Array(readFileSync(await download.path()));
  expect(await selfCheckDocx(bytes, 1)).toEqual({ ok: true, problems: [] });
  const { doc } = await readDocx(bytes);
  // 頁首：文字保留；水印所在嘅段落保留但已清空（同原檔排版一致）
  expect(allParagraphs(doc).map(paragraphText)).toEqual(["第一章", "營業額上升 12%", "季度報告", ""]);
  expect(allImages(doc)).toHaveLength(1); // 佔位圖
});

test("唔支援嘅格式會顯示錯誤", async ({ page }) => {
  await page.goto("/");
  await page.getByTestId("file-input").setInputFiles({ name: "old.doc", mimeType: "application/msword", buffer: Buffer.from("x") });
  await expect(page.getByTestId("error")).toContainText("另存為 .docx");
});

test("網頁有 CSP，禁止連去外部網址", async ({ page }) => {
  await page.goto("/");
  const csp = await page.locator('meta[http-equiv="Content-Security-Policy"]').getAttribute("content");
  expect(csp).toContain("connect-src 'self'");
});

test("可以切換英文", async ({ page }) => {
  await page.goto("/");
  await page.getByRole("button", { name: "English" }).click();
  await expect(page.getByText("Drop a Word file here")).toBeVisible();
});
