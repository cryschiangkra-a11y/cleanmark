// 用 Word 驗證排版：原檔同 CleanMark 新檔各自轉 PDF，逐頁比較畫面
// 只喺有 Microsoft Word 嘅 Mac 用得；用法：npx tsx scripts/verify-layout.mts <原檔.docx>
import { execFileSync } from "node:child_process";
import JSZip from "jszip";
import { mkdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";
import { analyze, build } from "../src/core/pipeline";
import { OoxmlPackage } from "../src/core/ooxml/package";
import { writeCleanDocx } from "../src/core/sanitize/package";
import { xmlDiff } from "./xml-diff.mts";

const src = process.argv.slice(2).find((a) => !a.startsWith("--"));
const strict = process.argv.includes("--strict");
if (!src) {
  console.error("用法：npx tsx scripts/verify-layout.mts <原檔.docx>");
  process.exit(2);
}
// Word 有沙盒，放喺佢自己嘅 folder 就唔會彈權限視窗
const dir = join(homedir(), "Library/Containers/com.microsoft.Word/Data/cleanmark-verify");
rmSync(dir, { recursive: true, force: true });
mkdirSync(dir, { recursive: true });

const original = join(dir, "original.docx");
const cleaned = join(dir, "cleaned.docx");
const bytes = new Uint8Array(readFileSync(src));
writeFileSync(original, await forWord(bytes));
const { report } = await analyze("original.docx", bytes);
const keepAll = report.images.filter((i) => i.keepable).map((i) => i.index);
// 用正式流程（包括自我檢查）；圖片全部保留，比較排版
const out = await build("original.docx", bytes, { removeParagraphs: [], keepImages: keepAll });
writeFileSync(cleaned, await forWord(out.bytes));

// 檔案層面：被拎走嘅嘢，同埋其餘部分係咪逐個元素一樣
const { dropped } = await writeCleanDocx(await OoxmlPackage.open(bytes), { removeParagraphs: [], keepImages: keepAll });
console.log("── 檔案層面 ──");
console.log("拎走或者改寫咗：");
for (const [k, n] of Object.entries(dropped).sort()) console.log(`  ${k} × ${n}`);
for (const line of await xmlDiff(bytes, out.bytes)) console.log(line);
console.log("── 畫面同排版 ──");

/** 驗證用副本：拎走「打開時更新欄位」，否則 Word 會彈視窗問，而且更新後兩份就唔係比較原本嘅樣 */
async function forWord(docx: Uint8Array): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(docx);
  const settings = await zip.file("word/settings.xml")?.async("string");
  if (settings) zip.file("word/settings.xml", settings.replace(/<w:updateFields\b[^>]*\/>/g, ""));
  return zip.generateAsync({ type: "uint8array" });
}

function toPdf(docx: string): string {
  const pdf = docx.replace(/\.docx$/, ".pdf");
  execFileSync("open", ["-a", "Microsoft Word", docx]);
  const script = `
    tell application "Microsoft Word"
      repeat 60 times
        if (count of documents) > 0 then exit repeat
        delay 0.5
      end repeat
      set d to document 1
      save as d file name ((POSIX file "${pdf}") as text) file format format PDF
      close d saving no
    end tell`;
  execFileSync("osascript", ["-e", script]);
  return pdf;
}

const pdfA = toPdf(original);
const pdfB = toPdf(cleaned);
mkdirSync(join(dir, "diff"), { recursive: true });
const result = execFileSync(
  "swift",
  [join(import.meta.dirname, "pdf-diff.swift"), pdfA, pdfB, join(dir, "diff"), ...(strict ? ["--strict"] : [])],
  { encoding: "utf8", maxBuffer: 64 * 1024 * 1024 },
);
console.log(result.trim());
console.log(`檔案喺：${dir}`);
