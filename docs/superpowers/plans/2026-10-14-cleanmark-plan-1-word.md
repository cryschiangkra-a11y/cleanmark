# CleanMark 第 1 期實作計劃：基礎 + Word 全套 + 網頁 + GitHub Pages

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 做一個純瀏覽器網頁：用家拖入 .docx，睇到水印報告、確認可疑字句同圖片處理方式、預覽，然後下載一份重新砌過、通過自我檢查嘅乾淨 .docx；並自動發佈去 GitHub Pages。

**Architecture:** `src/core/` 係唔依賴瀏覽器嘅核心：讀取器將 .docx 轉成「內容清單」(`CmDocument`)，檢查器出報告，`applyChoices` 按用家選擇清潔，輸出器由空白砌新 .docx，自我檢查再驗一次。`src/web/` 係 Svelte 單頁介面，核心喺 Web Worker 入面行。第二期 Terminal 版會直接重用 `src/core/`。

**Tech Stack:** TypeScript 6.0.3、Vite 8.3.4、Svelte 5.57.2、Vitest 5.0.4、Playwright 1.64.1、JSZip 3.10.2、@xmldom/xmldom 0.9.12、docx 9.11.0

**設計文件：** [docs/superpowers/specs/2026-10-14-cleanmark-design.md](../specs/2026-10-14-cleanmark-design.md)

## 成個計劃點分期

| 計劃 | 內容 | 做完之後 |
|---|---|---|
| **1（呢份）** | 基礎、Word 讀寫、檢查器、自我檢查、網頁、GitHub Pages | 可以 send link 俾人處理 Word 檔 |
| 2 | PowerPoint 讀寫同檢查 | 支援 .pptx |
| 3 | PDF 讀取同檢查 | PDF → .docx |
| 4 | PDF 輸出（可選）、Noto 字型、語言偵測 | 「同時輸出 PDF」選項 |

## Global Constraints

- 所有依賴用**準確版本號**（見 `package.json`），唔好用 `^`。TypeScript 用 6.0.3，因為 `svelte-check` 4.7.6 未支援 TypeScript 7
- Node 版本：本機 26，CI 用 24（Vitest 5 要求 `^22.12 || ^24 || >=26`）
- `src/core/` 唔准用 DOM、`window`、`document`；XML 一律用 `@xmldom/xmldom`
- 新檔案 metadata 一律中性：作者、標題、公司等留空，只准有 `cp:revision`、`dcterms:created`、`dcterms:modified`
- 白名單原則：只有 `src/core/model.ts` 定義咗嘅欄位可以搬去新檔案
- 檔案上限 50 MB
- 網頁正式版要有 CSP：`connect-src 'self'`，所有資源由自己網站提供，唔准用外部 CDN、字型或者分析工具
- 介面語言：繁中（書面語，方便台灣同其他地區用家）同 English，key 兩邊一定要一樣
- 圖片預設全部用佔位框；可疑字句預設剔咗「刪除」
- **Git：由用家自己 commit。** 每個 task 最後嘅「Commit」步驟係檢查點：執行者停低，話俾用家知可以 commit 咩檔案，等用家做完先繼續
- 呢個 folder 未係 git repo，第一次 commit 之前用家要先行 `git init`

## 同設計文件嘅差異（第 1 期）

- 檔案資料行暫時只顯示檔名同「找到 N 項」；**頁數同偵測到嘅語言**留到第 4 期（要配合 PDF 字型嘅文字偵測一齊做）
- 只接受 .docx；.pptx 同 .pdf 會顯示「即將推出」（錯誤代碼 `comingSoon`）
- 「同時輸出 PDF」選項第 4 期先加
- 只讀 Word 嘅預設頁首頁尾（`w:type="default"`）；「首頁不同」、「奇偶頁不同」嘅頁首頁尾唔搬

## 檔案結構

```
removewatermark/
├── package.json, tsconfig.json, vite.config.ts, svelte.config.js, playwright.config.ts
├── index.html
├── public/favicon.svg
├── src/core/
│   ├── model.ts                    內容清單、報告、用家選擇嘅型別
│   ├── walk.ts                     行勻成份文件嘅區塊、段落、圖片
│   ├── pipeline.ts                 analyze()、build()、輸入檢查、錯誤代碼
│   ├── ooxml/xml.ts                XML 解析同命名空間工具
│   ├── ooxml/package.ts            讀 zip 同 relationships
│   ├── readers/docx/styles.ts      樣式繼承
│   ├── readers/docx/numbering.ts   列表定義
│   ├── readers/docx/body.ts        段落、文字、表格、圖片、文字方塊
│   ├── readers/docx/index.ts       readDocx()：頁面設定、頁首頁尾、註腳
│   ├── inspect/bytes.ts            EXIF / XMP / C2PA 標記搜尋
│   ├── inspect/phrases.ts          可疑字句
│   ├── inspect/metadata.ts         docProps、customXml
│   ├── inspect/visible-docx.ts     可見水印
│   ├── inspect/docx.ts             inspectDocx()：合併成報告
│   ├── clean/hidden-chars.ts       隱形字元規則
│   ├── clean/apply-choices.ts      按用家選擇產生最終內容清單
│   ├── writers/docx.ts             由空白砌新 .docx
│   └── verify/self-check-docx.ts   新檔案自我檢查
├── src/web/
│   ├── main.ts, App.svelte, styles.css
│   ├── protocol.ts, worker.ts, worker-client.ts
│   ├── i18n/zh-Hant.ts, i18n/en.ts, i18n/index.svelte.ts
│   ├── list-label.ts, preview-style.ts, report-summary.ts
│   └── components/                 SiteHeader、DropZone、FileBar、Section、ReportPanel、
│                                   ConfirmPanel、Preview、PreviewBlocks、DownloadPanel、
│                                   ErrorNotice、SiteFooter
├── tests/                          Vitest 單元測試
│   └── fixtures/docx-builder.ts    用純 XML 砌測試用 .docx
├── e2e/                            Playwright 瀏覽器測試
├── .github/workflows/ci.yml
├── README.md, LICENSE
```

---

### Task 1：專案骨架

**Files:**
- Create: `package.json`, `tsconfig.json`, `vite.config.ts`, `svelte.config.js`, `.gitignore`
- Test: `tests/smoke.test.ts`

**Interfaces:**
- Produces: `npm test`（Vitest）、`npm run check`（svelte-check）、`npm run build`、`npm run e2e` 指令

- [ ] **Step 1：建立 `package.json`**

`package.json`

```json
{
  "name": "cleanmark",
  "version": "0.1.0",
  "private": true,
  "type": "module",
  "scripts": {
    "dev": "vite",
    "build": "vite build",
    "preview": "vite preview --port 4173 --strictPort",
    "check": "svelte-check --tsconfig ./tsconfig.json",
    "test": "vitest run",
    "e2e": "playwright test"
  },
  "dependencies": {
    "@xmldom/xmldom": "0.9.12",
    "docx": "9.11.0",
    "jszip": "3.10.2"
  },
  "devDependencies": {
    "@playwright/test": "1.64.1",
    "@sveltejs/vite-plugin-svelte": "7.3.1",
    "@tsconfig/svelte": "5.0.8",
    "svelte": "5.57.2",
    "svelte-check": "4.7.6",
    "typescript": "6.0.3",
    "vite": "8.3.4",
    "vitest": "5.0.4"
  }
}
```

- [ ] **Step 2：安裝依賴**

Run: `npm install`
Expected: `found 0 vulnerabilities`，產生 `package-lock.json`

- [ ] **Step 3：建立 `tsconfig.json`**

`tsconfig.json`

```json
{
  "extends": "@tsconfig/svelte/tsconfig.json",
  "compilerOptions": {
    "target": "ES2022",
    "module": "ESNext",
    "moduleResolution": "bundler",
    "lib": ["ES2022", "DOM", "DOM.Iterable", "WebWorker"],
    "types": ["vite/client"],
    "strict": true,
    "noEmit": true,
    "skipLibCheck": true,
    "verbatimModuleSyntax": true,
    "isolatedModules": true,
    "resolveJsonModule": true
  },
  "include": ["src/**/*.ts", "src/**/*.svelte", "tests/**/*.ts", "e2e/**/*.ts", "vite.config.ts", "playwright.config.ts"]
}
```

- [ ] **Step 4：建立 `vite.config.ts`**（正式打包先加 CSP；Vitest 設定都喺度）

`vite.config.ts`

```ts
/// <reference types="vitest/config" />
import { defineConfig, type Plugin } from "vite";
import { svelte } from "@sveltejs/vite-plugin-svelte";

// 只喺正式打包時加入 CSP；開發模式要用 inline script 同 websocket，加咗會壞
const CSP = [
  "default-src 'self'",
  "script-src 'self'",
  "style-src 'self' 'unsafe-inline'",
  "img-src 'self' data: blob:",
  "font-src 'self'",
  "connect-src 'self'",
  "worker-src 'self' blob:",
  "object-src 'none'",
  "base-uri 'none'",
  "form-action 'none'",
].join("; ");

function cspPlugin(): Plugin {
  return {
    name: "cleanmark-csp",
    apply: "build",
    transformIndexHtml(html) {
      return html.replace(
        "<head>",
        `<head>\n    <meta http-equiv="Content-Security-Policy" content="${CSP}" />`,
      );
    },
  };
}

export default defineConfig({
  base: "./",
  plugins: [svelte(), cspPlugin()],
  worker: { format: "es" },
  test: {
    include: ["tests/**/*.test.ts"],
    environment: "node",
  },
});
```

- [ ] **Step 5：建立 `svelte.config.js` 同 `.gitignore`**

`svelte.config.js`

```js
import { vitePreprocess } from "@sveltejs/vite-plugin-svelte";

export default {
  preprocess: vitePreprocess(),
};
```

`.gitignore`

```gitignore
node_modules/
dist/
test-results/
playwright-report/
e2e/.fixtures/
```

- [ ] **Step 6：寫煙霧測試**

`tests/smoke.test.ts`

```ts
import { expect, test } from "vitest";

test("toolchain works", () => {
  expect(1 + 1).toBe(2);
});
```

- [ ] **Step 7：行測試**

Run: `npm test`
Expected: `Tests  1 passed (1)`

- [ ] **Step 8：Commit（由用家做）**

```bash
git init
git add package.json package-lock.json tsconfig.json vite.config.ts svelte.config.js .gitignore tests/smoke.test.ts
git commit -m "chore: scaffold CleanMark project"
```

---

### Task 2：內容清單型別

**Files:**
- Create: `src/core/model.ts`

**Interfaces:**
- Produces: `CmDocument`、`Block`、`Paragraph`、`Table`、`TableCell`、`ImageSlot`、`PageBreak`、`Inline`、`TextInline`、`RunStyle`、`ListRef`、`ListLevel`、`PageSetup`、`Report`、`VisibleItem`、`VisibleCode`、`MetadataItem`、`PhraseItem`、`ImageItem`、`HiddenCharSummary`、`Choices`、`BuildOptions`、`KEEPABLE_IMAGE_TYPES`

呢個 task 只有型別，冇行為可以測；型別會喺 Task 3 之後每個 task 嘅測試入面用到，`npm run check` 會驗證。

- [ ] **Step 1：建立 `src/core/model.ts`**

`src/core/model.ts`

```ts
// 「內容清單」：讀取器嘅輸出、輸出器嘅輸入。
// 只有呢度定義咗嘅欄位會被搬去新檔案（白名單），水印同 metadata 冇欄位可以裝。

export type Align = "left" | "center" | "right" | "justify";

export interface RunStyle {
  font?: string; // 西文字型
  fontEastAsia?: string; // 中日韓字型
  size: number; // pt
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strike: boolean;
  color?: string; // RRGGBB，undefined = 自動（黑色）
  highlight?: string; // OOXML highlight 名稱，例如 "yellow"
  vertAlign?: "superscript" | "subscript";
}

export interface TextInline {
  type: "text";
  text: string; // 可以包含 "\t"
  style: RunStyle;
  link?: string; // 外部超連結網址
}

export interface BreakInline {
  type: "break"; // 行內換行
}

export interface FootnoteRefInline {
  type: "footnoteRef";
  id: number;
}

export type Inline = TextInline | BreakInline | FootnoteRefInline;

export interface ListRef {
  numId: string;
  level: number; // 0 起
}

export interface LineSpacing {
  line: number; // twip；rule 係 auto 時 240 = 單行
  rule: "auto" | "exact" | "atLeast";
}

export interface Paragraph {
  type: "paragraph";
  id: string; // 例如 "p12"，俾報告同用家選擇參照
  heading: number; // 0 = 內文，1–6 = 標題層級
  align: Align;
  indentLeft: number; // twip
  indentFirstLine: number; // twip，負數 = 懸掛縮排
  spacingBefore: number; // twip
  spacingAfter: number; // twip
  lineSpacing: LineSpacing;
  list?: ListRef;
  inlines: Inline[];
}

export interface TableCell {
  blocks: Block[];
  colSpan: number;
  rowSpan: number; // 垂直合併嘅起點格；被合併嘅格唔會出現
}

export interface Table {
  type: "table";
  columnWidths: number[]; // twip
  rows: TableCell[][];
  borders: boolean;
}

export interface ImageSlot {
  type: "image";
  index: number; // 1 起，全文件順序
  widthEmu: number;
  heightEmu: number;
  page: number; // 估計頁數，1 起
  mediaPath: string; // 原檔入面嘅路徑，例如 "word/media/image1.png"
  contentType: "png" | "jpg" | "gif" | "bmp" | "other";
  keep: boolean; // true = 保留原圖；false = 佔位框（由 applyChoices 設定）
}

/** 可以原圖放入新檔案嘅圖片格式 */
export const KEEPABLE_IMAGE_TYPES: ReadonlySet<ImageSlot["contentType"]> = new Set(["png", "jpg", "gif", "bmp"]);

export interface PageBreak {
  type: "pageBreak";
}

export type Block = Paragraph | Table | ImageSlot | PageBreak;

export interface ListLevel {
  format: string; // OOXML numFmt，例如 "decimal"、"bullet"、"lowerLetter"
  text: string; // OOXML lvlText，例如 "%1."、"•"
  start: number;
  indentLeft: number; // twip
  hanging: number; // twip
}

export interface PageSetup {
  widthTwip: number;
  heightTwip: number;
  landscape: boolean;
  margins: { top: number; right: number; bottom: number; left: number };
}

export interface CmDocument {
  kind: "docx";
  page: PageSetup;
  body: Block[];
  header: Block[]; // 只有段落同表格
  footer: Block[];
  footnotes: Record<number, Block[]>;
  lists: Record<string, ListLevel[]>; // numId → 每層設定
  media: Record<string, Uint8Array>; // mediaPath → 原圖 bytes（只喺保留原圖時用）
}

// ── 報告 ──────────────────────────────────────────────

export type VisibleCode =
  | "wordTextWatermark" // Word 內建文字水印
  | "wordPictureWatermark" // Word 內建圖片水印
  | "headerShape" // 頁首頁尾嘅圖形或者文字方塊
  | "headerImage" // 頁首頁尾嘅圖片
  | "behindText" // 擺喺文字後面嘅圖形
  | "pageBackground" // 頁面背景
  | "textArt"; // 內文入面嘅 WordArt

export interface VisibleItem {
  code: VisibleCode;
  part: string;
  text?: string;
}

export interface MetadataItem {
  part: string;
  field: string;
  value: string;
}

export interface PhraseItem {
  paragraphId: string;
  text: string; // 段落文字（最多 160 字）
  match: string;
}

export interface ImageItem {
  index: number;
  page: number;
  keepable: boolean; // 原圖格式可唔可以放入新檔案
  tags: ("exif" | "xmp" | "c2pa")[];
}

export interface HiddenCharSummary {
  removed: Record<string, number>; // 例如 { "U+200B": 23 }
  replaced: Record<string, number>; // 換成普通空格
  kept: number; // 因為語言需要而保留
}

export interface Report {
  fileKind: "docx";
  visible: VisibleItem[];
  metadata: MetadataItem[];
  c2pa: string[]; // 搵到 C2PA 嘅 part
  hidden: HiddenCharSummary;
  phrases: PhraseItem[];
  images: ImageItem[];
}

// ── 用家選擇 ─────────────────────────────────────────

export interface Choices {
  removeParagraphs: string[]; // 要刪除嘅段落 id
  keepImages: number[]; // 保留原圖嘅圖片序號，其餘用佔位框
}

export interface BuildOptions {
  placeholderText: string; // 例如 "圖片 {n}：請重新插入"
}
```

- [ ] **Step 2：型別檢查**

Run: `npx tsc --noEmit -p .`
Expected: 冇輸出（0 錯誤）

- [ ] **Step 3：Commit（由用家做）**

```bash
git add src/core/model.ts
git commit -m "feat(core): add document model types"
```

---

### Task 3：隱形字元清潔器

**Files:**
- Create: `src/core/clean/hidden-chars.ts`
- Test: `tests/hidden-chars.test.ts`

**Interfaces:**
- Consumes: `HiddenCharSummary`（Task 2）
- Produces:
  - `cleanText(text: string, ctx?: CleanContext): { text: string; summary: HiddenCharSummary }`
  - `interface CleanContext { before: string; after: string; rtl: boolean }`
  - `emptySummary(): HiddenCharSummary`
  - `mergeSummary(a, b): HiddenCharSummary`
  - `hasRtl(text: string): boolean`

規則（設計文件 5.6）：一定清走嘅、換成空格嘅、ZWJ/ZWNJ 睇前後文字、方向標記睇段落有冇由右至左文字、異體選擇符睇前一個字。

- [ ] **Step 1：寫失敗測試**

`tests/hidden-chars.test.ts`

```ts
import { describe, expect, test } from "vitest";
import { cleanText, emptySummary, mergeSummary } from "../src/core/clean/hidden-chars";

describe("cleanText", () => {
  test("清走零寬度空格同 BOM", () => {
    const r = cleanText("A\u200BB\uFEFFC");
    expect(r.text).toBe("ABC");
    expect(r.summary.removed).toEqual({ "U+200B": 1, "U+FEFF": 1 });
  });

  test("清走 Tag 字元（可以收埋隱形句子）", () => {
    const hidden = String.fromCodePoint(0xe0041, 0xe0042);
    const r = cleanText(`hi${hidden}`);
    expect(r.text).toBe("hi");
    expect(r.summary.removed).toEqual({ "U+E0041": 1, "U+E0042": 1 });
  });

  test("特殊空格換成普通空格，但保留全形空格 U+3000", () => {
    const r = cleanText("a\u202Fb\u2003c\u3000d");
    expect(r.text).toBe("a b c\u3000d");
    expect(r.summary.replaced).toEqual({ "U+202F": 1, "U+2003": 1 });
  });

  test("中英文之間嘅 ZWJ / ZWNJ 清走", () => {
    const r = cleanText("中\u200D文 a\u200Cb");
    expect(r.text).toBe("中文 ab");
    expect(r.summary.removed).toEqual({ "U+200D": 1, "U+200C": 1 });
  });

  test("emoji 家庭序列入面嘅 ZWJ 保留", () => {
    const family = "👨\u200D👩\u200D👧";
    const r = cleanText(family);
    expect(r.text).toBe(family);
    expect(r.summary.kept).toBe(2);
  });

  test("膚色 emoji 後面嘅 ZWJ 保留", () => {
    const coder = "👩🏽\u200D💻";
    expect(cleanText(coder).text).toBe(coder);
  });

  test("波斯文入面嘅 ZWNJ 保留", () => {
    const word = "می\u200Cخواهم";
    const r = cleanText(word);
    expect(r.text).toBe(word);
    expect(r.summary.kept).toBe(1);
  });

  test("方向標記：段落有希伯來文就保留，冇就清走", () => {
    expect(cleanText("שלום\u200F abc").text).toBe("שלום\u200F abc");
    expect(cleanText("abc\u200F def").text).toBe("abc def");
  });

  test("方向覆寫一定清走，就算有阿拉伯文", () => {
    expect(cleanText("مرحبا\u202Eabc").text).toBe("مرحباabc");
  });

  test("異體選擇符：emoji 同漢字後面保留，英文後面清走", () => {
    expect(cleanText("❤\uFE0F").text).toBe("❤\uFE0F");
    expect(cleanText("葛\uFE00").text).toBe("葛\uFE00");
    expect(cleanText("1\uFE0F⃣").text).toBe("1\uFE0F⃣");
    expect(cleanText("a\uFE0Fb").text).toBe("ab");
  });

  test("上下文可以由前後文字片段提供", () => {
    const r = cleanText("\u200D👩", { before: "👨", after: "", rtl: false });
    expect(r.text).toBe("\u200D👩");
  });

  test("冇隱形字元就原封不動", () => {
    const r = cleanText("Hello 你好");
    expect(r.text).toBe("Hello 你好");
    expect(r.summary).toEqual(emptySummary());
  });
});

test("mergeSummary 加埋兩份統計", () => {
  const a = { removed: { "U+200B": 1 }, replaced: {}, kept: 1 };
  const b = { removed: { "U+200B": 2, "U+FEFF": 1 }, replaced: { "U+202F": 1 }, kept: 0 };
  expect(mergeSummary(a, b)).toEqual({
    removed: { "U+200B": 3, "U+FEFF": 1 },
    replaced: { "U+202F": 1 },
    kept: 1,
  });
});
```

- [ ] **Step 2：行測試，確認失敗**

Run: `npx vitest run tests/hidden-chars.test.ts`
Expected: FAIL，`Failed to resolve import "../src/core/clean/hidden-chars"`

- [ ] **Step 3：實作**

`src/core/clean/hidden-chars.ts`

```ts
import type { HiddenCharSummary } from "../model";

// 規則見設計文件 5.6

function inRanges(cp: number, ranges: [number, number][]): boolean {
  return ranges.some(([lo, hi]) => cp >= lo && cp <= hi);
}

const ALWAYS_REMOVE: [number, number][] = [
  [0x00ad, 0x00ad], // 軟連字號
  [0x180e, 0x180e], // 蒙古文元音分隔符
  [0x200b, 0x200b], // 零寬度空格
  [0x202d, 0x202e], // 方向覆寫
  [0x2060, 0x2064], // 字詞連接符、隱形運算符
  [0xfeff, 0xfeff], // BOM
  [0xe0000, 0xe007f], // Tag 字元
];

const REPLACE_WITH_SPACE: [number, number][] = [
  [0x2000, 0x200a],
  [0x202f, 0x202f],
  [0x205f, 0x205f],
];

const JOINERS: [number, number][] = [[0x200c, 0x200d]]; // ZWNJ、ZWJ
const BIDI_MARKS: [number, number][] = [
  [0x061c, 0x061c],
  [0x200e, 0x200f],
  [0x202a, 0x202c],
  [0x2066, 0x2069],
];
const VARIATION_SELECTORS: [number, number][] = [[0xfe00, 0xfe0f]];

// 需要 ZWJ / ZWNJ 嘅文字同 emoji
const NEEDS_JOINER =
  /[\p{Script=Arabic}\p{Script=Syriac}\p{Script=Nko}\p{Script=Mongolian}\p{Script=Devanagari}\p{Script=Bengali}\p{Script=Gurmukhi}\p{Script=Gujarati}\p{Script=Oriya}\p{Script=Tamil}\p{Script=Telugu}\p{Script=Kannada}\p{Script=Malayalam}\p{Script=Sinhala}\p{Script=Khmer}\p{Script=Myanmar}\p{Extended_Pictographic}\p{Emoji_Modifier}\uFE0F]/u;
const NEEDS_VARIATION = /[\p{Extended_Pictographic}\p{Script=Han}0-9#*]/u;
const RTL = /[\p{Script=Arabic}\p{Script=Hebrew}\p{Script=Syriac}\p{Script=Thaana}\p{Script=Nko}]/u;

export function emptySummary(): HiddenCharSummary {
  return { removed: {}, replaced: {}, kept: 0 };
}

export function mergeSummary(a: HiddenCharSummary, b: HiddenCharSummary): HiddenCharSummary {
  const add = (x: Record<string, number>, y: Record<string, number>) => {
    const out = { ...x };
    for (const [k, v] of Object.entries(y)) out[k] = (out[k] ?? 0) + v;
    return out;
  };
  return { removed: add(a.removed, b.removed), replaced: add(a.replaced, b.replaced), kept: a.kept + b.kept };
}

export function hasRtl(text: string): boolean {
  return RTL.test(text);
}

function label(cp: number): string {
  return "U+" + cp.toString(16).toUpperCase().padStart(4, "0");
}

export interface CleanContext {
  before: string; // 前面嘅文字（同一段落）
  after: string; // 後面嘅文字（同一段落）
  rtl: boolean; // 段落有冇由右至左文字
}

export function cleanText(
  text: string,
  ctx: CleanContext = { before: "", after: "", rtl: hasRtl(text) },
): { text: string; summary: HiddenCharSummary } {
  const cps = Array.from(text);
  const prevOf = (i: number) => (i > 0 ? cps[i - 1] : Array.from(ctx.before).at(-1) ?? "");
  const nextOf = (i: number) => (i < cps.length - 1 ? cps[i + 1] : Array.from(ctx.after)[0] ?? "");
  const summary = emptySummary();
  let out = "";
  const bump = (bucket: Record<string, number>, cp: number) => {
    bucket[label(cp)] = (bucket[label(cp)] ?? 0) + 1;
  };

  cps.forEach((ch, i) => {
    const cp = ch.codePointAt(0)!;
    if (inRanges(cp, ALWAYS_REMOVE)) {
      bump(summary.removed, cp);
    } else if (inRanges(cp, REPLACE_WITH_SPACE)) {
      bump(summary.replaced, cp);
      out += " ";
    } else if (inRanges(cp, JOINERS)) {
      if (NEEDS_JOINER.test(prevOf(i)) || NEEDS_JOINER.test(nextOf(i))) {
        summary.kept++;
        out += ch;
      } else bump(summary.removed, cp);
    } else if (inRanges(cp, BIDI_MARKS)) {
      if (ctx.rtl) {
        summary.kept++;
        out += ch;
      } else bump(summary.removed, cp);
    } else if (inRanges(cp, VARIATION_SELECTORS)) {
      if (NEEDS_VARIATION.test(prevOf(i))) {
        summary.kept++;
        out += ch;
      } else bump(summary.removed, cp);
    } else {
      out += ch;
    }
  });
  return { text: out, summary };
}
```

- [ ] **Step 4：行測試，確認通過**

Run: `npx vitest run tests/hidden-chars.test.ts`
Expected: `Tests  13 passed (13)`

- [ ] **Step 5：Commit（由用家做）**

```bash
git add src/core/clean/hidden-chars.ts tests/hidden-chars.test.ts
git commit -m "feat(core): clean hidden characters with language-aware rules"
```

---

### Task 4：OOXML 工具

**Files:**
- Create: `src/core/ooxml/xml.ts`, `src/core/ooxml/package.ts`
- Test: `tests/ooxml.test.ts`

**Interfaces:**
- Produces:
  - `NS`（命名空間常數：`w`、`r`、`rel`、`wp`、`a`、`wps`、`mc`、`v`、`cp`、`ep`、`custom`）
  - `parseXml(text): Element`、`kids(el, ns?, local?): Element[]`、`kid(el, ns, local): Element | undefined`、`desc(el, ns, local): Element[]`
  - `attr(el, ns | null, local): string | undefined`、`wval(el)`、`wnum(el, local = "val")`、`toggle(el): boolean | undefined`
  - `class OoxmlPackage { static open(bytes): Promise<OoxmlPackage>; paths(): string[]; has(path); text(path); bytes(path); xml(path); rels(partPath): Promise<Map<string, Relationship>> }`
  - `interface Relationship { id; type; target; external }`、`resolvePath(baseDir, target): string`

- [ ] **Step 1：寫失敗測試**

`tests/ooxml.test.ts`

```ts
import JSZip from "jszip";
import { expect, test } from "vitest";
import { OoxmlPackage, resolvePath } from "../src/core/ooxml/package";
import { kids, NS, parseXml, toggle } from "../src/core/ooxml/xml";

test("resolvePath 處理相對路徑", () => {
  expect(resolvePath("word", "media/image1.png")).toBe("word/media/image1.png");
  expect(resolvePath("word", "../customXml/item1.xml")).toBe("customXml/item1.xml");
  expect(resolvePath("word", "/word/header1.xml")).toBe("word/header1.xml");
  expect(resolvePath("", "word/document.xml")).toBe("word/document.xml");
});

test("toggle 讀 OOXML 開關屬性", () => {
  const root = parseXml(
    `<w:rPr xmlns:w="${NS.w}"><w:b/><w:i w:val="0"/><w:u w:val="single"/></w:rPr>`,
  );
  const [b, i, u] = kids(root, NS.w);
  expect(toggle(b)).toBe(true);
  expect(toggle(i)).toBe(false);
  expect(toggle(u)).toBe(true);
  expect(toggle(undefined)).toBe(undefined);
});

test("OoxmlPackage 讀 relationships", async () => {
  const zip = new JSZip();
  zip.file("word/document.xml", "<x/>");
  zip.file(
    "word/_rels/document.xml.rels",
    `<Relationships xmlns="${NS.rel}">
      <Relationship Id="rId1" Type="http://x/relationships/image" Target="media/a.png"/>
      <Relationship Id="rId2" Type="http://x/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>
    </Relationships>`,
  );
  const pkg = await OoxmlPackage.open(await zip.generateAsync({ type: "uint8array" }));
  const rels = await pkg.rels("word/document.xml");
  expect(rels.get("rId1")).toEqual({ id: "rId1", type: "image", target: "word/media/a.png", external: false });
  expect(rels.get("rId2")?.target).toBe("https://example.com");
  expect(pkg.paths().sort()).toEqual(["word/_rels/document.xml.rels", "word/document.xml"]);
});
```

- [ ] **Step 2：行測試，確認失敗**

Run: `npx vitest run tests/ooxml.test.ts`
Expected: FAIL，`Failed to resolve import "../src/core/ooxml/package"`

- [ ] **Step 3：實作 `src/core/ooxml/xml.ts`**

`src/core/ooxml/xml.ts`

```ts
import { DOMParser, type Element } from "@xmldom/xmldom";

export type { Element };

export const NS = {
  w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  rel: "http://schemas.openxmlformats.org/package/2006/relationships",
  wp: "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
  a: "http://schemas.openxmlformats.org/drawingml/2006/main",
  wps: "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
  mc: "http://schemas.openxmlformats.org/markup-compatibility/2006",
  v: "urn:schemas-microsoft-com:vml",
  cp: "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
  ep: "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
  custom: "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties",
} as const;

export function parseXml(text: string): Element {
  const doc = new DOMParser({
    onError: (level, msg) => {
      if (level === "fatalError") throw new Error(`XML parse error: ${msg}`);
    },
  }).parseFromString(text, "text/xml");
  const root = doc.documentElement;
  if (!root) throw new Error("XML parse error: no root element");
  return root;
}

/** 直接子元素 */
export function kids(el: Element, ns?: string, local?: string): Element[] {
  const out: Element[] = [];
  for (let n = el.firstChild; n; n = n.nextSibling) {
    if (n.nodeType !== 1) continue;
    const e = n as Element;
    if (ns !== undefined && e.namespaceURI !== ns) continue;
    if (local !== undefined && e.localName !== local) continue;
    out.push(e);
  }
  return out;
}

export function kid(el: Element | undefined, ns: string, local: string): Element | undefined {
  return el ? kids(el, ns, local)[0] : undefined;
}

/** 所有後代元素 */
export function desc(el: Element, ns: string, local: string): Element[] {
  return Array.from(el.getElementsByTagNameNS(ns, local)) as Element[];
}

export function attr(el: Element | undefined, ns: string | null, local: string): string | undefined {
  if (!el) return undefined;
  const v = ns ? el.getAttributeNS(ns, local) : el.getAttribute(local);
  return v === null || v === "" ? undefined : v;
}

/** w:val 屬性 */
export function wval(el: Element | undefined): string | undefined {
  return attr(el, NS.w, "val");
}

export function wnum(el: Element | undefined, local = "val"): number | undefined {
  const v = attr(el, NS.w, local);
  if (v === undefined) return undefined;
  const n = Number(v);
  return Number.isFinite(n) ? n : undefined;
}

/** OOXML 開關屬性：<w:b/> = true，<w:b w:val="0"/> = false，冇元素 = undefined */
export function toggle(el: Element | undefined): boolean | undefined {
  if (!el) return undefined;
  const v = wval(el);
  return !(v === "0" || v === "false" || v === "off" || v === "none");
}

export function textContent(el: Element): string {
  return el.textContent ?? "";
}
```

- [ ] **Step 4：實作 `src/core/ooxml/package.ts`**

`src/core/ooxml/package.ts`

```ts
import JSZip from "jszip";
import { attr, kids, NS, parseXml, type Element } from "./xml";

export interface Relationship {
  id: string;
  type: string; // 最後一節，例如 "image"、"hyperlink"、"header"
  target: string; // 已解析成套件內嘅完整路徑；外部連結就係原網址
  external: boolean;
}

export class OoxmlPackage {
  private constructor(private zip: JSZip) {}

  static async open(bytes: Uint8Array): Promise<OoxmlPackage> {
    return new OoxmlPackage(await JSZip.loadAsync(bytes));
  }

  /** 所有檔案路徑（唔包括資料夾） */
  paths(): string[] {
    return Object.values(this.zip.files)
      .filter((f) => !f.dir)
      .map((f) => f.name);
  }

  has(path: string): boolean {
    const f = this.zip.file(path);
    return f !== null && !f.dir;
  }

  async text(path: string): Promise<string | undefined> {
    return this.zip.file(path)?.async("string");
  }

  async bytes(path: string): Promise<Uint8Array | undefined> {
    return this.zip.file(path)?.async("uint8array");
  }

  async xml(path: string): Promise<Element | undefined> {
    const t = await this.text(path);
    return t === undefined ? undefined : parseXml(t);
  }

  /** 讀取某個 part 嘅 relationships，例如 word/document.xml → word/_rels/document.xml.rels */
  async rels(partPath: string): Promise<Map<string, Relationship>> {
    const dir = partPath.includes("/") ? partPath.slice(0, partPath.lastIndexOf("/")) : "";
    const file = partPath.slice(partPath.lastIndexOf("/") + 1);
    const relsPath = `${dir ? dir + "/" : ""}_rels/${file}.rels`;
    const root = await this.xml(relsPath);
    const out = new Map<string, Relationship>();
    if (!root) return out;
    for (const r of kids(root, NS.rel, "Relationship")) {
      const id = attr(r, null, "Id");
      const target = attr(r, null, "Target");
      const type = attr(r, null, "Type") ?? "";
      if (!id || !target) continue;
      const external = attr(r, null, "TargetMode") === "External";
      out.set(id, {
        id,
        type: type.slice(type.lastIndexOf("/") + 1),
        target: external ? target : resolvePath(dir, target),
        external,
      });
    }
    return out;
  }
}

export function resolvePath(baseDir: string, target: string): string {
  if (target.startsWith("/")) return target.slice(1);
  const parts = baseDir ? baseDir.split("/") : [];
  for (const seg of target.split("/")) {
    if (seg === "..") parts.pop();
    else if (seg !== "." && seg !== "") parts.push(seg);
  }
  return parts.join("/");
}
```

- [ ] **Step 5：行測試，確認通過**

Run: `npx vitest run tests/ooxml.test.ts`
Expected: `Tests  3 passed (3)`

- [ ] **Step 6：Commit（由用家做）**

```bash
git add src/core/ooxml tests/ooxml.test.ts
git commit -m "feat(core): add OOXML xml and package helpers"
```

---

### Task 5：Word 讀取器

**Files:**
- Create: `tests/fixtures/docx-builder.ts`（測試用，之後所有 task 共用）
- Create: `src/core/readers/docx/styles.ts`, `src/core/readers/docx/numbering.ts`, `src/core/readers/docx/body.ts`, `src/core/readers/docx/index.ts`
- Test: `tests/read-docx.test.ts`

**Interfaces:**
- Consumes: Task 2 型別、Task 4 `OoxmlPackage` 同 XML 工具
- Produces:
  - `readDocx(bytes: Uint8Array): Promise<{ doc: CmDocument; pkg: OoxmlPackage }>`
  - `MAIN_PART = "word/document.xml"`
  - 測試工具：`buildDocx(f: DocxFixture): Promise<Uint8Array>`、`p(text, { style?, rPr?, pPr? })`、`imageP(rId, cx?, cy?)`、`wordTextWatermark(text)`、`PNG_1PX`、`pngWith(marker)`、`NS_DECL`

重點行為：
- 樣式繼承：文件預設 → 段落樣式鏈 → 字元樣式鏈 → 直接設定；冇設定時字體 10pt
- 標題：`outlineLvl` 0–5 → 1–6，或者樣式名「heading N」/「Title」
- 遇到圖片、分頁、文字方塊會切開段落；圖片段落唔會留低空段落
- 頁首頁尾、文字後面 (`behindDoc`) 嘅圖形、WordArt (`v:textpath`) 一律唔搬；段落因此變空就唔保留
- 修訂追蹤：`w:ins` 保留、`w:del` 唔要；功能變數只保留結果文字
- 表格：`gridSpan` → `colSpan`；`vMerge` → 起點格嘅 `rowSpan`，被合併嘅格唔出現
- Symbol / Wingdings 私人區項目符號換成 `•` `◦` `▪`
- 圖片頁數用 `w:lastRenderedPageBreak` 同分頁符估計

- [ ] **Step 1：建立測試用 .docx 產生器**

`tests/fixtures/docx-builder.ts`

```ts
import JSZip from "jszip";

// 用純 XML 砌測試用 .docx，唔使 Word，內容一睇就明

export const NS_DECL = [
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"',
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"',
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"',
  'xmlns:v="urn:schemas-microsoft-com:vml"',
  'xmlns:o="urn:schemas-microsoft-com:office:office"',
  'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"',
].join(" ");

export interface MediaEntry {
  rId: string;
  name: string; // 例如 "image1.png"
  bytes: Uint8Array;
}

export interface DocxFixture {
  body: string; // <w:body> 入面嘅內容（唔包 sectPr）
  sectPr?: string; // 唔提供就用 A4 預設
  styles?: string; // <w:styles> 入面嘅內容
  numbering?: string; // <w:numbering> 入面嘅內容
  header?: string; // <w:hdr> 入面嘅內容
  footer?: string; // <w:ftr> 入面嘅內容
  footnotes?: string; // <w:footnotes> 入面嘅一般註腳
  core?: Record<string, string>; // 例如 { "dc:creator": "ChatGPT" }
  app?: Record<string, string>; // 例如 { Application: "Some AI" }
  custom?: Record<string, string>;
  customXml?: string; // customXml/item1.xml 嘅內容
  media?: MediaEntry[]; // document.xml 用
  headerMedia?: MediaEntry[]; // header1.xml 用
  links?: Record<string, string>; // rId → 網址
  extraParts?: Record<string, string | Uint8Array>;
}

const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

function rels(entries: string[]): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${entries.join("")}</Relationships>`;
}

function rel(id: string, type: string, target: string, external = false): string {
  return `<Relationship Id="${id}" Type="${REL}/${type}" Target="${target}"${external ? ' TargetMode="External"' : ""}/>`;
}

export async function buildDocx(f: DocxFixture): Promise<Uint8Array> {
  const zip = new JSZip();
  const docRels: string[] = [rel("rIdStyles", "styles", "styles.xml")];
  const hdrRefs: string[] = [];

  zip.file(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="jpeg" ContentType="image/jpeg"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
  );
  zip.file("_rels/.rels", rels([
    rel("rId1", "officeDocument", "word/document.xml"),
    `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>`,
    rel("rId3", "extended-properties", "docProps/app.xml"),
    ...(f.custom ? [rel("rId4", "custom-properties", "docProps/custom.xml")] : []),
  ]));

  const core = Object.entries(f.core ?? {})
    .map(([k, v]) => `<${k}>${v}</${k}>`)
    .join("");
  zip.file(
    "docProps/core.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">${core}</cp:coreProperties>`,
  );
  const app = Object.entries(f.app ?? {})
    .map(([k, v]) => `<${k}>${v}</${k}>`)
    .join("");
  zip.file(
    "docProps/app.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">${app}</Properties>`,
  );
  if (f.custom) {
    const props = Object.entries(f.custom)
      .map(([k, v], i) => `<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="${i + 2}" name="${k}"><vt:lpwstr>${v}</vt:lpwstr></property>`)
      .join("");
    zip.file(
      "docProps/custom.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">${props}</Properties>`,
    );
  }
  if (f.customXml) zip.file("customXml/item1.xml", f.customXml);

  zip.file(
    "word/styles.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ${NS_DECL}>${f.styles ?? ""}</w:styles>`,
  );
  if (f.numbering) {
    docRels.push(rel("rIdNumbering", "numbering", "numbering.xml"));
    zip.file("word/numbering.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering ${NS_DECL}>${f.numbering}</w:numbering>`);
  }
  if (f.header !== undefined) {
    docRels.push(rel("rIdHeader", "header", "header1.xml"));
    hdrRefs.push(`<w:headerReference w:type="default" r:id="rIdHeader"/>`);
    zip.file("word/header1.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr ${NS_DECL}>${f.header}</w:hdr>`);
    zip.file("word/_rels/header1.xml.rels", rels((f.headerMedia ?? []).map((m) => rel(m.rId, "image", `media/${m.name}`))));
    for (const m of f.headerMedia ?? []) zip.file(`word/media/${m.name}`, m.bytes);
  }
  if (f.footer !== undefined) {
    docRels.push(rel("rIdFooter", "footer", "footer1.xml"));
    hdrRefs.push(`<w:footerReference w:type="default" r:id="rIdFooter"/>`);
    zip.file("word/footer1.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr ${NS_DECL}>${f.footer}</w:ftr>`);
  }
  if (f.footnotes !== undefined) {
    docRels.push(rel("rIdFootnotes", "footnotes", "footnotes.xml"));
    zip.file(
      "word/footnotes.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:footnotes ${NS_DECL}><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>${f.footnotes}</w:footnotes>`,
    );
  }
  for (const m of f.media ?? []) {
    docRels.push(rel(m.rId, "image", `media/${m.name}`));
    zip.file(`word/media/${m.name}`, m.bytes);
  }
  for (const [id, url] of Object.entries(f.links ?? {})) docRels.push(rel(id, "hyperlink", url, true));

  const sectPr =
    f.sectPr ??
    `<w:sectPr>${hdrRefs.join("")}<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>`;
  zip.file(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS_DECL}><w:body>${f.body}${sectPr}</w:body></w:document>`,
  );
  zip.file("word/_rels/document.xml.rels", rels(docRels));

  for (const [p, c] of Object.entries(f.extraParts ?? {})) zip.file(p, c);
  return zip.generateAsync({ type: "uint8array" });
}

// 1×1 透明 PNG
export const PNG_1PX = Uint8Array.from(
  atob("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=="),
  (c) => c.charCodeAt(0),
);

/** 喺 PNG 後面加一段 ASCII 標記，模擬帶 EXIF / XMP / C2PA 嘅圖片 */
export function pngWith(marker: string): Uint8Array {
  const m = new TextEncoder().encode(marker);
  const out = new Uint8Array(PNG_1PX.length + m.length);
  out.set(PNG_1PX);
  out.set(m, PNG_1PX.length);
  return out;
}

/** 一段文字嘅 <w:p> */
export function p(text: string, opts: { style?: string; rPr?: string; pPr?: string } = {}): string {
  const pPr = opts.style || opts.pPr ? `<w:pPr>${opts.style ? `<w:pStyle w:val="${opts.style}"/>` : ""}${opts.pPr ?? ""}</w:pPr>` : "";
  const rPr = opts.rPr ? `<w:rPr>${opts.rPr}</w:rPr>` : "";
  return `<w:p>${pPr}<w:r>${rPr}<w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
}

/** 內嵌圖片嘅 <w:p> */
export function imageP(rId: string, cx = 1905000, cy = 952500): string {
  return `<w:p><w:r><w:drawing><wp:inline><wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="1" name="Picture 1"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="${rId}"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
}

/** Word 內建文字水印（擺喺頁首） */
export function wordTextWatermark(text: string): string {
  return `<w:p><w:r><w:pict><v:shape id="PowerPlusWaterMarkObject357476642" o:spid="_x0000_s2049" type="#_x0000_t136" style="position:absolute;margin-left:0;margin-top:0;width:412.4pt;height:247.4pt;rotation:315;z-index:-251657216" fillcolor="silver" stroked="f"><v:fill opacity=".5"/><v:textpath style="font-family:&quot;Calibri&quot;;font-size:1pt" string="${text}"/></v:shape></w:pict></w:r></w:p>`;
}
```

- [ ] **Step 2：寫失敗測試**

`tests/read-docx.test.ts`

```ts
import { describe, expect, test } from "vitest";
import type { Block, ImageSlot, Paragraph, Table, TextInline } from "../src/core/model";
import { readDocx } from "../src/core/readers/docx";
import { buildDocx, imageP, p, PNG_1PX, wordTextWatermark } from "./fixtures/docx-builder";

const paras = (blocks: Block[]) => blocks.filter((b): b is Paragraph => b.type === "paragraph");
const texts = (para: Paragraph) =>
  para.inlines.filter((i): i is TextInline => i.type === "text").map((i) => i.text).join("");

const STYLES = `
  <w:docDefaults>
    <w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:eastAsia="PMingLiU"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/>
    <w:pPr><w:spacing w:before="240"/><w:outlineLvl w:val="0"/></w:pPr>
    <w:rPr><w:b/><w:color w:val="2F5496"/><w:sz w:val="32"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="MyHeading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="character" w:styleId="Emph"><w:name w:val="Emph"/><w:rPr><w:i/></w:rPr></w:style>
  <w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/>
    <w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4"/></w:tblBorders></w:tblPr></w:style>`;

describe("readDocx：格式", () => {
  test("樣式繼承：預設 → 段落樣式 → 字元樣式 → 直接設定", async () => {
    const bytes = await buildDocx({
      styles: STYLES,
      body:
        p("標題", { style: "Heading1" }) +
        `<w:p><w:r><w:rPr><w:rStyle w:val="Emph"/><w:u w:val="single"/><w:color w:val="FF0000"/></w:rPr><w:t>斜</w:t></w:r></w:p>`,
    });
    const { doc } = await readDocx(bytes);
    const [h, body] = paras(doc.body);
    expect(h.heading).toBe(1);
    expect(h.spacingBefore).toBe(240);
    expect(h.spacingAfter).toBe(160);
    expect(h.lineSpacing).toEqual({ line: 259, rule: "auto" });
    expect((h.inlines[0] as TextInline).style).toMatchObject({
      bold: true, color: "2F5496", size: 16, font: "Calibri", fontEastAsia: "PMingLiU",
    });
    expect((body.inlines[0] as TextInline).style).toMatchObject({
      bold: false, italic: true, underline: true, color: "FF0000", size: 11,
    });
  });

  test("用樣式名稱判斷標題層級", async () => {
    const { doc } = await readDocx(await buildDocx({ styles: STYLES, body: p("二", { style: "MyHeading2" }) }));
    expect(paras(doc.body)[0].heading).toBe(2);
  });

  test("對齊、縮排、懸掛縮排", async () => {
    const { doc } = await readDocx(
      await buildDocx({ body: p("x", { pPr: `<w:jc w:val="both"/><w:ind w:left="720" w:hanging="360"/>` }) }),
    );
    expect(paras(doc.body)[0]).toMatchObject({ align: "justify", indentLeft: 720, indentFirstLine: -360 });
  });

  test("tab、換行、超連結、註腳參照", async () => {
    const bytes = await buildDocx({
      links: { rIdL: "https://example.com" },
      footnotes: `<w:footnote w:id="1">${p("註腳內容")}</w:footnote>`,
      body: `<w:p><w:r><w:t>a</w:t><w:tab/><w:t>b</w:t><w:br/><w:t>c</w:t></w:r><w:hyperlink r:id="rIdL"><w:r><w:t>連結</w:t></w:r></w:hyperlink><w:r><w:footnoteReference w:id="1"/></w:r></w:p>`,
    });
    const { doc } = await readDocx(bytes);
    const inl = paras(doc.body)[0].inlines;
    expect(inl.map((i) => i.type)).toEqual(["text", "break", "text", "text", "footnoteRef"]);
    expect((inl[0] as TextInline).text).toBe("a\tb");
    expect((inl[3] as TextInline).link).toBe("https://example.com");
    expect(texts(paras(doc.footnotes[1])[0])).toBe("註腳內容");
  });

  test("修訂追蹤：新增保留、刪除唔要", async () => {
    const { doc } = await readDocx(
      await buildDocx({
        body: `<w:p><w:r><w:t>A</w:t></w:r><w:ins w:id="1" w:author="x"><w:r><w:t>B</w:t></w:r></w:ins><w:del w:id="2" w:author="x"><w:r><w:delText>C</w:delText></w:r></w:del></w:p>`,
      }),
    );
    expect(texts(paras(doc.body)[0])).toBe("AB");
  });

  test("功能變數只保留結果文字", async () => {
    const { doc } = await readDocx(
      await buildDocx({
        body: `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>3</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`,
      }),
    );
    expect(texts(paras(doc.body)[0])).toBe("3");
  });

  test("頁面設定", async () => {
    const { doc } = await readDocx(
      await buildDocx({
        body: p("x"),
        sectPr: `<w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/></w:sectPr>`,
      }),
    );
    expect(doc.page).toEqual({
      widthTwip: 16838, heightTwip: 11906, landscape: true, margins: { top: 720, right: 720, bottom: 720, left: 720 },
    });
  });
});

describe("readDocx：列表", () => {
  test("讀取編號格式同層級，Symbol 字型項目符號換成通用符號", async () => {
    const numbering = `
      <w:abstractNum w:abstractNumId="0">
        <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>
        <w:lvl w:ilvl="1"><w:numFmt w:val="bullet"/><w:lvlText w:val=""/></w:lvl>
      </w:abstractNum>
      <w:num w:numId="5"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="3"/></w:lvlOverride></w:num>`;
    const li = (t: string, lvl: number) => p(t, { pPr: `<w:numPr><w:ilvl w:val="${lvl}"/><w:numId w:val="5"/></w:numPr>` });
    const { doc } = await readDocx(await buildDocx({ numbering, body: li("一", 0) + li("甲", 1) + p("普通") }));
    const [a, b, c] = paras(doc.body);
    expect(a.list).toEqual({ numId: "5", level: 0 });
    expect(b.list).toEqual({ numId: "5", level: 1 });
    expect(c.list).toBeUndefined();
    expect(doc.lists["5"][0]).toEqual({ format: "decimal", text: "%1.", start: 3, indentLeft: 720, hanging: 360 });
    expect(doc.lists["5"][1].text).toBe("◦");
  });
});

describe("readDocx：表格", () => {
  test("合併儲存格同框線", async () => {
    const tc = (t: string, pr = "") => `<w:tc><w:tcPr>${pr}</w:tcPr>${p(t)}</w:tc>`;
    const body = `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="3000"/></w:tblGrid>
      <w:tr>${tc("A", `<w:gridSpan w:val="2"/>`)}</w:tr>
      <w:tr>${tc("B", `<w:vMerge w:val="restart"/>`)}${tc("C")}</w:tr>
      <w:tr>${tc("", `<w:vMerge/>`)}${tc("D")}</w:tr>
    </w:tbl>`;
    const { doc } = await readDocx(await buildDocx({ styles: STYLES, body }));
    const t = doc.body[0] as Table;
    expect(t.type).toBe("table");
    expect(t.columnWidths).toEqual([2000, 3000]);
    expect(t.borders).toBe(true);
    expect(t.rows.map((r) => r.map((c) => [texts(c.blocks[0] as Paragraph), c.colSpan, c.rowSpan]))).toEqual([
      [["A", 2, 1]],
      [["B", 1, 2], ["C", 1, 1]],
      [["D", 1, 1]],
    ]);
  });
});

describe("readDocx：圖片同圖形", () => {
  test("內嵌圖片切開段落，記錄尺寸同頁數", async () => {
    const bytes = await buildDocx({
      media: [{ rId: "rIdImg", name: "image1.png", bytes: PNG_1PX }],
      body:
        p("前") +
        `<w:p><w:r><w:lastRenderedPageBreak/><w:t>第二頁</w:t></w:r></w:p>` +
        imageP("rIdImg", 1905000, 952500),
    });
    const { doc } = await readDocx(bytes);
    const img = doc.body.find((b): b is ImageSlot => b.type === "image")!;
    expect(img).toEqual({
      type: "image", index: 1, widthEmu: 1905000, heightEmu: 952500, page: 2,
      mediaPath: "word/media/image1.png", contentType: "png", keep: false,
    });
    expect(doc.media["word/media/image1.png"]).toEqual(PNG_1PX);
    expect(doc.body.filter((b) => b.type === "paragraph")).toHaveLength(2); // 圖片段落唔會留低空段落
  });

  test("文字後面嘅圖形同 WordArt 唔搬", async () => {
    const behind = `<w:p><w:r><w:drawing><wp:anchor behindDoc="1"><wp:extent cx="100" cy="100"/><a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="rIdImg"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>`;
    const bytes = await buildDocx({
      media: [{ rId: "rIdImg", name: "image1.png", bytes: PNG_1PX }],
      body: p("內文") + behind + wordTextWatermark("CONFIDENTIAL"),
    });
    const { doc } = await readDocx(bytes);
    expect(doc.body.some((b) => b.type === "image")).toBe(false);
    expect(JSON.stringify(doc)).not.toContain("CONFIDENTIAL");
  });

  test("文字方塊內容變做普通段落", async () => {
    const box = `<w:p><w:r><w:drawing><wp:inline><wp:extent cx="100" cy="100"/><a:graphic><a:graphicData><wps:wsp><wps:txbx><w:txbxContent>${p("方塊入面")}</w:txbxContent></wps:txbx></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
    const { doc } = await readDocx(await buildDocx({ body: box }));
    expect(paras(doc.body).map(texts)).toContain("方塊入面");
  });

  test("方程式變做冇原圖嘅圖片位置", async () => {
    const { doc } = await readDocx(await buildDocx({ body: `<m:oMathPara><m:oMath><m:r><m:t>x</m:t></m:r></m:oMath></m:oMathPara>` }));
    expect(doc.body[0]).toMatchObject({ type: "image", mediaPath: "", contentType: "other" });
  });
});

describe("readDocx：頁首頁尾", () => {
  test("只搬文字，水印圖形同圖片唔搬", async () => {
    const bytes = await buildDocx({
      body: p("內文"),
      header: p("公司報告") + wordTextWatermark("AI Generated") + imageP("rIdLogo"),
      headerMedia: [{ rId: "rIdLogo", name: "logo.png", bytes: PNG_1PX }],
      footer: p("第 1 頁"),
    });
    const { doc } = await readDocx(bytes);
    expect(paras(doc.header).map(texts)).toEqual(["公司報告"]);
    expect(paras(doc.footer).map(texts)).toEqual(["第 1 頁"]);
    expect(doc.header.some((b) => b.type === "image")).toBe(false);
    expect(Object.keys(doc.media)).toEqual([]);
    expect(paras(doc.header)[0].id).toBe("h1");
  });
});
```

- [ ] **Step 3：行測試，確認失敗**

Run: `npx vitest run tests/read-docx.test.ts`
Expected: FAIL，`Failed to resolve import "../src/core/readers/docx"`

- [ ] **Step 4：實作樣式繼承 `src/core/readers/docx/styles.ts`**

`src/core/readers/docx/styles.ts`

```ts
import type { Align, LineSpacing, ListRef, RunStyle } from "../../model";
import { attr, kid, kids, NS, toggle, wnum, wval, type Element } from "../../ooxml/xml";

// 樣式繼承：文件預設 → 段落樣式（由最底嘅 basedOn 開始）→ 字元樣式 → 直接設定

interface StyleDef {
  type: string;
  basedOn?: string;
  name: string;
  pPr?: Element;
  rPr?: Element;
  tblPr?: Element;
}

type PartialRun = Partial<RunStyle>;

export interface ParaProps {
  heading: number;
  align: Align;
  indentLeft: number;
  indentFirstLine: number;
  spacingBefore: number;
  spacingAfter: number;
  lineSpacing: LineSpacing;
  list?: ListRef;
}

const DEFAULT_SIZE_PT = 10; // Word 喺冇任何設定時嘅預設字體大細

function applyRPr(acc: PartialRun, rPr: Element | undefined): void {
  if (!rPr) return;
  const w = (n: string) => kid(rPr, NS.w, n);
  const b = toggle(w("b"));
  if (b !== undefined) acc.bold = b;
  const i = toggle(w("i"));
  if (i !== undefined) acc.italic = i;
  const u = w("u");
  if (u) acc.underline = wval(u) !== "none";
  const strike = toggle(w("strike")) ?? toggle(w("dstrike"));
  if (strike !== undefined) acc.strike = strike;
  const color = wval(w("color"));
  if (color) acc.color = color === "auto" ? undefined : color.toUpperCase();
  const hl = wval(w("highlight"));
  if (hl) acc.highlight = hl === "none" ? undefined : hl;
  const sz = wnum(w("sz"));
  if (sz !== undefined) acc.size = sz / 2;
  const fonts = w("rFonts");
  if (fonts) {
    const ascii = attr(fonts, NS.w, "ascii") ?? attr(fonts, NS.w, "hAnsi");
    const ea = attr(fonts, NS.w, "eastAsia");
    if (ascii) acc.font = ascii;
    if (ea) acc.fontEastAsia = ea;
  }
  const va = wval(w("vertAlign"));
  if (va === "superscript" || va === "subscript") acc.vertAlign = va;
  else if (va === "baseline") acc.vertAlign = undefined;
}

interface PartialPara {
  heading?: number;
  align?: Align;
  indentLeft?: number;
  indentFirstLine?: number;
  spacingBefore?: number;
  spacingAfter?: number;
  lineSpacing?: LineSpacing;
  numId?: string;
  ilvl?: number;
}

const ALIGN: Record<string, Align> = {
  left: "left",
  start: "left",
  center: "center",
  right: "right",
  end: "right",
  both: "justify",
  distribute: "justify",
};

function applyPPr(acc: PartialPara, pPr: Element | undefined): void {
  if (!pPr) return;
  const w = (n: string) => kid(pPr, NS.w, n);
  const jc = wval(w("jc"));
  if (jc && ALIGN[jc]) acc.align = ALIGN[jc];
  const ind = w("ind");
  if (ind) {
    const left = wnum(ind, "left") ?? wnum(ind, "start");
    if (left !== undefined) acc.indentLeft = left;
    const first = wnum(ind, "firstLine");
    const hanging = wnum(ind, "hanging");
    if (hanging !== undefined) acc.indentFirstLine = -hanging;
    else if (first !== undefined) acc.indentFirstLine = first;
  }
  const sp = w("spacing");
  if (sp) {
    const before = wnum(sp, "before");
    const after = wnum(sp, "after");
    const line = wnum(sp, "line");
    if (before !== undefined) acc.spacingBefore = before;
    if (after !== undefined) acc.spacingAfter = after;
    if (line !== undefined) {
      const rule = attr(sp, NS.w, "lineRule");
      acc.lineSpacing = { line, rule: rule === "exact" ? "exact" : rule === "atLeast" ? "atLeast" : "auto" };
    }
  }
  const ol = wnum(w("outlineLvl"));
  if (ol !== undefined) acc.heading = ol >= 0 && ol <= 5 ? ol + 1 : 0;
  const numPr = w("numPr");
  if (numPr) {
    const numId = wval(kid(numPr, NS.w, "numId"));
    const ilvl = wnum(kid(numPr, NS.w, "ilvl"));
    if (numId !== undefined) acc.numId = numId;
    if (ilvl !== undefined) acc.ilvl = ilvl;
  }
}

function headingFromName(name: string): number | undefined {
  const m = /^heading\s*([1-6])$/i.exec(name.trim());
  if (m) return Number(m[1]);
  if (/^title$/i.test(name.trim())) return 1;
  return undefined;
}

export class StyleSheet {
  private styles = new Map<string, StyleDef>();
  private defaultRPr?: Element;
  private defaultPPr?: Element;
  private defaultParaStyle?: string;

  constructor(root: Element | undefined) {
    if (!root) return;
    const defaults = kid(root, NS.w, "docDefaults");
    this.defaultRPr = kid(kid(defaults, NS.w, "rPrDefault"), NS.w, "rPr");
    this.defaultPPr = kid(kid(defaults, NS.w, "pPrDefault"), NS.w, "pPr");
    for (const s of kids(root, NS.w, "style")) {
      const id = attr(s, NS.w, "styleId");
      if (!id) continue;
      const type = attr(s, NS.w, "type") ?? "paragraph";
      this.styles.set(id, {
        type,
        basedOn: wval(kid(s, NS.w, "basedOn")),
        name: wval(kid(s, NS.w, "name")) ?? id,
        pPr: kid(s, NS.w, "pPr"),
        rPr: kid(s, NS.w, "rPr"),
        tblPr: kid(s, NS.w, "tblPr"),
      });
      const isDefault = attr(s, NS.w, "default");
      if (type === "paragraph" && (isDefault === "1" || isDefault === "true")) this.defaultParaStyle = id;
    }
  }

  /** 由最底嘅 basedOn 開始排，避免循環 */
  private chain(id: string | undefined): StyleDef[] {
    const out: StyleDef[] = [];
    const seen = new Set<string>();
    let cur = id;
    while (cur && !seen.has(cur)) {
      seen.add(cur);
      const s = this.styles.get(cur);
      if (!s) break;
      out.unshift(s);
      cur = s.basedOn;
    }
    return out;
  }

  paraStyleId(pPr: Element | undefined): string | undefined {
    return wval(kid(pPr, NS.w, "pStyle")) ?? this.defaultParaStyle;
  }

  resolveRun(pStyleId: string | undefined, rPr: Element | undefined): RunStyle {
    const acc: PartialRun = {};
    applyRPr(acc, this.defaultRPr);
    for (const s of this.chain(pStyleId)) applyRPr(acc, s.rPr);
    const rStyle = wval(kid(rPr, NS.w, "rStyle"));
    for (const s of this.chain(rStyle)) applyRPr(acc, s.rPr);
    applyRPr(acc, rPr);
    return {
      font: acc.font,
      fontEastAsia: acc.fontEastAsia,
      size: acc.size ?? DEFAULT_SIZE_PT,
      bold: acc.bold ?? false,
      italic: acc.italic ?? false,
      underline: acc.underline ?? false,
      strike: acc.strike ?? false,
      color: acc.color,
      highlight: acc.highlight,
      vertAlign: acc.vertAlign,
    };
  }

  resolvePara(pPr: Element | undefined): ParaProps {
    const styleId = this.paraStyleId(pPr);
    const acc: PartialPara = {};
    applyPPr(acc, this.defaultPPr);
    for (const s of this.chain(styleId)) {
      const h = headingFromName(s.name);
      if (h !== undefined) acc.heading = h;
      applyPPr(acc, s.pPr);
    }
    applyPPr(acc, pPr);
    const list = acc.numId && acc.numId !== "0" ? { numId: acc.numId, level: acc.ilvl ?? 0 } : undefined;
    return {
      heading: acc.heading ?? 0,
      align: acc.align ?? "left",
      indentLeft: acc.indentLeft ?? 0,
      indentFirstLine: acc.indentFirstLine ?? 0,
      spacingBefore: acc.spacingBefore ?? 0,
      spacingAfter: acc.spacingAfter ?? 0,
      lineSpacing: acc.lineSpacing ?? { line: 240, rule: "auto" },
      list,
    };
  }

  tableHasBorders(tblPr: Element | undefined): boolean {
    const visible = (borders: Element | undefined) =>
      !!borders && kids(borders, NS.w).some((b) => !["none", "nil"].includes(wval(b) ?? "none"));
    const direct = kid(tblPr, NS.w, "tblBorders");
    if (direct) return visible(direct);
    const styleId = wval(kid(tblPr, NS.w, "tblStyle"));
    for (const s of this.chain(styleId).reverse()) {
      const b = kid(s.tblPr, NS.w, "tblBorders");
      if (b) return visible(b);
    }
    return false;
  }
}
```

- [ ] **Step 5：實作列表 `src/core/readers/docx/numbering.ts`**

`src/core/readers/docx/numbering.ts`

```ts
import type { ListLevel } from "../../model";
import { attr, kid, kids, NS, wnum, wval, type Element } from "../../ooxml/xml";

const BULLETS = ["•", "◦", "▪"];

/** Word 常用 Symbol / Wingdings 字型嘅私人區字元做項目符號，冇原字型就會變亂碼，所以換成通用符號 */
function normaliseBullet(text: string, level: number): string {
  const cp = text.codePointAt(0);
  if (!text || (cp !== undefined && cp >= 0xe000 && cp <= 0xf8ff)) return BULLETS[level % 3];
  return text;
}

function readLevel(lvl: Element, level: number): ListLevel {
  const format = wval(kid(lvl, NS.w, "numFmt")) ?? "decimal";
  const rawText = wval(kid(lvl, NS.w, "lvlText")) ?? "";
  const ind = kid(kid(lvl, NS.w, "pPr"), NS.w, "ind");
  return {
    format,
    text: format === "bullet" ? normaliseBullet(rawText, level) : rawText,
    start: wnum(kid(lvl, NS.w, "start")) ?? 1,
    indentLeft: wnum(ind, "left") ?? wnum(ind, "start") ?? 720 * (level + 1),
    hanging: wnum(ind, "hanging") ?? 360,
  };
}

/** numId → 九層設定 */
export function readNumbering(root: Element | undefined): Record<string, ListLevel[]> {
  const out: Record<string, ListLevel[]> = {};
  if (!root) return out;
  const abstracts = new Map<string, Element>();
  for (const a of kids(root, NS.w, "abstractNum")) {
    const id = attr(a, NS.w, "abstractNumId");
    if (id !== undefined) abstracts.set(id, a);
  }
  for (const num of kids(root, NS.w, "num")) {
    const numId = attr(num, NS.w, "numId");
    const absId = wval(kid(num, NS.w, "abstractNumId"));
    const abs = absId !== undefined ? abstracts.get(absId) : undefined;
    if (numId === undefined || !abs) continue;
    const levels: ListLevel[] = [];
    for (let i = 0; i < 9; i++) {
      const lvl = kids(abs, NS.w, "lvl").find((l) => attr(l, NS.w, "ilvl") === String(i));
      levels.push(
        lvl
          ? readLevel(lvl, i)
          : { format: "bullet", text: BULLETS[i % 3], start: 1, indentLeft: 720 * (i + 1), hanging: 360 },
      );
    }
    for (const ov of kids(num, NS.w, "lvlOverride")) {
      const i = Number(attr(ov, NS.w, "ilvl"));
      const start = wnum(kid(ov, NS.w, "startOverride"));
      if (levels[i] && start !== undefined) levels[i].start = start;
    }
    out[numId] = levels;
  }
  return out;
}
```

- [ ] **Step 6：實作內文 `src/core/readers/docx/body.ts`**

`src/core/readers/docx/body.ts`

```ts
import type { Block, ImageSlot, Inline, Paragraph, RunStyle, Table, TableCell } from "../../model";
import type { OoxmlPackage, Relationship } from "../../ooxml/package";
import { attr, desc, kid, kids, NS, wnum, wval, type Element } from "../../ooxml/xml";
import type { StyleSheet } from "./styles";

const NS_PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture";
const NS_M = "http://schemas.openxmlformats.org/officeDocument/2006/math";
const EMU_PER_PT = 12700;
const PLACEHOLDER_EMU = { w: 2743200, h: 457200 }; // 冇尺寸資料時：3 吋 × 0.5 吋

/** 成份文件共用嘅計數器同收集器 */
export interface ReadState {
  pkg: OoxmlPackage;
  styles: StyleSheet;
  media: Record<string, Uint8Array>;
  imageCount: number;
  page: number;
  paraCount: Record<string, number>;
}

/** 每個 part（內文、頁首、註腳）各自嘅設定 */
export interface PartCtx {
  rels: Map<string, Relationship>;
  allowGraphics: boolean; // 頁首頁尾 = false：圖形一律唔搬
  idPrefix: string; // "p"、"h"、"f"、"n"
}

function contentTypeOf(path: string): ImageSlot["contentType"] {
  const ext = path.slice(path.lastIndexOf(".") + 1).toLowerCase();
  if (ext === "png") return "png";
  if (ext === "jpg" || ext === "jpeg") return "jpg";
  if (ext === "gif") return "gif";
  if (ext === "bmp") return "bmp";
  return "other";
}

async function imageSlot(
  st: ReadState,
  ctx: PartCtx,
  rId: string | undefined,
  widthEmu: number,
  heightEmu: number,
): Promise<ImageSlot> {
  const rel = rId ? ctx.rels.get(rId) : undefined;
  let mediaPath = "";
  if (rel && !rel.external) {
    const bytes = await st.pkg.bytes(rel.target);
    if (bytes) {
      mediaPath = rel.target;
      st.media[mediaPath] = bytes;
    }
  }
  return {
    type: "image",
    index: ++st.imageCount,
    widthEmu: widthEmu || PLACEHOLDER_EMU.w,
    heightEmu: heightEmu || PLACEHOLDER_EMU.h,
    page: st.page,
    mediaPath,
    contentType: mediaPath ? contentTypeOf(mediaPath) : "other",
    keep: false,
  };
}

/** 段落讀取中途嘅狀態：遇到圖片、分頁、文字方塊就要切開段落 */
class ParaBuilder {
  out: Block[] = [];
  extra: Block[] = []; // 文字方塊內容，擺喺段落後面
  inlines: Inline[] = [];
  split = false; // 已經切開過（遇過圖片或者分頁）
  skipped = false; // 有圖形被略過：如果段落因此變空，就唔留低空段落

  constructor(
    private st: ReadState,
    private ctx: PartCtx,
    private base: Omit<Paragraph, "id" | "inlines">,
  ) {}

  addText(text: string, style: RunStyle, link?: string): void {
    const last = this.inlines.at(-1);
    if (last?.type === "text" && last.link === link && JSON.stringify(last.style) === JSON.stringify(style)) {
      last.text += text;
    } else {
      this.inlines.push({ type: "text", text, style: { ...style }, ...(link ? { link } : {}) });
    }
  }

  flush(force: boolean): void {
    if (this.inlines.length > 0 || (force && !this.split && !this.skipped)) {
      const n = (this.st.paraCount[this.ctx.idPrefix] = (this.st.paraCount[this.ctx.idPrefix] ?? 0) + 1);
      this.out.push({ ...this.base, id: `${this.ctx.idPrefix}${n}`, inlines: this.inlines });
    }
    this.inlines = [];
  }

  addBlock(b: Block): void {
    this.flush(false);
    this.split = true;
    this.out.push(b);
  }
}

export async function readBlocks(container: Element, st: ReadState, ctx: PartCtx): Promise<Block[]> {
  const out: Block[] = [];
  for (const el of kids(container)) {
    if (el.namespaceURI === NS.w) {
      if (el.localName === "p") out.push(...(await readParagraph(el, st, ctx)));
      else if (el.localName === "tbl") out.push(await readTable(el, st, ctx));
      else if (el.localName === "sdt") {
        const content = kid(el, NS.w, "sdtContent");
        if (content) out.push(...(await readBlocks(content, st, ctx)));
      } else if (el.localName === "customXml") out.push(...(await readBlocks(el, st, ctx)));
    } else if (el.namespaceURI === NS_M && el.localName === "oMathPara") {
      out.push(await imageSlot(st, ctx, undefined, 0, 0));
    }
  }
  return out;
}

async function readParagraph(p: Element, st: ReadState, ctx: PartCtx): Promise<Block[]> {
  const pPr = kid(p, NS.w, "pPr");
  const props = st.styles.resolvePara(pPr);
  const pStyleId = st.styles.paraStyleId(pPr);
  const pb = new ParaBuilder(st, ctx, { type: "paragraph", ...props });
  await readInlineContainer(p, pb, st, ctx, pStyleId, undefined);
  pb.flush(true);
  return [...pb.out, ...pb.extra];
}

const PASS_THROUGH = new Set(["ins", "moveTo", "smartTag", "customXml", "fldSimple", "dir", "bdo"]);

async function readInlineContainer(
  el: Element,
  pb: ParaBuilder,
  st: ReadState,
  ctx: PartCtx,
  pStyleId: string | undefined,
  link: string | undefined,
): Promise<void> {
  for (const c of kids(el)) {
    if (c.namespaceURI === NS_M && (c.localName === "oMath" || c.localName === "oMathPara")) {
      if (ctx.allowGraphics) pb.addBlock(await imageSlot(st, ctx, undefined, 0, 0));
      else pb.skipped = true;
      continue;
    }
    if (c.namespaceURI !== NS.w) continue;
    const name = c.localName ?? "";
    if (name === "r") await readRun(c, pb, st, ctx, pStyleId, link);
    else if (name === "hyperlink") {
      const rId = attr(c, NS.r, "id");
      const rel = rId ? ctx.rels.get(rId) : undefined;
      await readInlineContainer(c, pb, st, ctx, pStyleId, rel?.external ? rel.target : link);
    } else if (name === "sdt") {
      const content = kid(c, NS.w, "sdtContent");
      if (content) await readInlineContainer(content, pb, st, ctx, pStyleId, link);
    } else if (PASS_THROUGH.has(name)) {
      await readInlineContainer(c, pb, st, ctx, pStyleId, link);
    }
    // del、moveFrom（已刪除）、bookmark、proofErr 等一律略過
  }
}

async function readRun(
  r: Element,
  pb: ParaBuilder,
  st: ReadState,
  ctx: PartCtx,
  pStyleId: string | undefined,
  link: string | undefined,
): Promise<void> {
  const style = st.styles.resolveRun(pStyleId, kid(r, NS.w, "rPr"));
  await readRunChildren(kids(r), pb, st, ctx, style, link);
}

async function readRunChildren(
  children: Element[],
  pb: ParaBuilder,
  st: ReadState,
  ctx: PartCtx,
  style: RunStyle,
  link: string | undefined,
): Promise<void> {
  for (const c of children) {
    if (c.namespaceURI === NS.mc && c.localName === "AlternateContent") {
      // 只讀 Choice，避免同 Fallback 重複
      const choice = kid(c, NS.mc, "Choice") ?? kid(c, NS.mc, "Fallback");
      if (choice) await readRunChildren(kids(choice), pb, st, ctx, style, link);
      continue;
    }
    if (c.namespaceURI !== NS.w) continue;
    switch (c.localName) {
      case "t":
        pb.addText(c.textContent ?? "", style, link);
        break;
      case "tab":
        pb.addText("\t", style, link);
        break;
      case "noBreakHyphen":
        pb.addText("-", style, link);
        break;
      case "br":
        if (attr(c, NS.w, "type") === "page") {
          pb.addBlock({ type: "pageBreak" });
          st.page++;
        } else pb.inlines.push({ type: "break" });
        break;
      case "cr":
        pb.inlines.push({ type: "break" });
        break;
      case "lastRenderedPageBreak":
        st.page++;
        break;
      case "footnoteReference": {
        const id = wnum(c, "id");
        if (id !== undefined) pb.inlines.push({ type: "footnoteRef", id });
        break;
      }
      case "drawing":
        await readDrawing(c, pb, st, ctx);
        break;
      case "pict":
      case "object":
        await readVml(c, pb, st, ctx);
        break;
      // instrText（功能變數指令）、delText（已刪除文字）等略過
    }
  }
}

async function readDrawing(drawing: Element, pb: ParaBuilder, st: ReadState, ctx: PartCtx): Promise<void> {
  const holder = kid(drawing, NS.wp, "inline") ?? kid(drawing, NS.wp, "anchor");
  const behindText = holder?.localName === "anchor" && ["1", "true"].includes(attr(holder, null, "behindDoc") ?? "");
  // 頁首頁尾嘅圖形、擺喺文字後面嘅圖形 = 水印嘅常見做法，唔搬（檢查器會報告）
  if (!ctx.allowGraphics || !holder || behindText) {
    pb.skipped = true;
    return;
  }
  const extent = kid(holder, NS.wp, "extent");
  const cx = Number(attr(extent, null, "cx") ?? 0);
  const cy = Number(attr(extent, null, "cy") ?? 0);

  const txbx = desc(holder, NS.w, "txbxContent")[0];
  if (txbx) {
    pb.extra.push(...(await readBlocks(txbx, st, ctx)));
    return;
  }
  const blip = desc(holder, NS.a, "blip")[0];
  const isPicture = desc(holder, NS_PIC, "pic").length > 0;
  if (blip || isPicture) {
    pb.addBlock(await imageSlot(st, ctx, attr(blip, NS.r, "embed"), cx, cy));
    return;
  }
  // 圖表、SmartArt 等：當圖片處理，冇原圖可以保留
  pb.addBlock(await imageSlot(st, ctx, undefined, cx, cy));
}

function vmlSizeEmu(shape: Element | undefined): { w: number; h: number } {
  const style = attr(shape, null, "style") ?? "";
  const num = (key: string) => {
    const m = new RegExp(`(?:^|;)\\s*${key}:\\s*([\\d.]+)pt`).exec(style);
    return m ? Math.round(Number(m[1]) * EMU_PER_PT) : 0;
  };
  return { w: num("width"), h: num("height") };
}

async function readVml(pict: Element, pb: ParaBuilder, st: ReadState, ctx: PartCtx): Promise<void> {
  // WordArt（v:textpath）通常就係水印，唔搬
  if (!ctx.allowGraphics || desc(pict, NS.v, "textpath").length > 0) {
    pb.skipped = true;
    return;
  }
  const txbx = desc(pict, NS.w, "txbxContent")[0];
  if (txbx) {
    pb.extra.push(...(await readBlocks(txbx, st, ctx)));
    return;
  }
  const img = desc(pict, NS.v, "imagedata")[0];
  if (img) {
    const shape = desc(pict, NS.v, "shape")[0];
    const { w, h } = vmlSizeEmu(shape);
    pb.addBlock(await imageSlot(st, ctx, attr(img, NS.r, "id"), w, h));
  }
}

async function readTable(tbl: Element, st: ReadState, ctx: PartCtx): Promise<Table> {
  const tblPr = kid(tbl, NS.w, "tblPr");
  const grid = kids(kid(tbl, NS.w, "tblGrid") ?? tbl, NS.w, "gridCol").map((g) => wnum(g, "w") ?? 0);
  const rows: TableCell[][] = [];
  const open = new Map<number, TableCell>(); // 欄位置 → 未完嘅垂直合併起點

  const trs = kids(tbl, NS.w, "tr");
  for (const tr of trs) {
    const row: TableCell[] = [];
    let col = 0;
    const tcs = kids(tr).flatMap((c) =>
      c.namespaceURI === NS.w && c.localName === "sdt"
        ? kids(kid(c, NS.w, "sdtContent") ?? c, NS.w, "tc")
        : c.namespaceURI === NS.w && c.localName === "tc"
          ? [c]
          : [],
    );
    for (const tc of tcs) {
      const tcPr = kid(tc, NS.w, "tcPr");
      const span = wnum(kid(tcPr, NS.w, "gridSpan")) ?? 1;
      const vMerge = kid(tcPr, NS.w, "vMerge");
      const vVal = wval(vMerge);
      if (vMerge && vVal !== "restart" && open.has(col)) {
        open.get(col)!.rowSpan++;
        col += span;
        continue;
      }
      const cell: TableCell = { blocks: await readBlocks(tc, st, ctx), colSpan: span, rowSpan: 1 };
      if (vMerge && vVal === "restart") open.set(col, cell);
      else open.delete(col);
      row.push(cell);
      col += span;
    }
    rows.push(row);
  }
  const colCount = Math.max(1, ...rows.map((r) => r.reduce((n, c) => n + c.colSpan, 0)));
  const columnWidths = grid.length > 0 ? grid : Array<number>(colCount).fill(2000);
  return { type: "table", columnWidths, rows, borders: st.styles.tableHasBorders(tblPr) };
}
```

- [ ] **Step 7：實作入口 `src/core/readers/docx/index.ts`**

`src/core/readers/docx/index.ts`

```ts
import type { Block, CmDocument, PageSetup } from "../../model";
import { OoxmlPackage } from "../../ooxml/package";
import { kid, kids, NS, attr, wnum, type Element } from "../../ooxml/xml";
import { readBlocks, type ReadState } from "./body";
import { readNumbering } from "./numbering";
import { StyleSheet } from "./styles";

export const MAIN_PART = "word/document.xml";

function readPage(sectPr: Element | undefined): PageSetup {
  const pgSz = kid(sectPr, NS.w, "pgSz");
  const pgMar = kid(sectPr, NS.w, "pgMar");
  return {
    widthTwip: wnum(pgSz, "w") ?? 11906,
    heightTwip: wnum(pgSz, "h") ?? 16838,
    landscape: attr(pgSz, NS.w, "orient") === "landscape",
    margins: {
      top: wnum(pgMar, "top") ?? 1440,
      right: wnum(pgMar, "right") ?? 1440,
      bottom: wnum(pgMar, "bottom") ?? 1440,
      left: wnum(pgMar, "left") ?? 1440,
    },
  };
}

async function readHeaderFooter(
  st: ReadState,
  sectPr: Element | undefined,
  kind: "header" | "footer",
  docRels: Awaited<ReturnType<OoxmlPackage["rels"]>>,
): Promise<Block[]> {
  const refs = sectPr ? kids(sectPr, NS.w, `${kind}Reference`) : [];
  const ref = refs.find((r) => (attr(r, NS.w, "type") ?? "default") === "default");
  const rel = ref ? docRels.get(attr(ref, NS.r, "id") ?? "") : undefined;
  if (!rel) return [];
  const root = await st.pkg.xml(rel.target);
  if (!root) return [];
  return readBlocks(root, st, {
    rels: await st.pkg.rels(rel.target),
    allowGraphics: false,
    idPrefix: kind === "header" ? "h" : "f",
  });
}

async function readFootnotes(st: ReadState): Promise<Record<number, Block[]>> {
  const path = "word/footnotes.xml";
  const root = await st.pkg.xml(path);
  const out: Record<number, Block[]> = {};
  if (!root) return out;
  const rels = await st.pkg.rels(path);
  for (const fn of kids(root, NS.w, "footnote")) {
    if (attr(fn, NS.w, "type")) continue; // separator 等
    const id = wnum(fn, "id");
    if (id === undefined) continue;
    out[id] = await readBlocks(fn, st, { rels, allowGraphics: false, idPrefix: "n" });
  }
  return out;
}

export async function readDocx(bytes: Uint8Array): Promise<{ doc: CmDocument; pkg: OoxmlPackage }> {
  const pkg = await OoxmlPackage.open(bytes);
  const root = await pkg.xml(MAIN_PART);
  const body = root ? kid(root, NS.w, "body") : undefined;
  if (!body) throw new Error("word/document.xml 冇 <w:body>");

  const st: ReadState = {
    pkg,
    styles: new StyleSheet(await pkg.xml("word/styles.xml")),
    media: {},
    imageCount: 0,
    page: 1,
    paraCount: {},
  };
  const docRels = await pkg.rels(MAIN_PART);
  const blocks = await readBlocks(body, st, { rels: docRels, allowGraphics: true, idPrefix: "p" });
  const sectPr = kid(body, NS.w, "sectPr");

  const doc: CmDocument = {
    kind: "docx",
    page: readPage(sectPr),
    body: blocks,
    header: await readHeaderFooter(st, sectPr, "header", docRels),
    footer: await readHeaderFooter(st, sectPr, "footer", docRels),
    footnotes: await readFootnotes(st),
    lists: readNumbering(await pkg.xml("word/numbering.xml")),
    media: st.media,
  };
  return { doc, pkg };
}
```

- [ ] **Step 8：行測試，確認通過**

Run: `npx vitest run tests/read-docx.test.ts`
Expected: `Tests  14 passed (14)`

- [ ] **Step 9：型別檢查**

Run: `npx tsc --noEmit -p .`
Expected: 冇輸出

- [ ] **Step 10：Commit（由用家做）**

```bash
git add src/core/readers tests/fixtures/docx-builder.ts tests/read-docx.test.ts
git commit -m "feat(core): read docx into the content model"
```

---

### Task 6：走訪工具同用家選擇

**Files:**
- Create: `src/core/walk.ts`, `src/core/clean/apply-choices.ts`
- Test: `tests/apply-choices.test.ts`

**Interfaces:**
- Consumes: Task 2 型別、Task 3 `cleanText` / `hasRtl` / `emptySummary` / `mergeSummary`、Task 5 `readDocx` 同測試工具
- Produces:
  - `walkBlocks(blocks): Generator<Block>`、`allBlocks(doc): Generator<Block>`（次序：內文 → 頁首 → 頁尾 → 註腳）
  - `allParagraphs(doc): Paragraph[]`、`allImages(doc): ImageSlot[]`、`paragraphText(p): string`
  - `mapBlocks(blocks, fn: (b) => Block | null): Block[]`
  - `cleanParagraph(p): { paragraph: Paragraph; summary: HiddenCharSummary }`
  - `cleanParagraphSummary(p): HiddenCharSummary`
  - `applyChoices(doc: CmDocument, choices: Choices): CmDocument`（唔改原本嘅 doc；冇揀保留嘅原圖會喺 `media` 入面移除）

- [ ] **Step 1：寫失敗測試**

`tests/apply-choices.test.ts`

```ts
import { expect, test } from "vitest";
import { applyChoices } from "../src/core/clean/apply-choices";
import type { ImageSlot, Paragraph, TextInline } from "../src/core/model";
import { readDocx } from "../src/core/readers/docx";
import { allImages, allParagraphs, paragraphText } from "../src/core/walk";
import { buildDocx, imageP, p, PNG_1PX } from "./fixtures/docx-builder";

const fixture = () =>
  buildDocx({
    body:
      p("第一段\u200B") +
      p("本文由 AI 生成") +
      `<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc>${p("表格\u200B入面")}</w:tc></w:tr></w:tbl>` +
      imageP("rIdA") +
      imageP("rIdB"),
    media: [
      { rId: "rIdA", name: "a.png", bytes: PNG_1PX },
      { rId: "rIdB", name: "b.png", bytes: PNG_1PX },
    ],
    header: p("頁首\u200B"),
  });

test("刪除選咗嘅段落、清隱形字元（包括表格同頁首）", async () => {
  const { doc } = await readDocx(await fixture());
  const out = applyChoices(doc, { removeParagraphs: ["p2"], keepImages: [] });
  expect(allParagraphs(out).map(paragraphText)).toEqual(["第一段", "表格入面", "頁首"]);
});

test("圖片：揀咗保留先保留原圖，其餘佔位；冇保留嘅原圖唔會留喺內容清單", async () => {
  const { doc } = await readDocx(await fixture());
  const out = applyChoices(doc, { removeParagraphs: [], keepImages: [2] });
  expect(allImages(out).map((i: ImageSlot) => [i.index, i.keep])).toEqual([
    [1, false],
    [2, true],
  ]);
  expect(Object.keys(out.media)).toEqual(["word/media/b.png"]);
});

test("唔會改動原本嘅內容清單", async () => {
  const { doc } = await readDocx(await fixture());
  applyChoices(doc, { removeParagraphs: ["p1"], keepImages: [] });
  const first = doc.body[0] as Paragraph;
  expect((first.inlines[0] as TextInline).text).toBe("第一段\u200B");
});
```

- [ ] **Step 2：行測試，確認失敗**

Run: `npx vitest run tests/apply-choices.test.ts`
Expected: FAIL，`Failed to resolve import "../src/core/clean/apply-choices"`

- [ ] **Step 3：實作 `src/core/walk.ts`**

`src/core/walk.ts`

```ts
import type { Block, CmDocument, ImageSlot, Paragraph, TextInline } from "./model";

/** 逐個區塊行一次（包括表格入面），由上至下 */
export function* walkBlocks(blocks: Block[]): Generator<Block> {
  for (const b of blocks) {
    yield b;
    if (b.type === "table") for (const row of b.rows) for (const cell of row) yield* walkBlocks(cell.blocks);
  }
}

/** 文件入面所有區塊：內文、頁首、頁尾、註腳 */
export function* allBlocks(doc: CmDocument): Generator<Block> {
  yield* walkBlocks(doc.body);
  yield* walkBlocks(doc.header);
  yield* walkBlocks(doc.footer);
  for (const fn of Object.values(doc.footnotes)) yield* walkBlocks(fn);
}

export function allParagraphs(doc: CmDocument): Paragraph[] {
  return [...allBlocks(doc)].filter((b): b is Paragraph => b.type === "paragraph");
}

export function allImages(doc: CmDocument): ImageSlot[] {
  return [...allBlocks(doc)].filter((b): b is ImageSlot => b.type === "image");
}

export function paragraphText(p: Paragraph): string {
  return p.inlines
    .filter((i): i is TextInline => i.type === "text")
    .map((i) => i.text)
    .join("");
}

/** 由頂層開始重建區塊樹：fn 回傳 null 就刪除嗰個區塊 */
export function mapBlocks(blocks: Block[], fn: (b: Block) => Block | null): Block[] {
  const out: Block[] = [];
  for (const b of blocks) {
    const mapped = fn(b);
    if (!mapped) continue;
    if (mapped.type === "table") {
      out.push({
        ...mapped,
        rows: mapped.rows.map((row) => row.map((cell) => ({ ...cell, blocks: mapBlocks(cell.blocks, fn) }))),
      });
    } else out.push(mapped);
  }
  return out;
}
```

- [ ] **Step 4：實作 `src/core/clean/apply-choices.ts`**

`src/core/clean/apply-choices.ts`

```ts
import {
  KEEPABLE_IMAGE_TYPES,
  type Block,
  type Choices,
  type CmDocument,
  type HiddenCharSummary,
  type Inline,
  type Paragraph,
  type TextInline,
} from "../model";
import { mapBlocks } from "../walk";
import { cleanText, emptySummary, hasRtl, mergeSummary } from "./hidden-chars";

/** 清走段落入面嘅隱形字元；上下文用成個段落嘅文字 */
export function cleanParagraph(p: Paragraph): { paragraph: Paragraph; summary: HiddenCharSummary } {
  const texts = p.inlines.map((i) => (i.type === "text" ? i.text : ""));
  const rtl = hasRtl(texts.join(""));
  let summary = emptySummary();
  const inlines: Inline[] = [];
  p.inlines.forEach((inl, idx) => {
    if (inl.type !== "text") {
      inlines.push(inl);
      return;
    }
    const before = texts.slice(0, idx).join("");
    const after = texts.slice(idx + 1).join("");
    const r = cleanText(inl.text, { before, after, rtl });
    summary = mergeSummary(summary, r.summary);
    if (r.text) inlines.push({ ...inl, text: r.text } satisfies TextInline);
  });
  return { paragraph: { ...p, inlines }, summary };
}

export function cleanParagraphSummary(p: Paragraph): HiddenCharSummary {
  return cleanParagraph(p).summary;
}

/** 按用家選擇產生最終內容清單：刪段落、清隱形字元、決定圖片保留定佔位 */
export function applyChoices(doc: CmDocument, choices: Choices): CmDocument {
  const remove = new Set(choices.removeParagraphs);
  const keep = new Set(choices.keepImages);
  const keptMedia: Record<string, Uint8Array> = {};

  const fn = (b: Block): Block | null => {
    if (b.type === "paragraph") return remove.has(b.id) ? null : cleanParagraph(b).paragraph;
    if (b.type === "image") {
      const canKeep = keep.has(b.index) && KEEPABLE_IMAGE_TYPES.has(b.contentType) && !!doc.media[b.mediaPath];
      if (canKeep) keptMedia[b.mediaPath] = doc.media[b.mediaPath];
      return { ...b, keep: canKeep };
    }
    return b;
  };

  const footnotes: CmDocument["footnotes"] = {};
  for (const [id, blocks] of Object.entries(doc.footnotes)) footnotes[Number(id)] = mapBlocks(blocks, fn);

  return {
    ...doc,
    body: mapBlocks(doc.body, fn),
    header: mapBlocks(doc.header, fn),
    footer: mapBlocks(doc.footer, fn),
    footnotes,
    media: keptMedia, // 冇揀保留嘅原圖唔會再出現喺記憶體入面嘅內容清單
  };
}
```

- [ ] **Step 5：行測試，確認通過**

Run: `npx vitest run tests/apply-choices.test.ts`
Expected: `Tests  3 passed (3)`

- [ ] **Step 6：Commit（由用家做）**

```bash
git add src/core/walk.ts src/core/clean/apply-choices.ts tests/apply-choices.test.ts
git commit -m "feat(core): apply user choices and clean paragraphs"
```

---

### Task 7：檢查器

**Files:**
- Create: `src/core/inspect/bytes.ts`, `src/core/inspect/phrases.ts`, `src/core/inspect/metadata.ts`, `src/core/inspect/visible-docx.ts`, `src/core/inspect/docx.ts`
- Test: `tests/inspect-docx.test.ts`

**Interfaces:**
- Consumes: Task 4、Task 5 `readDocx`、Task 6 `allImages` / `allParagraphs` / `paragraphText` / `cleanParagraphSummary`
- Produces:
  - `inspectDocx(pkg: OoxmlPackage, doc: CmDocument): Promise<Report>`
  - `findVisibleDocx(pkg): Promise<VisibleItem[]>`（自我檢查會用）
  - `readOoxmlMetadata(pkg): Promise<MetadataItem[]>`（第 2 期 PowerPoint 會重用）
  - `hasC2pa(bytes): boolean`、`imageTags(bytes): ("exif" | "xmp" | "c2pa")[]`、`containsAscii(hay, needle)`
  - `findPhrases(doc): PhraseItem[]`、`PHRASE_PATTERNS: RegExp[]`

- [ ] **Step 1：寫失敗測試**

`tests/inspect-docx.test.ts`

```ts
import { describe, expect, test } from "vitest";
import { inspectDocx } from "../src/core/inspect/docx";
import { readDocx } from "../src/core/readers/docx";
import { buildDocx, imageP, p, PNG_1PX, pngWith, wordTextWatermark } from "./fixtures/docx-builder";

async function inspect(f: Parameters<typeof buildDocx>[0]) {
  const { doc, pkg } = await readDocx(await buildDocx(f));
  return inspectDocx(pkg, doc);
}

describe("inspectDocx：可見水印", () => {
  test("Word 內建文字水印同頁首 logo", async () => {
    const r = await inspect({
      body: p("內文"),
      header: wordTextWatermark("AI Generated") + imageP("rIdLogo"),
      headerMedia: [{ rId: "rIdLogo", name: "logo.png", bytes: PNG_1PX }],
    });
    expect(r.visible).toEqual([
      { code: "wordTextWatermark", part: "word/header1.xml", text: "AI Generated" },
      { code: "headerImage", part: "word/header1.xml" },
    ]);
  });

  test("文字後面嘅圖形", async () => {
    const behind = `<w:p><w:r><w:drawing><wp:anchor behindDoc="1"><wp:extent cx="1" cy="1"/><a:graphic><a:graphicData><wps:wsp><wps:txbx><w:txbxContent>${p("DRAFT")}</w:txbxContent></wps:txbx></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>`;
    const r = await inspect({ body: p("x") + behind });
    expect(r.visible).toEqual([{ code: "behindText", part: "word/document.xml", text: "DRAFT" }]);
  });

  test("乾淨文件冇可見水印", async () => {
    expect((await inspect({ body: p("hello") })).visible).toEqual([]);
  });
});

describe("inspectDocx：metadata", () => {
  test("core、app、custom、customXml", async () => {
    const r = await inspect({
      body: p("x"),
      core: { "dc:creator": "ChatGPT", "cp:lastModifiedBy": "Bot" },
      app: { Application: "Some AI Writer", Company: "" },
      custom: { GeneratedBy: "Gemini" },
      customXml: "<data><source>ai-platform</source></data>",
    });
    expect(r.metadata).toEqual([
      { part: "docProps/core.xml", field: "dc:creator", value: "ChatGPT" },
      { part: "docProps/core.xml", field: "cp:lastModifiedBy", value: "Bot" },
      { part: "docProps/app.xml", field: "Application", value: "Some AI Writer" },
      { part: "docProps/custom.xml", field: "GeneratedBy", value: "Gemini" },
      { part: "customXml/item1.xml", field: "customXml", value: "ai-platform" },
    ]);
  });
});

describe("inspectDocx：C2PA、圖片、可疑字句、隱形字元", () => {
  test("搵到 C2PA 同圖片標記", async () => {
    const r = await inspect({
      body: imageP("rIdA") + imageP("rIdB"),
      media: [
        { rId: "rIdA", name: "a.png", bytes: pngWith("Exif\0\0...c2pa") },
        { rId: "rIdB", name: "b.png", bytes: PNG_1PX },
      ],
    });
    expect(r.c2pa).toEqual(["word/media/a.png"]);
    expect(r.images).toEqual([
      { index: 1, page: 1, keepable: true, tags: ["exif", "c2pa"] },
      { index: 2, page: 1, keepable: true, tags: [] },
    ]);
  });

  test("可疑字句", async () => {
    const r = await inspect({ body: p("正常段落") + p("本文由 AI 協助生成。") + p("Drafted with ChatGPT") });
    expect(r.phrases.map((x) => [x.paragraphId, x.match])).toEqual([
      ["p2", "由 AI 協助生成"],
      ["p3", "ChatGPT"],
    ]);
  });

  test("隱形字元統計", async () => {
    const r = await inspect({ body: p("A\u200BB\u200BC") + p("👨\u200D👩") });
    expect(r.hidden).toEqual({ removed: { "U+200B": 2 }, replaced: {}, kept: 1 });
  });
});
```

- [ ] **Step 2：行測試，確認失敗**

Run: `npx vitest run tests/inspect-docx.test.ts`
Expected: FAIL，`Failed to resolve import "../src/core/inspect/docx"`

- [ ] **Step 3：實作 `src/core/inspect/bytes.ts`**

`src/core/inspect/bytes.ts`

```ts
// 喺 bytes 入面搜尋 ASCII 標記，用嚟認 EXIF / XMP / C2PA

function indexOf(hay: Uint8Array, needle: string): number {
  const n = Array.from(needle, (c) => c.charCodeAt(0));
  outer: for (let i = 0; i <= hay.length - n.length; i++) {
    for (let j = 0; j < n.length; j++) if (hay[i + j] !== n[j]) continue outer;
    return i;
  }
  return -1;
}

export function containsAscii(hay: Uint8Array, needle: string): boolean {
  return indexOf(hay, needle) >= 0;
}

export function hasC2pa(bytes: Uint8Array): boolean {
  // JUMBF 盒子類型 "jumb"、C2PA 標籤 "c2pa"、PNG 嘅 C2PA chunk "caBX"
  return containsAscii(bytes, "c2pa") || containsAscii(bytes, "jumb") || containsAscii(bytes, "caBX");
}

export function imageTags(bytes: Uint8Array): ("exif" | "xmp" | "c2pa")[] {
  const tags: ("exif" | "xmp" | "c2pa")[] = [];
  if (containsAscii(bytes, "Exif\0\0") || containsAscii(bytes, "eXIf")) tags.push("exif");
  if (containsAscii(bytes, "http://ns.adobe.com/xap/1.0/") || containsAscii(bytes, "XML:com.adobe.xmp")) tags.push("xmp");
  if (hasC2pa(bytes)) tags.push("c2pa");
  return tags;
}
```

- [ ] **Step 4：實作 `src/core/inspect/phrases.ts`**

`src/core/inspect/phrases.ts`

```ts
import type { CmDocument, PhraseItem } from "../model";
import { allParagraphs, paragraphText } from "../walk";

// 可疑字句：中英文，唔分大細楷
export const PHRASE_PATTERNS: RegExp[] = [
  /由\s*AI\s*(?:協助|輔助)?\s*(?:生成|產生|撰寫|製作)/i,
  /AI\s*(?:協助|輔助)/i,
  /人工智能\s*(?:生成|產生|協助)/,
  /人工智慧\s*(?:生成|產生|協助)/,
  /generated\s+(?:by|with|using)\b/i,
  /\bAI[-\s]generated\b/i,
  /\bwritten\s+(?:by|with)\s+AI\b/i,
  /\b(?:ChatGPT|OpenAI|Claude|Anthropic|Gemini|Bard|Copilot|DeepSeek|Perplexity|Grok|Llama)\b/i,
];

export function findPhrases(doc: CmDocument): PhraseItem[] {
  const out: PhraseItem[] = [];
  for (const p of allParagraphs(doc)) {
    const text = paragraphText(p);
    for (const re of PHRASE_PATTERNS) {
      const m = re.exec(text);
      if (m) {
        out.push({ paragraphId: p.id, text: text.length > 160 ? text.slice(0, 160) + "…" : text, match: m[0] });
        break;
      }
    }
  }
  return out;
}
```

- [ ] **Step 5：實作 `src/core/inspect/metadata.ts`**

`src/core/inspect/metadata.ts`

```ts
import type { MetadataItem } from "../model";
import type { OoxmlPackage } from "../ooxml/package";
import { attr, kids, NS, type Element } from "../ooxml/xml";

function leaves(el: Element): Element[] {
  const children = kids(el);
  return children.length === 0 ? [el] : children.flatMap(leaves);
}

/** Word / PowerPoint 共用：docProps 同 customXml */
export async function readOoxmlMetadata(pkg: OoxmlPackage): Promise<MetadataItem[]> {
  const out: MetadataItem[] = [];
  for (const part of ["docProps/core.xml", "docProps/app.xml"]) {
    const root = await pkg.xml(part);
    if (!root) continue;
    for (const leaf of kids(root).flatMap(leaves)) {
      const value = (leaf.textContent ?? "").trim();
      if (value) out.push({ part, field: leaf.nodeName, value });
    }
  }
  const custom = await pkg.xml("docProps/custom.xml");
  if (custom) {
    for (const prop of kids(custom, NS.custom, "property")) {
      out.push({ part: "docProps/custom.xml", field: attr(prop, null, "name") ?? "?", value: (prop.textContent ?? "").trim() });
    }
  }
  for (const path of pkg.paths()) {
    if (/^customXml\/item\d+\.xml$/.test(path)) {
      const text = ((await pkg.text(path)) ?? "").replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim();
      out.push({ part: path, field: "customXml", value: text.slice(0, 120) });
    }
  }
  return out;
}
```

- [ ] **Step 6：實作 `src/core/inspect/visible-docx.ts`**

`src/core/inspect/visible-docx.ts`

```ts
import type { VisibleItem } from "../model";
import type { OoxmlPackage } from "../ooxml/package";
import { attr, desc, NS, type Element } from "../ooxml/xml";

/** 唔計 mc:Fallback 入面嘅嘢，避免同 mc:Choice 重複 */
function inFallback(el: Element): boolean {
  for (let n = el.parentNode; n; n = n.parentNode) {
    const e = n as Element;
    if (e.namespaceURI === NS.mc && e.localName === "Fallback") return true;
  }
  return false;
}

function boxText(el: Element): string | undefined {
  const t = desc(el, NS.w, "t").map((x) => x.textContent ?? "").join("").trim();
  return t || undefined;
}

function scanVml(part: string, root: Element, inHeader: boolean): VisibleItem[] {
  const out: VisibleItem[] = [];
  for (const shape of desc(root, NS.v, "shape")) {
    if (inFallback(shape)) continue;
    const id = attr(shape, null, "id") ?? "";
    const textpath = desc(shape, NS.v, "textpath")[0];
    const text = attr(textpath, null, "string");
    if (id.startsWith("PowerPlusWaterMarkObject")) out.push({ code: "wordTextWatermark", part, text });
    else if (id.startsWith("WordPictureWatermark")) out.push({ code: "wordPictureWatermark", part });
    else if (textpath) out.push({ code: inHeader ? "headerShape" : "textArt", part, text });
    else if (inHeader) {
      out.push(desc(shape, NS.v, "imagedata").length > 0 ? { code: "headerImage", part } : { code: "headerShape", part, text: boxText(shape) });
    }
  }
  return out;
}

function scanDrawings(part: string, root: Element, inHeader: boolean): VisibleItem[] {
  const out: VisibleItem[] = [];
  for (const d of desc(root, NS.w, "drawing")) {
    if (inFallback(d)) continue;
    const anchor = desc(d, NS.wp, "anchor")[0];
    const behind = ["1", "true"].includes(attr(anchor, null, "behindDoc") ?? "");
    const isPicture = desc(d, NS.a, "blip").length > 0;
    if (inHeader) out.push(isPicture ? { code: "headerImage", part } : { code: "headerShape", part, text: boxText(d) });
    else if (behind) out.push({ code: "behindText", part, text: boxText(d) });
  }
  return out;
}

export async function findVisibleDocx(pkg: OoxmlPackage): Promise<VisibleItem[]> {
  const out: VisibleItem[] = [];
  for (const part of pkg.paths()) {
    const isHeader = /^word\/(header|footer)\d*\.xml$/.test(part);
    const isBody = part === "word/document.xml";
    if (!isHeader && !isBody) continue;
    const root = await pkg.xml(part);
    if (!root) continue;
    out.push(...scanVml(part, root, isHeader), ...scanDrawings(part, root, isHeader));
    if (isBody && desc(root, NS.w, "background").length > 0) out.push({ code: "pageBackground", part });
  }
  return out;
}
```

- [ ] **Step 7：實作 `src/core/inspect/docx.ts`**

`src/core/inspect/docx.ts`

```ts
import { cleanParagraphSummary } from "../clean/apply-choices";
import { emptySummary, mergeSummary } from "../clean/hidden-chars";
import { KEEPABLE_IMAGE_TYPES, type CmDocument, type ImageItem, type Report } from "../model";
import type { OoxmlPackage } from "../ooxml/package";
import { allImages, allParagraphs } from "../walk";
import { hasC2pa, imageTags } from "./bytes";
import { readOoxmlMetadata } from "./metadata";
import { findPhrases } from "./phrases";
import { findVisibleDocx } from "./visible-docx";

export async function inspectDocx(pkg: OoxmlPackage, doc: CmDocument): Promise<Report> {
  const c2pa: string[] = [];
  for (const path of pkg.paths()) {
    const bytes = await pkg.bytes(path);
    if (bytes && hasC2pa(bytes)) c2pa.push(path);
  }

  const images: ImageItem[] = allImages(doc).map((img) => ({
    index: img.index,
    page: img.page,
    keepable: KEEPABLE_IMAGE_TYPES.has(img.contentType) && !!doc.media[img.mediaPath],
    tags: doc.media[img.mediaPath] ? imageTags(doc.media[img.mediaPath]) : [],
  }));

  let hidden = emptySummary();
  for (const p of allParagraphs(doc)) hidden = mergeSummary(hidden, cleanParagraphSummary(p));

  return {
    fileKind: "docx",
    visible: await findVisibleDocx(pkg),
    metadata: await readOoxmlMetadata(pkg),
    c2pa,
    hidden,
    phrases: findPhrases(doc),
    images,
  };
}
```

- [ ] **Step 8：行測試，確認通過**

Run: `npx vitest run tests/inspect-docx.test.ts`
Expected: `Tests  7 passed (7)`

- [ ] **Step 9：Commit（由用家做）**

```bash
git add src/core/inspect tests/inspect-docx.test.ts
git commit -m "feat(core): inspect docx for watermarks, metadata, C2PA and hidden text"
```

---

### Task 8：Word 輸出器

**Files:**
- Create: `src/core/writers/docx.ts`
- Test: `tests/write-docx.test.ts`

**Interfaces:**
- Consumes: Task 2 型別；測試用 Task 5、Task 6
- Produces: `writeDocx(doc: CmDocument, opts: BuildOptions): Promise<Uint8Array>`

重點：
- 每個字嘅格式全部寫明（顏色冇設定就寫 `auto`），因為 docx 套件嘅預設標題樣式係藍色
- 段落間距、縮排、行距都明確寫出
- 佔位框 = 同原圖一樣大嘅虛線單格表格（EMU ÷ 635 = twip）
- 註腳只接受段落，表格會攤平
- `creator`、`title`、`subject`、`description`、`keywords`、`lastModifiedBy` 全部係空字串

- [ ] **Step 1：寫失敗測試**（讀 → 寫 → 再讀，前後內容清單一樣）

`tests/write-docx.test.ts`

```ts
import { describe, expect, test } from "vitest";
import { applyChoices } from "../src/core/clean/apply-choices";
import type { Block, CmDocument, Paragraph, Table } from "../src/core/model";
import { OoxmlPackage } from "../src/core/ooxml/package";
import { readDocx } from "../src/core/readers/docx";
import { writeDocx } from "../src/core/writers/docx";
import { allParagraphs, paragraphText } from "../src/core/walk";
import { buildDocx, imageP, p, PNG_1PX } from "./fixtures/docx-builder";

const OPTS = { placeholderText: "圖片 {n}：請重新插入" };

/** 去掉 id 同 numId 呢類「編號」，淨係比較內容同格式 */
function normalise(blocks: Block[]): unknown {
  return blocks.map((b) => {
    if (b.type === "paragraph") {
      const { id: _id, list, ...rest } = b;
      return { ...rest, list: list ? { level: list.level } : undefined };
    }
    if (b.type === "table") return { ...b, rows: b.rows.map((r) => r.map((c) => ({ ...c, blocks: normalise(c.blocks) }))) };
    if (b.type === "image") {
      const { mediaPath: _m, ...rest } = b;
      return rest;
    }
    return b;
  });
}

async function roundTrip(doc: CmDocument): Promise<CmDocument> {
  return (await readDocx(await writeDocx(doc, OPTS))).doc;
}

const RICH = {
  styles: `<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:pPr><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:sz w:val="28"/></w:rPr></w:style>`,
  numbering: `<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl><w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%2)"/><w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>`,
  links: { rIdL: "https://example.com/" },
  footnotes: `<w:footnote w:id="1">${p("註腳")}</w:footnote>`,
  header: p("頁首文字"),
  footer: p("頁尾文字", { pPr: `<w:jc w:val="center"/>` }),
  body:
    p("第二級標題", { style: "Heading2" }) +
    p("紅色粗斜底線 Arial 14pt", {
      rPr: `<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Microsoft JhengHei"/><w:b/><w:i/><w:u w:val="single"/><w:strike/><w:color w:val="C00000"/><w:sz w:val="28"/><w:highlight w:val="yellow"/>`,
    }) +
    p("置中、縮排、行距", { pPr: `<w:jc w:val="center"/><w:ind w:left="567" w:firstLine="425"/><w:spacing w:before="120" w:after="240" w:line="360" w:lineRule="auto"/>` }) +
    `<w:p><w:r><w:t>上標</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:t>2</w:t></w:r><w:r><w:t>a</w:t><w:tab/><w:t>b</w:t><w:br/><w:t>c</w:t></w:r><w:hyperlink r:id="rIdL"><w:r><w:t>連結</w:t></w:r></w:hyperlink><w:r><w:footnoteReference w:id="1"/></w:r></w:p>` +
    p("項目一", { pPr: `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>` }) +
    p("子項目", { pPr: `<w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr>` }) +
    `<w:tbl><w:tblPr><w:tblBorders><w:top w:val="single"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="3000"/></w:tblGrid>
      <w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr>${p("合併")}</w:tc></w:tr>
      <w:tr><w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr>${p("直")}</w:tc><w:tc>${p("C")}</w:tc></w:tr>
      <w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr>${p("")}</w:tc><w:tc>${p("D")}</w:tc></w:tr></w:tbl>` +
    `<w:p><w:r><w:br w:type="page"/></w:r></w:p>` +
    p("中文 English 日本語 한국어 العربية"),
};

describe("writeDocx：讀 → 寫 → 再讀，格式唔走樣", () => {
  test("段落、文字格式、列表、表格、頁首頁尾、註腳", async () => {
    const { doc } = await readDocx(await buildDocx(RICH));
    const clean = applyChoices(doc, { removeParagraphs: [], keepImages: [] });
    const again = await roundTrip(clean);
    expect(normalise(again.body)).toEqual(normalise(clean.body));
    expect(normalise(again.header)).toEqual(normalise(clean.header));
    expect(normalise(again.footer)).toEqual(normalise(clean.footer));
    expect(again.page).toEqual(clean.page);
    expect(Object.values(again.footnotes).map((b) => allParagraphs({ ...again, body: b, header: [], footer: [], footnotes: {} }).map(paragraphText))).toEqual([["註腳"]]);
    const listId = (again.body.find((b) => b.type === "paragraph" && b.list) as Paragraph).list!.numId;
    expect(again.lists[listId].slice(0, 2)).toEqual(clean.lists["1"].slice(0, 2));
  });
});

describe("writeDocx：圖片", () => {
  const withImages = () =>
    buildDocx({
      body: imageP("rIdA", 1905000, 952500) + imageP("rIdB", 952500, 952500),
      media: [
        { rId: "rIdA", name: "a.png", bytes: PNG_1PX },
        { rId: "rIdB", name: "b.png", bytes: PNG_1PX },
      ],
    });

  test("保留原圖：放返入新檔案，尺寸一樣", async () => {
    const { doc } = await readDocx(await withImages());
    const again = await roundTrip(applyChoices(doc, { removeParagraphs: [], keepImages: [1] }));
    const img = again.body.find((b) => b.type === "image")!;
    expect(img).toMatchObject({ widthEmu: 1905000, heightEmu: 952500, contentType: "png" });
    expect(Object.values(again.media)).toEqual([PNG_1PX]);
  });

  test("佔位框：同原圖一樣大嘅框，寫住提示字句", async () => {
    const { doc } = await readDocx(await withImages());
    const out = await writeDocx(applyChoices(doc, { removeParagraphs: [], keepImages: [] }), OPTS);
    const again = (await readDocx(out)).doc;
    const tables = again.body.filter((b): b is Table => b.type === "table");
    expect(tables).toHaveLength(2);
    expect(tables[0].columnWidths).toEqual([3000]); // 1905000 EMU ÷ 635
    expect(paragraphText(tables[0].rows[0][0].blocks[0] as Paragraph)).toBe("圖片 1：請重新插入");
    const pkg = await OoxmlPackage.open(out);
    expect(pkg.paths().filter((x) => x.startsWith("word/media/"))).toEqual([]);
  });
});

test("writeDocx：新檔案 metadata 係中性", async () => {
  const { doc } = await readDocx(await buildDocx({ body: p("x"), core: { "dc:creator": "ChatGPT" } }));
  const pkg = await OoxmlPackage.open(await writeDocx(doc, OPTS));
  const core = (await pkg.text("docProps/core.xml"))!;
  expect(core).not.toContain("ChatGPT");
  expect(core).not.toMatch(/<dc:creator>[^<]/);
  expect(core).not.toMatch(/<cp:lastModifiedBy>[^<]/);
});
```

- [ ] **Step 2：行測試，確認失敗**

Run: `npx vitest run tests/write-docx.test.ts`
Expected: FAIL，`Failed to resolve import "../src/core/writers/docx"`

- [ ] **Step 3：實作**

`src/core/writers/docx.ts`

```ts
import {
  AlignmentType,
  BorderStyle,
  Document,
  ExternalHyperlink,
  FootnoteReferenceRun,
  Footer,
  Header,
  HeadingLevel,
  ImageRun,
  LevelFormat,
  LineRuleType,
  Packer,
  PageBreak,
  PageOrientation,
  Paragraph as DocxParagraph,
  Tab,
  Table as DocxTable,
  TableCell as DocxTableCell,
  TableLayoutType,
  TableRow,
  TextRun,
  UnderlineType,
  VerticalAlignTable,
  WidthType,
  type ILevelsOptions,
  type ParagraphChild,
} from "docx";
import type { Block, BuildOptions, CmDocument, ImageSlot, Paragraph, RunStyle, Table, TextInline } from "../model";

const EMU_PER_TWIP = 635;
const EMU_PER_PX = 9525;

const HEADINGS = [
  undefined,
  HeadingLevel.HEADING_1,
  HeadingLevel.HEADING_2,
  HeadingLevel.HEADING_3,
  HeadingLevel.HEADING_4,
  HeadingLevel.HEADING_5,
  HeadingLevel.HEADING_6,
] as const;

const ALIGN = {
  left: AlignmentType.LEFT,
  center: AlignmentType.CENTER,
  right: AlignmentType.RIGHT,
  justify: AlignmentType.JUSTIFIED,
} as const;

const LINE_RULE = { auto: LineRuleType.AUTO, exact: LineRuleType.EXACT, atLeast: LineRuleType.AT_LEAST } as const;

/** 每個字嘅格式都寫明，唔依賴 docx 套件預設樣式（預設標題係藍色） */
function runOptions(style: RunStyle) {
  return {
    bold: style.bold,
    italics: style.italic,
    strike: style.strike,
    underline: style.underline ? { type: UnderlineType.SINGLE } : undefined,
    color: style.color ?? "auto",
    size: Math.round(style.size * 2),
    font: style.font || style.fontEastAsia
      ? { ascii: style.font, hAnsi: style.font, eastAsia: style.fontEastAsia ?? style.font }
      : undefined,
    highlight: style.highlight as never,
    superScript: style.vertAlign === "superscript",
    subScript: style.vertAlign === "subscript",
  };
}

function textRun(inl: TextInline, breakBefore: number): TextRun {
  const parts = inl.text.split("\t");
  const children: (string | Tab)[] = [];
  parts.forEach((part, i) => {
    if (i > 0) children.push(new Tab());
    if (part) children.push(part);
  });
  return new TextRun({ ...runOptions(inl.style), children, break: breakBefore || undefined });
}

function paragraphChildren(p: Paragraph): ParagraphChild[] {
  const out: ParagraphChild[] = [];
  let pendingBreaks = 0;
  for (const inl of p.inlines) {
    if (inl.type === "break") {
      pendingBreaks++;
    } else if (inl.type === "footnoteRef") {
      if (pendingBreaks) out.push(new TextRun({ break: pendingBreaks }));
      pendingBreaks = 0;
      out.push(new FootnoteReferenceRun(inl.id));
    } else {
      const run = textRun(inl, pendingBreaks);
      pendingBreaks = 0;
      out.push(inl.link ? new ExternalHyperlink({ link: inl.link, children: [run] }) : run);
    }
  }
  if (pendingBreaks) out.push(new TextRun({ break: pendingBreaks }));
  return out;
}

class Writer {
  constructor(
    private doc: CmDocument,
    private opts: BuildOptions,
  ) {}

  paragraph(p: Paragraph): DocxParagraph {
    return new DocxParagraph({
      heading: HEADINGS[p.heading],
      alignment: ALIGN[p.align],
      indent:
        p.indentFirstLine < 0
          ? { left: p.indentLeft, hanging: -p.indentFirstLine }
          : { left: p.indentLeft, firstLine: p.indentFirstLine },
      spacing: {
        before: p.spacingBefore,
        after: p.spacingAfter,
        line: p.lineSpacing.line,
        lineRule: LINE_RULE[p.lineSpacing.rule],
      },
      numbering: p.list && this.doc.lists[p.list.numId] ? { reference: `list-${p.list.numId}`, level: p.list.level } : undefined,
      children: paragraphChildren(p),
    });
  }

  image(img: ImageSlot): DocxParagraph | DocxTable {
    const bytes = this.doc.media[img.mediaPath];
    if (img.keep && bytes && img.contentType !== "other") {
      return new DocxParagraph({
        children: [
          new ImageRun({
            type: img.contentType,
            data: bytes,
            transformation: {
              width: Math.round(img.widthEmu / EMU_PER_PX),
              height: Math.round(img.heightEmu / EMU_PER_PX),
            },
          }),
        ],
      });
    }
    // 佔位框：同原圖一樣大嘅虛線框
    const dashed = { style: BorderStyle.DASHED, size: 8, color: "999999" };
    const width = Math.max(720, Math.round(img.widthEmu / EMU_PER_TWIP));
    const height = Math.max(360, Math.round(img.heightEmu / EMU_PER_TWIP));
    return new DocxTable({
      layout: TableLayoutType.FIXED,
      width: { size: width, type: WidthType.DXA },
      columnWidths: [width],
      borders: { top: dashed, bottom: dashed, left: dashed, right: dashed, insideHorizontal: dashed, insideVertical: dashed },
      rows: [
        new TableRow({
          height: { value: height, rule: "exact" },
          children: [
            new DocxTableCell({
              width: { size: width, type: WidthType.DXA },
              verticalAlign: VerticalAlignTable.CENTER,
              children: [
                new DocxParagraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({ text: this.opts.placeholderText.replace("{n}", String(img.index)), color: "777777", size: 20 }),
                  ],
                }),
              ],
            }),
          ],
        }),
      ],
    });
  }

  table(t: Table): DocxTable {
    const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
    const single = { style: BorderStyle.SINGLE, size: 4, color: "000000" };
    const b = t.borders ? single : none;
    const widthOf = (startCol: number, span: number) =>
      t.columnWidths.slice(startCol, startCol + span).reduce((a, x) => a + x, 0);
    return new DocxTable({
      layout: TableLayoutType.FIXED,
      width: { size: t.columnWidths.reduce((a, x) => a + x, 0), type: WidthType.DXA },
      columnWidths: t.columnWidths,
      borders: { top: b, bottom: b, left: b, right: b, insideHorizontal: b, insideVertical: b },
      rows: t.rows.map((row) => {
        let col = 0;
        return new TableRow({
          children: row.map((cell) => {
            const w = widthOf(col, cell.colSpan);
            col += cell.colSpan;
            const children = this.blocks(cell.blocks);
            // Word 規定儲存格一定要以段落結尾
            if (children.length === 0 || !(children.at(-1) instanceof DocxParagraph)) children.push(new DocxParagraph({}));
            return new DocxTableCell({
              width: { size: w, type: WidthType.DXA },
              columnSpan: cell.colSpan > 1 ? cell.colSpan : undefined,
              rowSpan: cell.rowSpan > 1 ? cell.rowSpan : undefined,
              children,
            });
          }),
        });
      }),
    });
  }

  blocks(blocks: Block[]): (DocxParagraph | DocxTable)[] {
    return blocks.map((b) => {
      if (b.type === "paragraph") return this.paragraph(b);
      if (b.type === "table") return this.table(b);
      if (b.type === "image") return this.image(b);
      return new DocxParagraph({ children: [new PageBreak()] });
    });
  }

  /** 註腳只接受段落：表格攤平做段落 */
  footnoteParagraphs(blocks: Block[]): DocxParagraph[] {
    const out: DocxParagraph[] = [];
    for (const b of blocks) {
      if (b.type === "paragraph") out.push(this.paragraph(b));
      else if (b.type === "table") for (const row of b.rows) for (const c of row) out.push(...this.footnoteParagraphs(c.blocks));
    }
    return out;
  }

  numberingConfig() {
    return Object.entries(this.doc.lists).map(([numId, levels]) => ({
      reference: `list-${numId}`,
      levels: levels.map(
        (l, i): ILevelsOptions => ({
          level: i,
          format: (l.format === "none" ? LevelFormat.NONE : l.format) as ILevelsOptions["format"],
          text: l.text,
          start: l.start,
          alignment: AlignmentType.START,
          style: { paragraph: { indent: { left: l.indentLeft, hanging: l.hanging } } },
        }),
      ),
    }));
  }

  build(): Document {
    const { page } = this.doc;
    const footnotes: Record<number, { children: DocxParagraph[] }> = {};
    for (const [id, blocks] of Object.entries(this.doc.footnotes)) {
      footnotes[Number(id)] = { children: this.footnoteParagraphs(blocks) };
    }
    const header = this.blocks(this.doc.header);
    const footer = this.blocks(this.doc.footer);
    return new Document({
      // 中性 metadata：所有欄位留空
      creator: "",
      title: "",
      subject: "",
      description: "",
      keywords: "",
      lastModifiedBy: "",
      revision: 1,
      numbering: { config: this.numberingConfig() },
      footnotes,
      sections: [
        {
          properties: {
            page: {
              size: {
                width: page.widthTwip,
                height: page.heightTwip,
                orientation: page.landscape ? PageOrientation.LANDSCAPE : PageOrientation.PORTRAIT,
              },
              margin: page.margins,
            },
          },
          headers: header.length ? { default: new Header({ children: header }) } : undefined,
          footers: footer.length ? { default: new Footer({ children: footer }) } : undefined,
          children: this.blocks(this.doc.body),
        },
      ],
    });
  }
}

export async function writeDocx(doc: CmDocument, opts: BuildOptions): Promise<Uint8Array> {
  const blob = await Packer.toBlob(new Writer(doc, opts).build());
  return new Uint8Array(await blob.arrayBuffer());
}
```

- [ ] **Step 4：行測試，確認通過**

Run: `npx vitest run tests/write-docx.test.ts`
Expected: `Tests  4 passed (4)`

- [ ] **Step 5：Commit（由用家做）**

```bash
git add src/core/writers/docx.ts tests/write-docx.test.ts
git commit -m "feat(core): write a fresh docx from the content model"
```

---

### Task 9：新檔案自我檢查

**Files:**
- Create: `src/core/verify/self-check-docx.ts`
- Test: `tests/self-check.test.ts`

**Interfaces:**
- Consumes: Task 5 `readDocx`、Task 7 `inspectDocx` / `readOoxmlMetadata` / `findVisibleDocx` / `hasC2pa`
- Produces:
  - `selfCheckDocx(bytes: Uint8Array, keptImages: number): Promise<SelfCheckResult>`
  - `interface SelfCheckResult { ok: boolean; problems: Problem[] }`
  - `interface Problem { code: ProblemCode; detail: string }`
  - `type ProblemCode = "unreadable" | "unexpectedPart" | "unexpectedMedia" | "metadata" | "visible" | "c2pa" | "hiddenChars"`

- [ ] **Step 1：寫失敗測試**

`tests/self-check.test.ts`

```ts
import JSZip from "jszip";
import { expect, test } from "vitest";
import { applyChoices } from "../src/core/clean/apply-choices";
import { readDocx } from "../src/core/readers/docx";
import { selfCheckDocx } from "../src/core/verify/self-check-docx";
import { writeDocx } from "../src/core/writers/docx";
import { buildDocx, imageP, p, PNG_1PX, pngWith, wordTextWatermark } from "./fixtures/docx-builder";

const OPTS = { placeholderText: "圖片 {n}" };

/** 「乜都有」嘅原檔：水印、logo、metadata、C2PA、隱形字元 */
const dirty = () =>
  buildDocx({
    body: p("內文\u200B\u202F") + imageP("rIdA"),
    media: [{ rId: "rIdA", name: "a.png", bytes: pngWith("c2pa") }],
    header: p("頁首") + wordTextWatermark("AI Generated") + imageP("rIdLogo"),
    headerMedia: [{ rId: "rIdLogo", name: "logo.png", bytes: PNG_1PX }],
    core: { "dc:creator": "ChatGPT" },
    app: { Application: "AI Writer" },
    custom: { Source: "Gemini" },
    customXml: "<x>ai</x>",
  });

test("原檔處理之後，自我檢查通過", async () => {
  const { doc } = await readDocx(await dirty());
  const out = await writeDocx(applyChoices(doc, { removeParagraphs: [], keepImages: [] }), OPTS);
  expect(await selfCheckDocx(out, 0)).toEqual({ ok: true, problems: [] });
});

test("保留原圖都通過（圖片入面嘅 C2PA 已經喺報告提示過）", async () => {
  const { doc } = await readDocx(await dirty());
  const out = await writeDocx(applyChoices(doc, { removeParagraphs: [], keepImages: [1] }), OPTS);
  expect((await selfCheckDocx(out, 1)).ok).toBe(true);
});

test("原檔本身唔會通過，每類問題都捉到", async () => {
  const r = await selfCheckDocx(await dirty(), 0);
  expect(r.ok).toBe(false);
  expect(new Set(r.problems.map((x) => x.code))).toEqual(
    new Set(["unexpectedPart", "unexpectedMedia", "metadata", "visible", "hiddenChars"]),
  );
});

test("輸出器以外多咗一個檔案就唔通過", async () => {
  const { doc } = await readDocx(await buildDocx({ body: p("x") }));
  const zip = await JSZip.loadAsync(await writeDocx(doc, OPTS));
  zip.file("word/vbaProject.bin", "x");
  zip.file("word/settings2.xml", "<x/>");
  const r = await selfCheckDocx(await zip.generateAsync({ type: "uint8array" }), 0);
  expect(r.problems).toEqual([
    { code: "unexpectedPart", detail: "word/vbaProject.bin" },
    { code: "unexpectedPart", detail: "word/settings2.xml" },
  ]);
});

test("損壞檔案", async () => {
  const r = await selfCheckDocx(new Uint8Array([1, 2, 3]), 0);
  expect(r.problems[0].code).toBe("unreadable");
});
```

- [ ] **Step 2：行測試，確認失敗**

Run: `npx vitest run tests/self-check.test.ts`
Expected: FAIL，`Failed to resolve import "../src/core/verify/self-check-docx"`

- [ ] **Step 3：實作**

`src/core/verify/self-check-docx.ts`

```ts
import { hasC2pa } from "../inspect/bytes";
import { inspectDocx } from "../inspect/docx";
import { readOoxmlMetadata } from "../inspect/metadata";
import { findVisibleDocx } from "../inspect/visible-docx";
import { readDocx } from "../readers/docx";

export type ProblemCode = "unreadable" | "unexpectedPart" | "unexpectedMedia" | "metadata" | "visible" | "c2pa" | "hiddenChars";

export interface Problem {
  code: ProblemCode;
  detail: string;
}

export interface SelfCheckResult {
  ok: boolean;
  problems: Problem[];
}

/** 輸出器應該產生嘅檔案，其他一律當有問題 */
const ALLOWED_PARTS: RegExp[] = [
  /^\[Content_Types\]\.xml$/,
  /^_rels\/\.rels$/,
  /^docProps\/(core|app|custom)\.xml$/,
  /^word\/(document|styles|settings|fontTable|numbering|footnotes|endnotes)\.xml$/,
  /^word\/(header|footer)\d+\.xml$/,
  /^word\/theme\/theme1\.xml$/,
  /^word\/media\/[^/]+$/,
  /^word\/_rels\/[^/]+\.xml\.rels$/,
];

/** 只准有建立時間、修改時間同版本號 */
const ALLOWED_METADATA = new Set(["cp:revision", "dcterms:created", "dcterms:modified"]);

export async function selfCheckDocx(bytes: Uint8Array, keptImages: number): Promise<SelfCheckResult> {
  const problems: Problem[] = [];
  let read: Awaited<ReturnType<typeof readDocx>>;
  try {
    read = await readDocx(bytes);
  } catch (e) {
    return { ok: false, problems: [{ code: "unreadable", detail: String(e) }] };
  }
  const { doc, pkg } = read;

  const paths = pkg.paths();
  for (const path of paths) {
    if (!ALLOWED_PARTS.some((re) => re.test(path))) problems.push({ code: "unexpectedPart", detail: path });
  }
  const mediaCount = paths.filter((x) => x.startsWith("word/media/")).length;
  if (mediaCount > keptImages) problems.push({ code: "unexpectedMedia", detail: `${mediaCount} > ${keptImages}` });

  for (const m of await readOoxmlMetadata(pkg)) {
    if (!ALLOWED_METADATA.has(m.field)) problems.push({ code: "metadata", detail: `${m.field}=${m.value}` });
  }
  for (const v of await findVisibleDocx(pkg)) problems.push({ code: "visible", detail: `${v.code} ${v.part}` });

  for (const path of paths) {
    if (path.startsWith("word/media/")) continue; // 用家揀咗保留嘅原圖，報告已經提示過
    const b = await pkg.bytes(path);
    if (b && hasC2pa(b)) problems.push({ code: "c2pa", detail: path });
  }

  const { hidden } = await inspectDocx(pkg, doc);
  const leftover = [...Object.entries(hidden.removed), ...Object.entries(hidden.replaced)];
  for (const [cp, n] of leftover) problems.push({ code: "hiddenChars", detail: `${cp}×${n}` });

  return { ok: problems.length === 0, problems };
}
```

- [ ] **Step 4：行測試，確認通過**

Run: `npx vitest run tests/self-check.test.ts`
Expected: `Tests  5 passed (5)`

- [ ] **Step 5：Commit（由用家做）**

```bash
git add src/core/verify tests/self-check.test.ts
git commit -m "feat(core): self-check generated docx before download"
```

---

### Task 10：處理流程同輸入檢查

**Files:**
- Create: `src/core/pipeline.ts`
- Test: `tests/pipeline.test.ts`

**Interfaces:**
- Consumes: Task 5–9
- Produces:
  - `analyze(name: string, bytes: Uint8Array): Promise<Analysis>`，`Analysis = { doc: CmDocument; report: Report }`
  - `build(name, doc, choices, opts): Promise<BuiltFile>`，`BuiltFile = { name: string; bytes: Uint8Array }`；自我檢查唔通過就 throw `CleanMarkError("selfCheckFailed", problems)`
  - `class CleanMarkError extends Error { code: ErrorCode; problems: Problem[] }`
  - `type ErrorCode = "legacyFormat" | "unsupported" | "comingSoon" | "tooLarge" | "encrypted" | "corrupt" | "selfCheckFailed"`
  - `checkInput(name, bytes): "docx"`、`extensionOf(name)`、`outputName(name, ext)`（`報告.docx` → `報告-clean.docx`）、`MAX_BYTES`

- [ ] **Step 1：寫失敗測試**

`tests/pipeline.test.ts`

```ts
import { describe, expect, test } from "vitest";
import { analyze, build, CleanMarkError, MAX_BYTES, outputName } from "../src/core/pipeline";
import { readDocx } from "../src/core/readers/docx";
import { allParagraphs, paragraphText } from "../src/core/walk";
import { buildDocx, p, wordTextWatermark } from "./fixtures/docx-builder";

const OPTS = { placeholderText: "圖片 {n}" };

async function codeOf(name: string, bytes: Uint8Array): Promise<string> {
  try {
    await analyze(name, bytes);
    return "ok";
  } catch (e) {
    return e instanceof CleanMarkError ? e.code : "other";
  }
}

describe("analyze：輸入檢查", () => {
  test("各種錯誤", async () => {
    const zip = await buildDocx({ body: p("x") });
    expect(await codeOf("a.doc", zip)).toBe("legacyFormat");
    expect(await codeOf("a.ppt", zip)).toBe("legacyFormat");
    expect(await codeOf("a.pptx", zip)).toBe("comingSoon");
    expect(await codeOf("a.pdf", zip)).toBe("comingSoon");
    expect(await codeOf("a.txt", zip)).toBe("unsupported");
    expect(await codeOf("a.docx", new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0]))).toBe("encrypted");
    expect(await codeOf("a.docx", new Uint8Array([1, 2, 3, 4]))).toBe("corrupt");
    expect(await codeOf("a.docx", new Uint8Array([0x50, 0x4b, 0x03, 0x04, 9, 9]))).toBe("corrupt");
    expect(await codeOf("A.DOCX", zip)).toBe("ok");
  });

  test("超過 50 MB", async () => {
    const big = new Uint8Array(MAX_BYTES + 1);
    big.set([0x50, 0x4b, 0x03, 0x04]);
    expect(await codeOf("a.docx", big)).toBe("tooLarge");
  });
});

test("analyze → build：成個流程", async () => {
  const bytes = await buildDocx({
    body: p("保留呢段") + p("Generated by ChatGPT"),
    header: wordTextWatermark("DRAFT"),
    core: { "dc:creator": "AI" },
  });
  const { doc, report } = await analyze("報告.docx", bytes);
  expect(report.visible).toHaveLength(1);
  expect(report.phrases).toHaveLength(1);

  const out = await build("報告.docx", doc, { removeParagraphs: [report.phrases[0].paragraphId], keepImages: [] }, OPTS);
  expect(out.name).toBe("報告-clean.docx");
  const again = (await readDocx(out.bytes)).doc;
  expect(allParagraphs(again).map(paragraphText)).toEqual(["保留呢段"]);
});

test("outputName", () => {
  expect(outputName("a.b.docx", "docx")).toBe("a.b-clean.docx");
  expect(outputName("noext", "docx")).toBe("noext-clean.docx");
});
```

- [ ] **Step 2：行測試，確認失敗**

Run: `npx vitest run tests/pipeline.test.ts`
Expected: FAIL，`Failed to resolve import "../src/core/pipeline"`

- [ ] **Step 3：實作**

`src/core/pipeline.ts`

```ts
import { applyChoices } from "./clean/apply-choices";
import { inspectDocx } from "./inspect/docx";
import type { BuildOptions, Choices, CmDocument, Report } from "./model";
import { readDocx } from "./readers/docx";
import { selfCheckDocx, type Problem } from "./verify/self-check-docx";
import { writeDocx } from "./writers/docx";

export const MAX_BYTES = 50 * 1024 * 1024;

export type ErrorCode =
  | "legacyFormat" // .doc / .ppt
  | "unsupported" // 其他格式
  | "comingSoon" // .pptx / .pdf：之後幾期先支援
  | "tooLarge"
  | "encrypted"
  | "corrupt"
  | "selfCheckFailed";

export class CleanMarkError extends Error {
  constructor(
    public code: ErrorCode,
    public problems: Problem[] = [],
  ) {
    super(code);
  }
}

const ZIP = [0x50, 0x4b, 0x03, 0x04];
const CFB = [0xd0, 0xcf, 0x11, 0xe0]; // 舊 Office 格式、或者加咗密碼嘅新格式

const startsWith = (b: Uint8Array, sig: number[]) => sig.every((x, i) => b[i] === x);

export function extensionOf(name: string): string {
  const i = name.lastIndexOf(".");
  return i < 0 ? "" : name.slice(i + 1).toLowerCase();
}

export function checkInput(name: string, bytes: Uint8Array): "docx" {
  const ext = extensionOf(name);
  if (ext === "doc" || ext === "ppt") throw new CleanMarkError("legacyFormat");
  if (ext === "pptx" || ext === "pdf") throw new CleanMarkError("comingSoon");
  if (ext !== "docx") throw new CleanMarkError("unsupported");
  if (bytes.length > MAX_BYTES) throw new CleanMarkError("tooLarge");
  if (startsWith(bytes, CFB)) throw new CleanMarkError("encrypted");
  if (!startsWith(bytes, ZIP)) throw new CleanMarkError("corrupt");
  return "docx";
}

export interface Analysis {
  doc: CmDocument;
  report: Report;
}

export async function analyze(name: string, bytes: Uint8Array): Promise<Analysis> {
  checkInput(name, bytes);
  try {
    const { doc, pkg } = await readDocx(bytes);
    return { doc, report: await inspectDocx(pkg, doc) };
  } catch (e) {
    if (e instanceof CleanMarkError) throw e;
    throw new CleanMarkError("corrupt");
  }
}

export interface BuiltFile {
  name: string;
  bytes: Uint8Array;
}

export async function build(
  name: string,
  doc: CmDocument,
  choices: Choices,
  opts: BuildOptions,
): Promise<BuiltFile> {
  const finalDoc = applyChoices(doc, choices);
  const bytes = await writeDocx(finalDoc, opts);
  const keptImages = Object.keys(finalDoc.media).length;
  const check = await selfCheckDocx(bytes, keptImages);
  if (!check.ok) throw new CleanMarkError("selfCheckFailed", check.problems);
  return { name: outputName(name, "docx"), bytes };
}

export function outputName(name: string, ext: string): string {
  const base = name.includes(".") ? name.slice(0, name.lastIndexOf(".")) : name;
  return `${base}-clean.${ext}`;
}
```

- [ ] **Step 4：行全部單元測試**

Run: `npm test`
Expected: `Test Files  9 passed (9)`，`Tests  54 passed (54)`

- [ ] **Step 5：Commit（由用家做）**

```bash
git add src/core/pipeline.ts tests/pipeline.test.ts
git commit -m "feat(core): add analyze/build pipeline with input validation"
```

---

### Task 11：介面用嘅純函數（翻譯、列表編號、預覽樣式、報告統計）

**Files:**
- Create: `src/web/i18n/zh-Hant.ts`, `src/web/i18n/en.ts`, `src/web/i18n/index.svelte.ts`, `src/web/list-label.ts`, `src/web/report-summary.ts`, `src/web/preview-style.ts`
- Test: `tests/i18n.test.ts`, `tests/list-label.test.ts`, `tests/report-summary.test.ts`

**Interfaces:**
- Produces:
  - `t(key: MessageKey, params?): string`、`setLang(lang)`、`i18n`（`$state`，`{ lang: Lang }`）、`type Lang = "zh-Hant" | "en"`、`type MessageKey`、`type Messages`
  - `listLabels(blocks: Iterable<Block>, lists): Map<string, string>`（段落 id → 「1.」、「1.a)」、「•」）、`formatNumber(n, format)`
  - `hiddenTotal(report): number`、`findingCount(report): number`（隱形字元當一項）
  - `runCss(style: RunStyle): string`、`paragraphCss(p: Paragraph): string`、`EMU_TO_PX`

- [ ] **Step 1：寫失敗測試**

`tests/i18n.test.ts`

```ts
import { expect, test } from "vitest";
import en from "../src/web/i18n/en";
import zhHant from "../src/web/i18n/zh-Hant";

test("兩種語言嘅 key 一樣", () => {
  expect(Object.keys(en).sort()).toEqual(Object.keys(zhHant).sort());
});

test("參數 {x} 兩種語言都一樣", () => {
  const params = (s: string) => (s.match(/\{\w+\}/g) ?? []).sort();
  for (const k of Object.keys(zhHant) as (keyof typeof zhHant)[]) {
    expect(params(en[k]), k).toEqual(params(zhHant[k]));
  }
});
```

`tests/list-label.test.ts`

```ts
import { expect, test } from "vitest";
import type { Block, ListLevel } from "../src/core/model";
import { formatNumber, listLabels } from "../src/web/list-label";

test("formatNumber", () => {
  expect(formatNumber(4, "upperRoman")).toBe("IV");
  expect(formatNumber(28, "lowerLetter")).toBe("ab");
  expect(formatNumber(3, "decimal")).toBe("3");
});

test("listLabels 計多層編號，上層前進時重設下層", () => {
  const lvl = (format: string, text: string): ListLevel => ({ format, text, start: 1, indentLeft: 0, hanging: 0 });
  const lists = { "1": [lvl("decimal", "%1."), lvl("lowerLetter", "%1.%2)"), lvl("bullet", "▪")] };
  const item = (id: string, level: number): Block => ({
    type: "paragraph", id, heading: 0, align: "left", indentLeft: 0, indentFirstLine: 0,
    spacingBefore: 0, spacingAfter: 0, lineSpacing: { line: 240, rule: "auto" },
    list: { numId: "1", level }, inlines: [],
  });
  const labels = listLabels([item("a", 0), item("b", 1), item("c", 1), item("d", 2), item("e", 0), item("f", 1)], lists);
  expect([...labels.values()]).toEqual(["1.", "1.a)", "1.b)", "▪", "2.", "2.a)"]);
});
```

`tests/report-summary.test.ts`

```ts
import { expect, test } from "vitest";
import type { Report } from "../src/core/model";
import { findingCount, hiddenTotal } from "../src/web/report-summary";

const base: Report = {
  fileKind: "docx", visible: [], metadata: [], c2pa: [],
  hidden: { removed: {}, replaced: {}, kept: 0 }, phrases: [], images: [],
};

test("乾淨文件 = 0 項", () => {
  expect(findingCount(base)).toBe(0);
});

test("隱形字元無論幾多個都當一項", () => {
  const r = {
    ...base,
    visible: [{ code: "headerImage" as const, part: "word/header1.xml" }],
    metadata: [{ part: "docProps/core.xml", field: "dc:creator", value: "x" }],
    hidden: { removed: { "U+200B": 23 }, replaced: { "U+202F": 2 }, kept: 1 },
  };
  expect(hiddenTotal(r)).toBe(25);
  expect(findingCount(r)).toBe(3);
});
```

- [ ] **Step 2：行測試，確認失敗**

Run: `npx vitest run tests/i18n.test.ts tests/list-label.test.ts tests/report-summary.test.ts`
Expected: FAIL，搵唔到 `../src/web/i18n/en` 等模組

- [ ] **Step 3：建立繁中翻譯 `src/web/i18n/zh-Hant.ts`**

`src/web/i18n/zh-Hant.ts`

```ts
const zhHant = {
  tagline: "移除 Word 文件的水印與隱藏資料",
  langSwitch: "English",
  dropTitle: "把 Word 檔案拖到這裡",
  dropButton: "選擇檔案",
  dropHint: "支援 .docx，PowerPoint 與 PDF 即將推出",
  privacy: "檔案只在你的電腦上處理，不會上傳到任何地方",
  step1: "檢查",
  step1Desc: "找出水印與隱藏資料",
  step2: "確認",
  step2Desc: "決定圖片與可疑字句",
  step3: "下載",
  step3Desc: "取得全新、乾淨的檔案",
  analyzing: "正在檢查…",
  changeFile: "換檔案",
  foundN: "找到 {n} 項",
  foundNone: "沒有發現問題",
  reportTitle: "檢查報告",
  visibleTitle: "可見水印",
  metadataTitle: "檔案資料",
  c2paTitle: "Content Credentials (C2PA)",
  hiddenTitle: "隱形字元",
  hiddenRemoved: "會清除 {n} 個",
  hiddenReplaced: "會換成普通空格 {n} 個",
  hiddenKept: "保留 {n} 個（語言需要）",
  none: "沒有",
  statisticalNote: "統計水印（藏在用字規律中）無法檢查。",
  v_wordTextWatermark: "Word 文字水印",
  v_wordPictureWatermark: "Word 圖片水印",
  v_headerShape: "頁首／頁尾的圖形",
  v_headerImage: "頁首／頁尾的圖片",
  v_behindText: "文字後面的圖形",
  v_pageBackground: "頁面背景",
  v_textArt: "文字藝術師 (WordArt)",
  confirmTitle: "需要你確認",
  nothingToConfirm: "沒有需要確認的項目",
  phrasesTitle: "可疑字句",
  phraseRemove: "刪除這段",
  imagesTitle: "圖片",
  imageN: "圖片 {n}",
  pageN: "約第 {n} 頁",
  usePlaceholder: "佔位框",
  keepOriginal: "保留原圖",
  allPlaceholder: "全部用佔位框",
  allKeep: "全部保留",
  keepWarning: "原圖可能帶有水印或隱藏資料，建議重新製作。",
  imageTagged: "帶有 {tags} 標記",
  notKeepable: "這種圖片只能用佔位框",
  previewTitle: "預覽新檔案",
  placeholderText: "圖片 {n}：請重新插入",
  downloadTitle: "下載",
  buildButton: "清除並下載",
  building: "正在製作新檔案…",
  selfCheckPassed: "新檔案自我檢查通過",
  downloadAgain: "再次下載",
  another: "處理另一個檔案",
  err_legacyFormat: "這是舊格式，請先用 Word 另存為 .docx。",
  err_unsupported: "目前只支援 .docx 檔案。",
  err_comingSoon: "PowerPoint 與 PDF 支援即將推出，目前只支援 .docx。",
  err_tooLarge: "檔案太大，上限是 50 MB。",
  err_encrypted: "檔案有密碼保護，請先移除密碼。",
  err_corrupt: "無法開啟這個檔案，試試用原本的軟體重新儲存。",
  err_selfCheckFailed: "新檔案未通過自我檢查，為安全起見不提供下載。",
  err_unknown: "發生未預期的錯誤，請重新整理頁面再試。",
  limitsTitle: "做不到的事",
  limitStatistical: "統計水印：藏在用字規律裡，需要 AI 公司的私人密鑰才能驗證。",
  limitDropped: "Word 註解不會保留；圖表、SmartArt、方程式會變成佔位框。",
  usageRule: "只可用於你有權處理的文件。請勿用來移除他人的版權、機密或其他權利標記。",
};

export default zhHant;
export type Messages = typeof zhHant;
```

- [ ] **Step 4：建立英文翻譯 `src/web/i18n/en.ts`**

`src/web/i18n/en.ts`

```ts
import type { Messages } from "./zh-Hant";

const en: Messages = {
  tagline: "Remove watermarks and hidden data from Word documents",
  langSwitch: "繁中",
  dropTitle: "Drop a Word file here",
  dropButton: "Choose file",
  dropHint: "Supports .docx. PowerPoint and PDF are coming soon",
  privacy: "Your file is processed on your computer and never uploaded",
  step1: "Check",
  step1Desc: "Find watermarks and hidden data",
  step2: "Confirm",
  step2Desc: "Decide on images and flagged text",
  step3: "Download",
  step3Desc: "Get a brand-new, clean file",
  analyzing: "Checking…",
  changeFile: "Change file",
  foundN: "{n} found",
  foundNone: "Nothing found",
  reportTitle: "Report",
  visibleTitle: "Visible watermarks",
  metadataTitle: "File metadata",
  c2paTitle: "Content Credentials (C2PA)",
  hiddenTitle: "Hidden characters",
  hiddenRemoved: "{n} will be removed",
  hiddenReplaced: "{n} will become normal spaces",
  hiddenKept: "{n} kept (needed by the language)",
  none: "None",
  statisticalNote: "Statistical watermarks (hidden in word choice) can't be checked.",
  v_wordTextWatermark: "Word text watermark",
  v_wordPictureWatermark: "Word picture watermark",
  v_headerShape: "Shape in header or footer",
  v_headerImage: "Image in header or footer",
  v_behindText: "Shape behind text",
  v_pageBackground: "Page background",
  v_textArt: "WordArt",
  confirmTitle: "Your call",
  nothingToConfirm: "Nothing to confirm",
  phrasesTitle: "Flagged text",
  phraseRemove: "Remove this paragraph",
  imagesTitle: "Images",
  imageN: "Image {n}",
  pageN: "around page {n}",
  usePlaceholder: "Placeholder",
  keepOriginal: "Keep original",
  allPlaceholder: "All placeholders",
  allKeep: "Keep all",
  keepWarning: "Original images may carry watermarks or hidden data. Recreating them is safer.",
  imageTagged: "Tagged with {tags}",
  notKeepable: "This image type can only be a placeholder",
  previewTitle: "Preview",
  placeholderText: "Image {n}: insert again",
  downloadTitle: "Download",
  buildButton: "Clean and download",
  building: "Building the new file…",
  selfCheckPassed: "New file passed the self-check",
  downloadAgain: "Download again",
  another: "Clean another file",
  err_legacyFormat: "This is an old format. Save it as .docx in Word first.",
  err_unsupported: "Only .docx files are supported right now.",
  err_comingSoon: "PowerPoint and PDF support is coming soon. Only .docx works right now.",
  err_tooLarge: "That file's too big. The limit is 50 MB.",
  err_encrypted: "This file is password-protected. Remove the password first.",
  err_corrupt: "Couldn't open this file. Try saving it again in the app that made it.",
  err_selfCheckFailed: "The new file didn't pass the self-check, so it isn't offered for download.",
  err_unknown: "Something unexpected went wrong. Reload the page and try again.",
  limitsTitle: "What it can't do",
  limitStatistical: "Statistical watermarks: they live in word choice and need the AI company's private key to verify.",
  limitDropped: "Word comments aren't kept. Charts, SmartArt, and equations become placeholders.",
  usageRule: "Only use this on documents you have the right to modify. Don't use it to strip someone else's copyright, confidentiality, or other rights notices.",
};

export default en;
```

- [ ] **Step 5：建立 `src/web/i18n/index.svelte.ts`**

`src/web/i18n/index.svelte.ts`

```ts
import en from "./en";
import zhHant, { type Messages } from "./zh-Hant";

export type Lang = "zh-Hant" | "en";
export type MessageKey = keyof Messages;

const DICTS: Record<Lang, Messages> = { "zh-Hant": zhHant, en };
const STORAGE_KEY = "cleanmark.lang";

function initialLang(): Lang {
  try {
    const saved = localStorage.getItem(STORAGE_KEY);
    if (saved === "zh-Hant" || saved === "en") return saved;
  } catch {
    // 私密瀏覽等情況讀唔到，用瀏覽器語言
  }
  return navigator.language.toLowerCase().startsWith("zh") ? "zh-Hant" : "en";
}

export const i18n = $state({ lang: initialLang() });

export function setLang(lang: Lang): void {
  i18n.lang = lang;
  document.documentElement.lang = lang;
  try {
    localStorage.setItem(STORAGE_KEY, lang);
  } catch {
    // 儲存唔到就算，下次用返預設
  }
}

export function t(key: MessageKey, params: Record<string, string | number> = {}): string {
  let s = DICTS[i18n.lang][key];
  for (const [k, v] of Object.entries(params)) s = s.replaceAll(`{${k}}`, String(v));
  return s;
}
```

- [ ] **Step 6：建立 `src/web/list-label.ts`**

`src/web/list-label.ts`

```ts
import type { Block, ListLevel } from "../core/model";

function roman(n: number): string {
  const table: [number, string][] = [
    [1000, "m"], [900, "cm"], [500, "d"], [400, "cd"], [100, "c"], [90, "xc"],
    [50, "l"], [40, "xl"], [10, "x"], [9, "ix"], [5, "v"], [4, "iv"], [1, "i"],
  ];
  let out = "";
  for (const [v, s] of table) while (n >= v) { out += s; n -= v; }
  return out;
}

function letters(n: number): string {
  let out = "";
  while (n > 0) { n--; out = String.fromCharCode(97 + (n % 26)) + out; n = Math.floor(n / 26); }
  return out;
}

export function formatNumber(n: number, format: string): string {
  switch (format) {
    case "lowerLetter": return letters(n);
    case "upperLetter": return letters(n).toUpperCase();
    case "lowerRoman": return roman(n);
    case "upperRoman": return roman(n).toUpperCase();
    case "decimalZero": return n < 10 ? `0${n}` : String(n);
    default: return String(n);
  }
}

/** 預覽用：計出每個列表段落前面嘅編號或者符號，key = 段落 id */
export function listLabels(blocks: Iterable<Block>, lists: Record<string, ListLevel[]>): Map<string, string> {
  const counters = new Map<string, number[]>();
  const out = new Map<string, string>();
  for (const b of blocks) {
    if (b.type !== "paragraph" || !b.list) continue;
    const levels = lists[b.list.numId];
    if (!levels) continue;
    const { level } = b.list;
    const c = counters.get(b.list.numId) ?? levels.map((l) => l.start - 1);
    c[level] = (c[level] ?? 0) + 1;
    for (let i = level + 1; i < c.length; i++) c[i] = levels[i].start - 1;
    counters.set(b.list.numId, c);
    const def = levels[level];
    out.set(
      b.id,
      def.format === "bullet"
        ? def.text
        : def.format === "none"
          ? ""
          : def.text.replace(/%(\d)/g, (_, d) => formatNumber(Math.max(c[Number(d) - 1] ?? 1, 1), levels[Number(d) - 1]?.format ?? "decimal")),
    );
  }
  return out;
}
```

- [ ] **Step 7：建立 `src/web/report-summary.ts`**

`src/web/report-summary.ts`

```ts
import type { Report } from "../core/model";

export function hiddenTotal(r: Report): number {
  const sum = (x: Record<string, number>) => Object.values(x).reduce((a, b) => a + b, 0);
  return sum(r.hidden.removed) + sum(r.hidden.replaced);
}

/** 檔案資料行顯示嘅「找到 N 項」：隱形字元當一項計 */
export function findingCount(r: Report): number {
  return r.visible.length + r.metadata.length + r.c2pa.length + (hiddenTotal(r) > 0 ? 1 : 0) + r.phrases.length;
}
```

- [ ] **Step 8：建立 `src/web/preview-style.ts`**

`src/web/preview-style.ts`

```ts
import type { Paragraph, RunStyle } from "../core/model";

const HIGHLIGHT: Record<string, string> = {
  yellow: "#ffff00", green: "#00ff00", cyan: "#00ffff", magenta: "#ff00ff", blue: "#0000ff",
  red: "#ff0000", darkBlue: "#000080", darkCyan: "#008080", darkGreen: "#008000",
  darkMagenta: "#800080", darkRed: "#800000", darkYellow: "#808000", darkGray: "#808080",
  lightGray: "#c0c0c0", black: "#000000", white: "#ffffff",
};

const quote = (f: string) => `"${f.replace(/"/g, "")}"`;
const TWIP_TO_PX = 96 / 1440;

export function runCss(s: RunStyle): string {
  const fonts = [s.font, s.fontEastAsia].filter(Boolean).map((f) => quote(f!));
  const deco = [s.underline && "underline", s.strike && "line-through"].filter(Boolean).join(" ");
  return [
    fonts.length ? `font-family:${fonts.join(",")},inherit` : "",
    `font-size:${s.size}pt`,
    `font-weight:${s.bold ? 700 : 400}`,
    `font-style:${s.italic ? "italic" : "normal"}`,
    deco ? `text-decoration:${deco}` : "",
    s.color ? `color:#${s.color}` : "",
    s.highlight && HIGHLIGHT[s.highlight] ? `background:${HIGHLIGHT[s.highlight]}` : "",
    s.vertAlign === "superscript" ? "vertical-align:super;font-size:0.7em" : "",
    s.vertAlign === "subscript" ? "vertical-align:sub;font-size:0.7em" : "",
  ]
    .filter(Boolean)
    .join(";");
}

export function paragraphCss(p: Paragraph): string {
  const line = p.lineSpacing.rule === "auto" ? String(p.lineSpacing.line / 240 * 1.15) : `${(p.lineSpacing.line / 20).toFixed(1)}pt`;
  return [
    `text-align:${p.align}`,
    `margin:${Math.round(p.spacingBefore * TWIP_TO_PX)}px 0 ${Math.round(p.spacingAfter * TWIP_TO_PX)}px ${Math.round(p.indentLeft * TWIP_TO_PX)}px`,
    `text-indent:${Math.round(p.indentFirstLine * TWIP_TO_PX)}px`,
    `line-height:${line}`,
  ].join(";");
}

export const EMU_TO_PX = 96 / 914400;
```

- [ ] **Step 9：行測試，確認通過**

Run: `npx vitest run tests/i18n.test.ts tests/list-label.test.ts tests/report-summary.test.ts`
Expected: `Tests  6 passed (6)`

- [ ] **Step 10：Commit（由用家做）**

```bash
git add src/web/i18n src/web/list-label.ts src/web/report-summary.ts src/web/preview-style.ts tests/i18n.test.ts tests/list-label.test.ts tests/report-summary.test.ts
git commit -m "feat(web): add i18n, list labels, preview styles and report summary"
```

---

### Task 12：網頁介面同背景執行緒

**Files:**
- Create: `index.html`, `public/favicon.svg`, `src/web/main.ts`, `src/web/styles.css`, `src/web/protocol.ts`, `src/web/worker.ts`, `src/web/worker-client.ts`, `src/web/App.svelte`
- Create: `src/web/components/SiteHeader.svelte`, `DropZone.svelte`, `FileBar.svelte`, `Section.svelte`, `ReportPanel.svelte`, `ConfirmPanel.svelte`, `Preview.svelte`, `PreviewBlocks.svelte`, `DownloadPanel.svelte`, `ErrorNotice.svelte`, `SiteFooter.svelte`

**Interfaces:**
- Consumes: Task 6 `applyChoices` / `allBlocks`、Task 10 `analyze` / `build` / `CleanMarkError`、Task 11 全部
- Produces: 可以喺瀏覽器用嘅網頁；`data-testid`：`file-input`、`found-count`、`preview`、`self-check-ok`、`error`（Task 13 會用）

介面測試由 Task 13 嘅 Playwright 負責。呢個 task 用型別檢查、打包同人手試用驗證。

兩個實作陷阱（開發時真係遇過）：
1. `postMessage` 複製唔到 Svelte `$state` 嘅 Proxy，所以 `onbuild` 要先用 `$state.snapshot()` 轉做普通物件
2. 「改咗選擇就作廢舊檔案」嘅 `$effect` 唔可以追蹤 `phase`，否則 phase 變成 `done` 嗰下會即刻被打返去 `ready`；要用 `untrack()`

- [ ] **Step 1：建立 `index.html`、`public/favicon.svg`、`src/web/main.ts`**

`index.html`

```html
<!doctype html>
<html lang="zh-Hant">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <meta name="referrer" content="no-referrer" />
    <link rel="icon" type="image/svg+xml" href="./favicon.svg" />
    <title>CleanMark</title>
  </head>
  <body>
    <div id="app"></div>
    <script type="module" src="/src/web/main.ts"></script>
  </body>
</html>
```

`public/favicon.svg`

```svg
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 32 32"><rect width="32" height="32" rx="8" fill="#1f6f5c"/><path d="M9 17l4.5 4.5L23 11" fill="none" stroke="#fff" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"/></svg>
```

`src/web/main.ts`

```ts
import { mount } from "svelte";
import App from "./App.svelte";
import { i18n } from "./i18n/index.svelte";
import "./styles.css";

document.documentElement.lang = i18n.lang;
mount(App, { target: document.getElementById("app")! });
```

- [ ] **Step 2：建立全域樣式 `src/web/styles.css`**（淺色同深色模式）

`src/web/styles.css`

```css
:root {
  --bg: #f6f5f1;
  --surface: #ffffff;
  --surface-muted: #f1efe8;
  --text: #1f1e1b;
  --text-muted: #5f5e5a;
  --border: #e3e0d6;
  --border-strong: #c9c5b8;
  --accent: #1f6f5c;
  --accent-text: #ffffff;
  --accent-soft: #e1f2ec;
  --accent-soft-text: #0f5040;
  --warn-soft: #faeeda;
  --warn-text: #7a4a0b;
  --ok-soft: #e6f3dc;
  --ok-text: #2c5a10;
  --danger-soft: #fcebeb;
  --danger-text: #8a2424;
  --radius: 10px;
  --radius-lg: 16px;
  color-scheme: light;
}

@media (prefers-color-scheme: dark) {
  :root {
    --bg: #161614;
    --surface: #1f1f1d;
    --surface-muted: #2a2a27;
    --text: #efede6;
    --text-muted: #a9a79f;
    --border: #34332f;
    --border-strong: #4a4944;
    --accent: #5dcaa5;
    --accent-text: #08281f;
    --accent-soft: #123a30;
    --accent-soft-text: #9fe1cb;
    --warn-soft: #3a2a10;
    --warn-text: #f5c77a;
    --ok-soft: #1f3312;
    --ok-text: #b5dd8f;
    --danger-soft: #3a1717;
    --danger-text: #f3a9a9;
    color-scheme: dark;
  }
}

* {
  box-sizing: border-box;
}

body {
  margin: 0;
  background: var(--bg);
  color: var(--text);
  font-family: system-ui, -apple-system, "Segoe UI", "PingFang TC", "Microsoft JhengHei", "Noto Sans TC", sans-serif;
  font-size: 16px;
  line-height: 1.6;
}

button {
  font: inherit;
  cursor: pointer;
  border-radius: var(--radius);
  border: 1px solid var(--border-strong);
  background: var(--surface);
  color: var(--text);
  padding: 8px 14px;
}

button:hover {
  background: var(--surface-muted);
}

button.primary {
  background: var(--accent);
  border-color: var(--accent);
  color: var(--accent-text);
  font-weight: 600;
}

button.primary:hover {
  filter: brightness(1.08);
}

button:focus-visible,
input:focus-visible,
label:focus-within {
  outline: 2px solid var(--accent);
  outline-offset: 2px;
}

.card {
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: var(--radius-lg);
  padding: 20px;
}

.muted {
  color: var(--text-muted);
}

.pill {
  display: inline-flex;
  align-items: center;
  gap: 4px;
  font-size: 13px;
  padding: 2px 10px;
  border-radius: 999px;
  white-space: nowrap;
}

.pill.warn {
  background: var(--warn-soft);
  color: var(--warn-text);
}

.pill.ok {
  background: var(--ok-soft);
  color: var(--ok-text);
}

.sr-only {
  position: absolute;
  width: 1px;
  height: 1px;
  overflow: hidden;
  clip: rect(0 0 0 0);
  white-space: nowrap;
}
```

- [ ] **Step 3：建立背景執行緒 `src/web/protocol.ts`、`src/web/worker.ts`、`src/web/worker-client.ts`**

`src/web/protocol.ts`

```ts
import type { Choices, CmDocument, Report } from "../core/model";
import type { ErrorCode } from "../core/pipeline";
import type { Problem } from "../core/verify/self-check-docx";

export type WorkerRequest =
  | { id: number; type: "analyze"; name: string; bytes: ArrayBuffer }
  | { id: number; type: "build"; choices: Choices; placeholderText: string };

export type AnalyzeResult = { doc: CmDocument; report: Report };
export type BuildResult = { name: string; bytes: Uint8Array };

export type WorkerResponse =
  | { id: number; ok: true; result: AnalyzeResult | BuildResult }
  | { id: number; ok: false; code: ErrorCode | "unknown"; problems: Problem[] };
```

`src/web/worker.ts`

```ts
import type { CmDocument } from "../core/model";
import { analyze, build, CleanMarkError } from "../core/pipeline";
import type { WorkerRequest, WorkerResponse } from "./protocol";

// 所有讀取、檢查、輸出都喺呢個背景執行緒做，畫面唔會卡

let current: { name: string; doc: CmDocument } | undefined;

function reply(msg: WorkerResponse, transfer: Transferable[] = []): void {
  (self as unknown as Worker).postMessage(msg, transfer);
}

self.onmessage = async (ev: MessageEvent<WorkerRequest>) => {
  const req = ev.data;
  try {
    if (req.type === "analyze") {
      const result = await analyze(req.name, new Uint8Array(req.bytes));
      current = { name: req.name, doc: result.doc };
      reply({ id: req.id, ok: true, result });
    } else {
      if (!current) throw new CleanMarkError("corrupt");
      const result = await build(current.name, current.doc, req.choices, { placeholderText: req.placeholderText });
      reply({ id: req.id, ok: true, result }, [result.bytes.buffer as ArrayBuffer]);
    }
  } catch (e) {
    if (e instanceof CleanMarkError) reply({ id: req.id, ok: false, code: e.code, problems: e.problems });
    else {
      console.error(e);
      reply({ id: req.id, ok: false, code: "unknown", problems: [] });
    }
  }
};
```

`src/web/worker-client.ts`

```ts
import type { Choices } from "../core/model";
import type { AnalyzeResult, BuildResult, WorkerRequest, WorkerResponse } from "./protocol";

type DistributiveOmit<T, K extends keyof T> = T extends unknown ? Omit<T, K> : never;

export class WorkerError extends Error {
  constructor(
    public code: string,
    public problems: { code: string; detail: string }[],
  ) {
    super(code);
  }
}

type Pending = { resolve: (v: unknown) => void; reject: (e: unknown) => void };

export class CleanMarkWorker {
  private worker = new Worker(new URL("./worker.ts", import.meta.url), { type: "module" });
  private nextId = 1;
  private pending = new Map<number, Pending>();

  constructor() {
    this.worker.onmessage = (ev: MessageEvent<WorkerResponse>) => {
      const msg = ev.data;
      const p = this.pending.get(msg.id);
      if (!p) return;
      this.pending.delete(msg.id);
      if (msg.ok) p.resolve(msg.result);
      else p.reject(new WorkerError(msg.code, msg.problems));
    };
  }

  private send<T>(req: DistributiveOmit<WorkerRequest, "id">, transfer: Transferable[] = []): Promise<T> {
    const id = this.nextId++;
    return new Promise<T>((resolve, reject) => {
      this.pending.set(id, { resolve: resolve as (v: unknown) => void, reject });
      this.worker.postMessage({ ...req, id }, transfer);
    });
  }

  analyze(name: string, bytes: ArrayBuffer): Promise<AnalyzeResult> {
    return this.send({ type: "analyze", name, bytes }, [bytes]);
  }

  build(choices: Choices, placeholderText: string): Promise<BuildResult> {
    return this.send({ type: "build", choices, placeholderText });
  }
}
```

- [ ] **Step 4：建立頁首、拖放區、檔案資料行、區塊外框**

`src/web/components/SiteHeader.svelte`

```svelte
<script lang="ts">
  import { i18n, setLang, t } from "../i18n/index.svelte";
</script>

<header>
  <div class="brand">
    <svg viewBox="0 0 32 32" width="32" height="32" aria-hidden="true">
      <rect width="32" height="32" rx="8" fill="var(--accent)" />
      <path d="M9 17l4.5 4.5L23 11" fill="none" stroke="var(--accent-text)" stroke-width="3" stroke-linecap="round" stroke-linejoin="round" />
    </svg>
    <div>
      <h1>CleanMark</h1>
      <p class="muted">{t("tagline")}</p>
    </div>
  </div>
  <button class="lang" onclick={() => setLang(i18n.lang === "en" ? "zh-Hant" : "en")}>{t("langSwitch")}</button>
</header>

<style>
  header {
    display: flex;
    justify-content: space-between;
    align-items: flex-start;
    gap: 12px;
    margin-bottom: 28px;
  }
  .brand {
    display: flex;
    gap: 12px;
    align-items: center;
  }
  h1 {
    margin: 0;
    font-size: 22px;
    line-height: 1.2;
  }
  p {
    margin: 2px 0 0;
    font-size: 14px;
  }
  .lang {
    white-space: nowrap;
    flex-shrink: 0;
    font-size: 14px;
    padding: 6px 12px;
  }
</style>
```

`src/web/components/DropZone.svelte`

```svelte
<script lang="ts">
  import { t } from "../i18n/index.svelte";

  let { onfile }: { onfile: (file: File) => void } = $props();
  let input: HTMLInputElement;
  let dragging = $state(false);

  function pick(files: FileList | null | undefined) {
    const f = files?.[0];
    if (f) onfile(f);
  }
</script>

<div
  class="zone"
  class:dragging
  role="button"
  tabindex="0"
  aria-label={t("dropTitle")}
  onclick={() => input.click()}
  onkeydown={(e) => (e.key === "Enter" || e.key === " ") && (e.preventDefault(), input.click())}
  ondragover={(e) => {
    e.preventDefault();
    dragging = true;
  }}
  ondragleave={() => (dragging = false)}
  ondrop={(e) => {
    e.preventDefault();
    dragging = false;
    pick(e.dataTransfer?.files);
  }}
>
  <svg viewBox="0 0 24 24" width="36" height="36" aria-hidden="true">
    <path d="M12 16V4m0 0l-4 4m4-4l4 4M4 16v2a2 2 0 002 2h12a2 2 0 002-2v-2" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" />
  </svg>
  <p class="title">{t("dropTitle")}</p>
  <span class="choose">{t("dropButton")}</span>
  <p class="muted hint">{t("dropHint")}</p>
  <input
    bind:this={input}
    type="file"
    accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    hidden
    data-testid="file-input"
    onchange={(e) => pick(e.currentTarget.files)}
  />
</div>
<p class="privacy">
  <svg viewBox="0 0 24 24" width="16" height="16" aria-hidden="true">
    <path d="M12 3l7 3v5c0 4.5-3 8.5-7 10-4-1.5-7-5.5-7-10V6l7-3z" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linejoin="round" />
  </svg>
  {t("privacy")}
</p>

<ol class="steps">
  <li><strong>1 · {t("step1")}</strong><span class="muted">{t("step1Desc")}</span></li>
  <li><strong>2 · {t("step2")}</strong><span class="muted">{t("step2Desc")}</span></li>
  <li><strong>3 · {t("step3")}</strong><span class="muted">{t("step3Desc")}</span></li>
</ol>

<style>
  .zone {
    border: 2px dashed var(--border-strong);
    border-radius: var(--radius-lg);
    background: var(--surface);
    padding: 48px 20px;
    text-align: center;
    cursor: pointer;
    transition: border-color 0.15s, background 0.15s;
  }
  .zone:hover,
  .zone.dragging {
    border-color: var(--accent);
    background: var(--accent-soft);
  }
  .title {
    font-size: 18px;
    font-weight: 600;
    margin: 8px 0 12px;
  }
  .choose {
    display: inline-block;
    background: var(--accent);
    color: var(--accent-text);
    border-radius: var(--radius);
    padding: 8px 18px;
    font-weight: 600;
  }
  .hint {
    font-size: 13px;
    margin: 12px 0 0;
  }
  .privacy {
    display: flex;
    justify-content: center;
    align-items: center;
    gap: 6px;
    font-size: 14px;
    color: var(--accent-soft-text);
    margin: 14px 0 32px;
  }
  .steps {
    list-style: none;
    padding: 0;
    margin: 0;
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(160px, 1fr));
    gap: 12px;
  }
  .steps li {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    padding: 12px 14px;
    display: flex;
    flex-direction: column;
    font-size: 14px;
  }
</style>
```

`src/web/components/FileBar.svelte`

```svelte
<script lang="ts">
  import { t } from "../i18n/index.svelte";

  let { name, count, onchange }: { name: string; count: number; onchange: () => void } = $props();
</script>

<div class="bar card">
  <svg viewBox="0 0 24 24" width="22" height="22" aria-hidden="true">
    <path d="M14 3H7a2 2 0 00-2 2v14a2 2 0 002 2h10a2 2 0 002-2V8l-5-5zm0 0v5h5M9 13h6M9 17h6" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linejoin="round" />
  </svg>
  <span class="name">{name}</span>
  {#if count > 0}
    <span class="pill warn" data-testid="found-count">{t("foundN", { n: count })}</span>
  {:else}
    <span class="pill ok" data-testid="found-count">{t("foundNone")}</span>
  {/if}
  <button onclick={onchange}>{t("changeFile")}</button>
</div>

<style>
  .bar {
    display: flex;
    align-items: center;
    gap: 10px;
    padding: 12px 16px;
    flex-wrap: wrap;
  }
  .name {
    flex: 1;
    min-width: 0;
    overflow-wrap: anywhere;
    font-weight: 600;
  }
  button {
    font-size: 14px;
    padding: 4px 12px;
  }
</style>
```

`src/web/components/Section.svelte`

```svelte
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
```

- [ ] **Step 5：建立檢查報告同確認面板**

`src/web/components/ReportPanel.svelte`

```svelte
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
```

`src/web/components/ConfirmPanel.svelte`

```svelte
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
          <span class="muted">· {t("pageN", { n: img.page })}</span>
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
    <p class="warning">{t("keepWarning")}</p>
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
  .small {
    font-size: 13px;
  }
  .warning {
    background: var(--warn-soft);
    color: var(--warn-text);
    border-radius: var(--radius);
    padding: 8px 12px;
    font-size: 14px;
    margin: 0;
  }
</style>
```

- [ ] **Step 6：建立預覽**

`src/web/components/PreviewBlocks.svelte`

```svelte
<script lang="ts">
  import type { Block } from "../../core/model";
  import { t } from "../i18n/index.svelte";
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
```

`src/web/components/Preview.svelte`

```svelte
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
```

- [ ] **Step 7：建立下載、錯誤、頁底**

`src/web/components/DownloadPanel.svelte`

```svelte
<script lang="ts">
  import { t } from "../i18n/index.svelte";

  let {
    phase,
    onbuild,
    ondownload,
    onreset,
  }: {
    phase: "ready" | "building" | "done";
    onbuild: () => void;
    ondownload: () => void;
    onreset: () => void;
  } = $props();
</script>

<div class="row">
  {#if phase === "ready"}
    <button class="primary" onclick={onbuild}>{t("buildButton")}</button>
  {:else if phase === "building"}
    <span class="muted" role="status">{t("building")}</span>
  {:else}
    <span class="pill ok" role="status" data-testid="self-check-ok">✓ {t("selfCheckPassed")}</span>
    <span class="spacer"></span>
    <button class="primary" onclick={ondownload}>{t("downloadAgain")}</button>
    <button onclick={onreset}>{t("another")}</button>
  {/if}
</div>

<style>
  .row {
    display: flex;
    gap: 10px;
    align-items: center;
    flex-wrap: wrap;
  }
  .spacer {
    flex: 1;
  }
</style>
```

`src/web/components/ErrorNotice.svelte`

```svelte
<script lang="ts">
  import { t, type MessageKey } from "../i18n/index.svelte";

  let { code, problems, onreset }: { code: string; problems: { code: string; detail: string }[]; onreset: () => void } =
    $props();
  const KNOWN = ["legacyFormat", "unsupported", "comingSoon", "tooLarge", "encrypted", "corrupt", "selfCheckFailed"];
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
```

`src/web/components/SiteFooter.svelte`

```svelte
<script lang="ts">
  import { t } from "../i18n/index.svelte";
</script>

<footer>
  <h2>{t("limitsTitle")}</h2>
  <ul>
    <li>{t("limitStatistical")}</li>
    <li>{t("limitDropped")}</li>
  </ul>
  <p class="rule">{t("usageRule")}</p>
</footer>

<style>
  footer {
    margin-top: 40px;
    padding-top: 20px;
    border-top: 1px solid var(--border);
    color: var(--text-muted);
    font-size: 14px;
  }
  h2 {
    font-size: 15px;
    color: var(--text);
    margin: 0 0 6px;
  }
  ul {
    margin: 0 0 12px;
    padding-left: 20px;
  }
  .rule {
    margin: 0;
  }
</style>
```

- [ ] **Step 8：建立 `src/web/App.svelte`**

`src/web/App.svelte`

```svelte
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
  let built = $state.raw<BuildResult | null>(null);
  let error = $state<{ code: string; problems: { code: string; detail: string }[] } | null>(null);

  const report = $derived<Report | null>(analysis?.report ?? null);
  const previewDoc = $derived(analysis ? applyChoices(analysis.doc, { removeParagraphs: removeIds, keepImages }) : null);

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
      removeIds = result.report.phrases.map((p) => p.paragraphId); // 預設剔咗刪除
      keepImages = [];
      phase = "ready";
    } catch (e) {
      fail(e);
    }
  }

  // 用家改咗選擇，之前整好嘅檔案就作廢。只追蹤選擇，唔追蹤 phase
  $effect(() => {
    $state.snapshot(removeIds);
    $state.snapshot(keepImages);
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
      const choices = $state.snapshot({ removeParagraphs: removeIds, keepImages });
      built = await worker.build(choices, t("placeholderText"));
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
      <ConfirmPanel {report} bind:removeIds bind:keepImages />
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
```

- [ ] **Step 9：型別檢查**

Run: `npm run check`
Expected: `0 ERRORS 0 WARNINGS`

- [ ] **Step 10：打包，確認有 CSP**

Run: `npm run build && grep -o "connect-src 'self'" dist/index.html`
Expected: 打包成功，並輸出 `connect-src 'self'`

- [ ] **Step 11：人手試用**

Run: `npm run preview`，打開 `http://localhost:4173/`

逐項確認：
1. 拖放區、私隱提示、三個步驟、頁底守則都有顯示
2. 揀一個有水印嘅 .docx：檔案資料行顯示「找到 N 項」，報告逐項撳得開
3. 可疑字句預設剔咗，預覽入面嗰段消失；取消剔選就出返
4. 圖片轉「保留原圖」會出現警告，預覽顯示原圖
5. 撳「清除並下載」：出現「新檔案自我檢查通過」，下載 `<原名>-clean.docx`，用 Word 打開確認冇水印、格式正常
6. 撳右上角「English」：成頁轉英文；重新整理之後仍然係英文
7. 瀏覽器調到 375px 闊：冇橫向捲動，語言掣唔會斷行

- [ ] **Step 12：Commit（由用家做）**

```bash
git add index.html public src/web
git commit -m "feat(web): single-page UI with worker-based processing"
```

---

### Task 13：瀏覽器端到端測試

**Files:**
- Create: `playwright.config.ts`, `e2e/global-setup.ts`, `e2e/clean-docx.spec.ts`

**Interfaces:**
- Consumes: Task 12 嘅 `data-testid`、Task 5 測試工具、Task 9 `selfCheckDocx`

- [ ] **Step 1：安裝瀏覽器**

Run: `npx playwright install chromium`
Expected: 下載 Chromium（大約 150 MB，第一次先需要）

- [ ] **Step 2：建立 `playwright.config.ts`**

`playwright.config.ts`

```ts
import { defineConfig, devices } from "@playwright/test";

export default defineConfig({
  testDir: "e2e",
  globalSetup: "./e2e/global-setup.ts",
  use: { baseURL: "http://localhost:4173", acceptDownloads: true },
  projects: [{ name: "chromium", use: { ...devices["Desktop Chrome"] } }],
  webServer: {
    command: "npm run build && npm run preview",
    url: "http://localhost:4173",
    reuseExistingServer: !process.env.CI,
    timeout: 120_000,
  },
});
```

- [ ] **Step 3：建立 `e2e/global-setup.ts`**（每次測試前產生樣本檔）

`e2e/global-setup.ts`

```ts
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
    body: p("第一章") + p("營業額\u200B上升 12%") + imageP("rIdA") + p("本文由 AI 協助生成。"),
  });
  writeFileSync("e2e/.fixtures/sample.docx", bytes);
}
```

- [ ] **Step 4：寫測試 `e2e/clean-docx.spec.ts`**

`e2e/clean-docx.spec.ts`

```ts
import { readFileSync } from "node:fs";
import { expect, test } from "@playwright/test";
import { readDocx } from "../src/core/readers/docx";
import { selfCheckDocx } from "../src/core/verify/self-check-docx";
import { allParagraphs, paragraphText } from "../src/core/walk";

test.use({ locale: "zh-TW" });

test("拖入 Word → 報告 → 改選擇 → 下載乾淨檔案", async ({ page }) => {
  await page.goto("/");
  await expect(page.getByText("把 Word 檔案拖到這裡")).toBeVisible();

  await page.getByTestId("file-input").setInputFiles("e2e/.fixtures/sample.docx");
  await expect(page.getByTestId("found-count")).toHaveText("找到 6 項");
  await expect(page.getByTestId("preview")).not.toContainText("本文由 AI");
  await expect(page.getByTestId("preview")).toContainText("圖片 1：請重新插入");

  await page.getByText("保留原圖").click();
  await expect(page.getByText("原圖可能帶有水印")).toBeVisible();
  await page.getByText("佔位框", { exact: true }).click();

  const downloadPromise = page.waitForEvent("download");
  await page.getByRole("button", { name: "清除並下載" }).click();
  const download = await downloadPromise;
  expect(download.suggestedFilename()).toBe("sample-clean.docx");
  await expect(page.getByTestId("self-check-ok")).toBeVisible();

  const bytes = new Uint8Array(readFileSync(await download.path()));
  expect(await selfCheckDocx(bytes, 0)).toEqual({ ok: true, problems: [] });
  const { doc } = await readDocx(bytes);
  expect(allParagraphs(doc).map(paragraphText)).toEqual(["第一章", "營業額上升 12%", "圖片 1：請重新插入", "季度報告"]);
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
```

- [ ] **Step 5：行測試**

Run: `npm run e2e`
Expected: `4 passed`

- [ ] **Step 6：Commit（由用家做）**

```bash
git add playwright.config.ts e2e/global-setup.ts e2e/clean-docx.spec.ts
git commit -m "test(e2e): cover upload, report, download and self-check"
```

---

### Task 14：README、授權、CI 同 GitHub Pages

**Files:**
- Create: `README.md`, `LICENSE`, `.github/workflows/ci.yml`

- [ ] **Step 1：建立 `README.md`**

`README.md`

````markdown
# CleanMark（乾淨文件）

移除 Word 文件嘅水印同隱藏資料。**檔案只喺你部電腦嘅瀏覽器入面處理，唔會上傳去任何地方。**

Remove watermarks and hidden data from Word documents, entirely in your browser. Nothing is uploaded.

## 點運作

CleanMark 唔係「刪除水印」，而係**重新打一份**：

1. 讀出原檔嘅文字同格式（標題、粗體、斜體、字體大細、顏色、列表、表格）
2. 喺一份空白新檔案入面重新砌出嚟
3. 水印、頁首 logo、作者同軟件資料、C2PA 簽章、隱形字元一律唔搬
4. 新檔案整好之後再自我檢查一次，通過先俾你下載

## 檢查同清除

| 類別 | 例子 |
|---|---|
| 可見水印 | Word 內建水印、頁首頁尾嘅圖形同 logo、文字後面嘅圖形 |
| 檔案資料 | 作者、最後修改人、建立軟件、自訂欄位、customXml |
| Content Credentials | C2PA 簽章 |
| 隱形字元 | 零寬度空格、Tag 字元、方向覆寫等（阿拉伯文、波斯文、emoji 需要嘅會保留） |
| 可疑字句 | 例如「由 AI 生成」、「Generated by」，由你決定刪唔刪 |

## 做唔到嘅嘢

- **統計水印**（藏喺用字規律入面，例如 SynthID-Text）：要 AI 公司嘅私人密鑰先驗證到，CleanMark 唔會檢查亦唔會清除
- Word 註解唔會保留；圖表、SmartArt、方程式會變成佔位框
- 暫時只支援 .docx，PowerPoint 同 PDF 即將推出

## 私隱

- 網頁設定咗 Content Security Policy（`connect-src 'self'`），瀏覽器層面禁止連去任何外部網址
- 冇任何分析或者追蹤工具

## 使用守則

只可以用喺你有權處理嘅文件。唔好用嚟移除他人嘅版權、機密或者其他權利標記。

## 開發

```bash
npm install
npm run dev      # 本機開發
npm test         # 單元測試
npm run e2e      # 瀏覽器測試（第一次要先行 npx playwright install chromium）
npm run build    # 打包到 dist/
```

## 授權

MIT
````

- [ ] **Step 2：建立 `LICENSE`**（MIT；`CleanMark contributors` 可以換成你嘅名）

`LICENSE`

```text
MIT License

Copyright (c) 2026 CleanMark contributors

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

- [ ] **Step 3：建立 `.github/workflows/ci.yml`**

`.github/workflows/ci.yml`

```yaml
name: CI

on:
  push:
    branches: [main]
  pull_request:

permissions:
  contents: read

jobs:
  test:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-node@v4
        with:
          node-version: 24
          cache: npm
      - run: npm ci
      - run: npm run check
      - run: npm test
      - run: npx playwright install --with-deps chromium
      - run: npm run e2e

  deploy:
    # 只有 main 分支、而且測試全部通過先發佈
    if: github.event_name == 'push' && github.ref == 'refs/heads/main'
    needs: test
    runs-on: ubuntu-latest
    permissions:
      pages: write
      id-token: write
    environment:
      name: github-pages
      url: ${{ steps.deployment.outputs.page_url }}
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-node@v4
        with:
          node-version: 24
          cache: npm
      - run: npm ci
      - run: npm run build
      - uses: actions/configure-pages@v5
      - uses: actions/upload-pages-artifact@v3
        with:
          path: dist
      - id: deployment
        uses: actions/deploy-pages@v4
```

- [ ] **Step 4：本機行齊所有檢查**

Run: `npm run check && npm test && npm run e2e`
Expected: `0 ERRORS`、`Tests  60 passed (60)`、`4 passed`

- [ ] **Step 5：Commit（由用家做）**

```bash
git add README.md LICENSE .github/workflows/ci.yml
git commit -m "docs: add README, MIT license and CI with GitHub Pages deploy"
```

- [ ] **Step 6：上 GitHub（由用家做，或者叫 Claude 幫手）**

```bash
gh repo create cleanmark --public --source . --push
```

之後喺 GitHub repo 嘅 **Settings → Pages → Build and deployment → Source** 揀 **GitHub Actions**。下一次 push 去 `main`，CI 通過之後就會發佈到 `https://<你嘅帳戶>.github.io/cleanmark/`。

---

## 完成標準

- `npm run check`：0 錯誤 0 警告
- `npm test`：60 個單元測試全部通過
- `npm run e2e`：4 個瀏覽器測試全部通過
- 人手試用（Task 12 Step 11）七項全部正常
- GitHub Pages 網址打得開，用一份真實、有水印嘅 Word 檔試一次，下載到嘅檔案用 Word 打開冇水印、格式正確

---

## 追加：試用後修訂（2026-10-15）

用一份真實嘅 50 頁學術報告試用之後，做咗以下修改。以下內容以 repo 入面嘅 code 為準，上面各 task 嘅 code 係修改前嘅版本。

| 修改 | 原因 | 涉及檔案 |
|---|---|---|
| 方程式按白名單原樣保留，唔再變佔位框 | 學術報告有大量方程式；方程式係純數學文字，藏唔到圖形水印 | 新增 `src/core/math.ts`、`src/web/math-text.ts`、`tests/math.test.ts`；改 `model.ts`、`readers/docx/body.ts`、`clean/apply-choices.ts`、`writers/docx.ts`、`PreviewBlocks.svelte` |
| 方程式入面嘅排版空格（U+2000–U+200A、U+205F）保留 | 原本會被換成普通空格，令方程式間距走樣 | `clean/hidden-chars.ts`（新增 `math` 選項） |
| 方程式逐個元素用 `ImportedXmlComponent` 砌，唔用 `fromXmlString` | `fromXmlString` 會丟走只有空格嘅 `m:t` | `writers/docx.ts` |
| 可疑字句預設保留 | AI 使用聲明等段落可能係規定要保留 | `App.svelte`、`e2e/clean-docx.spec.ts` |
| 圖片顯示位置提示（最近嘅標題），唔再顯示頁數 | `w:lastRenderedPageBreak` 唔可靠，50 頁嘅文件全部圖片都顯示第 7–10 頁 | `model.ts`（`ImageSlot.context`、`ImageItem.context`）、`readers/docx/body.ts`、`ConfirmPanel.svelte`、翻譯檔 |
| 原始隱形字元改成 `\uXXXX` 轉義寫法 | 測試同 code 入面嘅隱形字元肉眼睇唔到 | 測試檔、`hidden-chars.ts`、本計劃 |

修訂後：單元測試 68 個、瀏覽器測試 4 個全部通過；試用文件嘅 1357 段內文同 18 條方程式寫出之後同原檔一字不差。

---

## 第 1b 期：白名單重新抄寫（2026-10-15）

**問題：** 用 Word 將一份 50 頁學術報告嘅原檔同新檔轉 PDF 逐頁比較，依家嘅做法（內容清單 → docx 套件）由第 1 頁開始已經唔同，新檔仲長咗好多頁。目錄點線、定位點、樣式、主題字型、相容性設定都喺翻譯途中冇咗。

**決定（用家已同意）：** 繼續「重新打一次」，但改成**白名單重新抄寫**：逐個 Word 元素讀，只有白名單入面、數值驗證過嘅元素同屬性先會重新寫入全新檔案。範圍只做 Word（PDF、PowerPoint 唔做）。

**完成標準：**
1. `scripts/verify-layout.mts` 用 Word 將測試報告原檔同新檔轉 PDF，頁數一樣、每頁畫面一樣（`IDENTICAL`）
2. 新檔通過自我檢查；搜尋唔到 `rsid`、`paraId`、`docId`、`durableId`、原作者名、customXml
3. 全部單元測試、瀏覽器測試通過

**Tasks：**
1. `src/core/sanitize/policy.ts`：w、m、wp、a、pic、wps、wpg、v 嘅元素同屬性白名單；要拆走外殼嘅元素（`ins`、`sdt`、`smartTag`、`customXml`）；一律刪除嘅元素（`del`、註解、`rsid`、`docVars` 等）
2. `src/core/sanitize/prepass.ts`：處理前先改 DOM：隱形字元（按段落上下文）、只准白名單欄位類型、刪除用家揀咗嘅段落
3. `src/core/sanitize/serialize.ts`：按白名單重新寫 XML，重新編 relationship ID
4. `src/core/sanitize/images.ts`：清走 PNG / JPEG / GIF 入面嘅 EXIF、XMP、C2PA、文字資料；產生佔位圖
5. `src/core/sanitize/package.ts`：砌全新 zip：Content Types、rels、樣式、編號、設定（白名單）、字型表、主題、頁首頁尾、註腳尾註、圖片、中性 docProps
6. 讀取器嘅段落 ID 改成「part 名 + 段落序號」，令報告、預覽同重新抄寫對得上
7. 圖片預設「保留（已清除圖片資料）」，佔位框變成選項
8. 移除 `docx` 套件同舊輸出器；更新測試、自我檢查、介面、設計文件
9. 用 Word 逐頁比對測試報告，直到 `IDENTICAL`

**結果（2026-10-15）：** 全部完成。測試報告用 Word 轉 PDF 逐頁比較：104 頁對 104 頁，`IDENTICAL`。原檔入面嘅 `rsid`、`paraId`、`docId`、`durableId`、作者名、`Normal.dotm` 等，新檔全部搵唔到。單元測試 80 個、瀏覽器測試 4 個全部通過。過程中發現兩個問題並已修正：主題顏色用冇前綴嘅 `val` 屬性（漏咗白名單令 Word 打唔開）；全形空格被替換令圖表目錄移位（特殊空格規則收窄至只替換 U+202F）。
