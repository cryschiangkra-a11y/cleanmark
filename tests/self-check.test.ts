import JSZip from "jszip";
import { expect, test } from "vitest";
import type { Choices } from "../src/core/model";
import { OoxmlPackage } from "../src/core/ooxml/package";
import { writeCleanDocx } from "../src/core/sanitize/package";
import { selfCheckDocx } from "../src/core/verify/self-check-docx";
import { buildDocx, imageP, p, PNG_1PX, pngWith, wordTextWatermark } from "./fixtures/docx-builder";

const clean = async (bytes: Uint8Array, choices: Choices) => writeCleanDocx(await OoxmlPackage.open(bytes), choices);

/** 「乜都有」嘅原檔：水印、logo、metadata、C2PA、隱形字元 */
const dirty = () =>
  buildDocx({
    body: p("內文​ ") + imageP("rIdA"),
    media: [{ rId: "rIdA", name: "a.png", bytes: pngWith("c2pa") }],
    header: p("頁首") + wordTextWatermark("AI Generated") + imageP("rIdLogo"),
    headerMedia: [{ rId: "rIdLogo", name: "logo.png", bytes: PNG_1PX }],
    core: { "dc:creator": "ChatGPT" },
    app: { Application: "AI Writer" },
    custom: { Source: "Gemini" },
    customXml: "<x>ai</x>",
  });

test("原檔處理之後，自我檢查通過", async () => {
  const out = await clean(await dirty(), { removeParagraphs: [], keepImages: [] });
  expect(await selfCheckDocx(out.bytes, out.mediaCount)).toEqual({ ok: true, problems: [] });
});

test("保留原圖都通過：圖片入面嘅 C2PA 已經清走", async () => {
  const out = await clean(await dirty(), { removeParagraphs: [], keepImages: [1] });
  expect(await selfCheckDocx(out.bytes, out.mediaCount)).toEqual({ ok: true, problems: [] });
});

test("原檔本身唔會通過，每類問題都捉到", async () => {
  const r = await selfCheckDocx(await dirty(), 0);
  expect(r.ok).toBe(false);
  expect(new Set(r.problems.map((x) => x.code))).toEqual(
    new Set(["unexpectedPart", "unexpectedMedia", "metadata", "visible", "hiddenChars", "c2pa", "imageMetadata"]),
  );
});

test("輸出器以外多咗一個檔案就唔通過", async () => {
  const out = await clean(await buildDocx({ body: p("x") }), { removeParagraphs: [], keepImages: [] });
  const zip = await JSZip.loadAsync(out.bytes);
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

test("揀咗換同形字，但新檔仍然有：唔通過", async () => {
  const bytes = await buildDocx({ body: p("Clаude") });
  const notFixed = await clean(bytes, { removeParagraphs: [], keepImages: [], fixHomoglyphs: false });
  const fix = { homoglyphs: true, apostrophes: true, specialSpaces: false };
  expect((await selfCheckDocx(notFixed.bytes, notFixed.mediaCount, fix)).problems).toEqual([
    { code: "suspiciousChars", detail: "homoglyphs×1" },
  ]);
  const fixed = await clean(bytes, { removeParagraphs: [], keepImages: [] });
  expect((await selfCheckDocx(fixed.bytes, fixed.mediaCount, fix)).ok).toBe(true);
});
