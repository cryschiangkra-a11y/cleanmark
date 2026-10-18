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
      { index: 1, context: "", keepable: true, tags: ["exif", "c2pa"] },
      { index: 2, context: "", keepable: true, tags: [] },
    ]);
  });

  test("可疑字句", async () => {
    const r = await inspect({ body: p("正常段落") + p("本文由 AI 協助生成。") + p("Drafted with ChatGPT") });
    expect(r.phrases.map((x) => [x.paragraphId, x.match])).toEqual([
      ["document:2", "由 AI 協助生成"],
      ["document:3", "ChatGPT"],
    ]);
  });

  test("隱形字元統計", async () => {
    const r = await inspect({
      body: p("A\u200BB\u200BC") + p("👨\u200D👩") + `<w:p><m:oMath><m:r><m:t>x\u200B</m:t></m:r></m:oMath></w:p>`,
    });
    expect(r.hidden).toEqual({ removed: { "U+200B": 3 }, replaced: {}, kept: 1 });
  });
});

test("inspectDocx：可疑字元", async () => {
  const r = await inspect({
    body: p("Clаude said") + p("don\u02BCt") + p("a\u00A0b\u2009c") + `<w:p><m:oMath><m:r><m:t>Ο</m:t></m:r></m:oMath></w:p>`,
  });
  expect(r.suspicious).toEqual({
    homoglyphs: [{ paragraphId: "document:1", word: "Clаude", fixed: "Claude" }],
    apostrophes: 1,
    specialSpaces: { "U+00A0": 1, "U+2009": 1 },
  });
});
