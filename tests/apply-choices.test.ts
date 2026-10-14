import { expect, test } from "vitest";
import { applyChoices } from "../src/core/clean/apply-choices";
import { mathPlainText } from "../src/core/math";
import type { ImageSlot, MathInline, Paragraph, TextInline } from "../src/core/model";
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
  const out = applyChoices(doc, { removeParagraphs: ["document:2"], keepImages: [] });
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
  applyChoices(doc, { removeParagraphs: ["document:1"], keepImages: [] });
  const first = doc.body[0] as Paragraph;
  expect((first.inlines[0] as TextInline).text).toBe("第一段\u200B");
});

test("方程式入面嘅隱形字元都清走", async () => {
  const { doc } = await readDocx(await buildDocx({ body: `<w:p><m:oMath><m:r><m:t>x\u200B+1</m:t></m:r></m:oMath></w:p>` }));
  const out = applyChoices(doc, { removeParagraphs: [], keepImages: [] });
  const m = (out.body[0] as Paragraph).inlines[0] as MathInline;
  expect(mathPlainText(m.root)).toBe("x+1");
});

test("方程式入面嘅細空格保留", async () => {
  const { doc } = await readDocx(await buildDocx({ body: `<w:p><m:oMath><m:r><m:t>M\u2009G</m:t></m:r></m:oMath></w:p>` }));
  const out = applyChoices(doc, { removeParagraphs: [], keepImages: [] });
  expect(mathPlainText(((out.body[0] as Paragraph).inlines[0] as MathInline).root)).toBe("M\u2009G");
});
