import { describe, expect, test } from "vitest";
import { mathToXml } from "../src/core/math";
import type { Block, ImageSlot, MathInline, Paragraph, Table, TextInline } from "../src/core/model";
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
  test("內嵌圖片切開段落，記錄尺寸同前面嘅文字", async () => {
    const bytes = await buildDocx({
      media: [{ rId: "rIdImg", name: "image1.png", bytes: PNG_1PX }],
      body: p("前") + p("") + imageP("rIdImg", 1905000, 952500),
    });
    const { doc } = await readDocx(bytes);
    const img = doc.body.find((b): b is ImageSlot => b.type === "image")!;
    expect(img).toEqual({
      type: "image", index: 1, widthEmu: 1905000, heightEmu: 952500, context: "前",
      mediaPath: "word/media/image1.png", contentType: "png", keep: false,
    });
    expect(doc.media["word/media/image1.png"]).toEqual(PNG_1PX);
    expect(doc.body.filter((b) => b.type === "paragraph")).toHaveLength(2); // 圖片段落唔會留低空段落
  });

  test("圖片位置優先用最近嘅標題", async () => {
    const bytes = await buildDocx({
      styles: STYLES,
      media: [{ rId: "rIdImg", name: "image1.png", bytes: PNG_1PX }],
      body: p("3.1 Methodology and Data Collection", { style: "Heading1" }) + p("內文") + imageP("rIdImg"),
    });
    const { doc } = await readDocx(bytes);
    expect((doc.body.find((b) => b.type === "image") as ImageSlot).context).toBe("3.1 Methodology and Data Collection");
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

  test("方程式原樣保留，只留 OMML 元素同基本字型格式", async () => {
    const math =
      `<m:oMath><m:f><m:num><m:r><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/><w:i/><w:color w:val="FF0000"/></w:rPr><m:t>a</m:t></m:r></m:num>` +
      `<m:den><m:r><m:t>b</m:t></m:r></m:den></m:f><w:bookmarkStart w:id="1" w:name="x"/><m:r><w:drawing/><m:t>c</m:t></m:r></m:oMath>`;
    const { doc } = await readDocx(
      await buildDocx({ body: `<w:p><w:r><w:t xml:space="preserve">x = </w:t></w:r>${math}</w:p><w:p><m:oMathPara>${math}</m:oMathPara></w:p>` }),
    );
    const [inlineP, displayP] = paras(doc.body);
    expect(inlineP.inlines.map((i) => i.type)).toEqual(["text", "math"]);
    const m = inlineP.inlines[1] as MathInline;
    expect(m.display).toBe(false);
    expect(mathToXml(m.root)).toBe(
      `<m:oMath><m:f><m:num><m:r><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/><w:i/><w:color w:val="FF0000"/></w:rPr><m:t xml:space="preserve">a</m:t></m:r></m:num>` +
        `<m:den><m:r><m:t xml:space="preserve">b</m:t></m:r></m:den></m:f><m:r><m:t xml:space="preserve">c</m:t></m:r></m:oMath>`,
    );
    expect((displayP.inlines[0] as MathInline).display).toBe(true);
    expect(doc.body.some((b) => b.type === "image")).toBe(false);
  });

  test("直接喺內文層嘅方程式段落都保留", async () => {
    const { doc } = await readDocx(await buildDocx({ body: `<m:oMathPara><m:oMath><m:r><m:t>x</m:t></m:r></m:oMath></m:oMathPara>` }));
    expect(paras(doc.body)[0].inlines).toMatchObject([{ type: "math", display: true }]);
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
    expect(paras(doc.header)[0].id).toBe("header1:1");
  });
});
