import JSZip from "jszip";
import { describe, expect, test } from "vitest";
import { applyChoices } from "../src/core/clean/apply-choices";
import { imageTags } from "../src/core/inspect/bytes";
import type { Block, Choices, Paragraph } from "../src/core/model";
import { OoxmlPackage } from "../src/core/ooxml/package";
import { readDocx } from "../src/core/readers/docx";
import { writeCleanDocx } from "../src/core/sanitize/package";
import { allImages, allParagraphs, paragraphText } from "../src/core/walk";
import { buildDocx, imageP, p, PNG_1PX, pngWith, wordTextWatermark } from "./fixtures/docx-builder";

const NONE: Choices = { removeParagraphs: [], keepImages: [] };

async function clean(bytes: Uint8Array, choices: Choices = NONE): Promise<Uint8Array> {
  return (await writeCleanDocx(await OoxmlPackage.open(bytes), choices)).bytes;
}

async function partText(bytes: Uint8Array, path: string): Promise<string> {
  return (await (await JSZip.loadAsync(bytes)).file(path)?.async("string")) ?? "";
}

/** 去掉 id、圖片路徑，淨係比較內容同格式 */
function normalise(blocks: Block[]): unknown {
  return blocks.map((b) => {
    if (b.type === "paragraph") {
      const { id: _id, ...rest } = b;
      return rest;
    }
    if (b.type === "table") return { ...b, rows: b.rows.map((r) => r.map((c) => ({ ...c, blocks: normalise(c.blocks) }))) };
    if (b.type === "image") {
      const { mediaPath: _m, keep: _k, ...rest } = b;
      return rest;
    }
    return b;
  });
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
    p("中文 English 日本語 한국어 العربية") +
    `<w:p><w:r><w:t xml:space="preserve">公式 </w:t></w:r><m:oMath><m:sSup><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup><m:r><m:t xml:space="preserve"> </m:t></m:r><m:r><m:t>&lt;y</m:t></m:r></m:oMath></w:p>` +
    `<w:p><m:oMathPara><m:oMath><m:f><m:num><m:r><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr><m:t>a</m:t></m:r></m:num><m:den><m:r><m:t>b</m:t></m:r></m:den></m:f></m:oMath></m:oMathPara></w:p>`,
};

describe("writeCleanDocx：讀 → 寫 → 再讀，格式唔走樣", () => {
  test("段落、文字格式、列表、表格、頁首頁尾、註腳、方程式", async () => {
    const bytes = await buildDocx(RICH);
    const { doc } = await readDocx(bytes);
    const expected = applyChoices(doc, NONE);
    const again = (await readDocx(await clean(bytes))).doc;
    expect(normalise(again.body)).toEqual(normalise(expected.body));
    expect(normalise(again.header)).toEqual(normalise(expected.header));
    expect(normalise(again.footer)).toEqual(normalise(expected.footer));
    expect(again.page).toEqual(expected.page);
    expect(again.lists).toEqual(expected.lists);
    expect(Object.values(again.footnotes).map((b) => b.filter((x): x is Paragraph => x.type === "paragraph").map(paragraphText))).toEqual([["註腳"]]);
  });
});

describe("writeCleanDocx：唔搬嘅嘢", () => {
  const W14 = 'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';
  const dirty = () =>
    buildDocx({
      core: { "dc:creator": "ChatGPT", "cp:lastModifiedBy": "AI" },
      app: { Application: "AI Writer" },
      custom: { Source: "Gemini" },
      customXml: "<x>ai</x>",
      settings: `<w:zoom w:percent="120"/><w:defaultTabStop w:val="720"/><w:rsids><w:rsidRoot w:val="00AB12CD"/></w:rsids><w:docVars><w:docVar w:name="aiTool" w:val="secret"/></w:docVars><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>`,
      header: p("頁首") + wordTextWatermark("AI Generated") + imageP("rIdLogo"),
      headerMedia: [{ rId: "rIdLogo", name: "logo.png", bytes: PNG_1PX }],
      body:
        `<w:p ${W14} w14:paraId="1A2B3C4D" w:rsidR="00AB12CD"><w:r w:rsidRPr="00AB12CD"><w:t>內文</w:t></w:r></w:p>` +
        `<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>有註解</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>` +
        `<w:p><w:ins w:id="1" w:author="Bot" w:date="2026-01-01T00:00:00Z"><w:r><w:t>新增</w:t></w:r></w:ins><w:del w:id="2" w:author="Bot"><w:r><w:delText>刪除</w:delText></w:r></w:del></w:p>`,
    });

  test("metadata、編輯紀錄碼、段落 ID、customXml、註解、修訂作者都冇咗", async () => {
    const out = await clean(await dirty());
    const zip = await JSZip.loadAsync(out);
    const all = (await Promise.all(Object.values(zip.files).filter((f) => !f.dir).map((f) => f.async("string")))).join("\n");
    for (const s of ["ChatGPT", "AI Writer", "Gemini", "customXml", "rsid", "paraId", "aiTool", "secret", "Bot", "comment", "AI Generated", "zoom"]) {
      expect(all, s).not.toContain(s);
    }
    expect(Object.values(zip.files).filter((f) => !f.dir).map((f) => f.name).sort()).toEqual([
      "[Content_Types].xml", "_rels/.rels", "docProps/app.xml", "docProps/core.xml",
      "word/_rels/document.xml.rels", "word/document.xml", "word/header1.xml", "word/settings.xml", "word/styles.xml",
    ]);
  });

  test("設定只保留影響排版嘅項目", async () => {
    const settings = await partText(await clean(await dirty()), "word/settings.xml");
    expect(settings).toContain('<w:defaultTabStop w:val="720"/>');
    expect(settings).toContain('w:name="compatibilityMode"');
  });

  test("頁首文字保留，圖形同 WordArt 唔搬", async () => {
    const out = await clean(await dirty());
    const header = await partText(out, "word/header1.xml");
    expect(header).toContain("頁首");
    expect(header).not.toContain("w:pict");
    expect(header).not.toContain("w:drawing");
  });

  test("修訂：新增保留、刪除唔要", async () => {
    const doc = (await readDocx(await clean(await dirty()))).doc;
    expect(allParagraphs(doc).map(paragraphText)).toContain("新增");
    expect(allParagraphs(doc).map(paragraphText).join()).not.toContain("刪除");
  });
});

describe("writeCleanDocx：欄位、連結、書籤", () => {
  const field = (instr: string, result: string) =>
    `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve">${instr}</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>${result}</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`;

  test("准嘅欄位（目錄、頁碼）保留指令；其他欄位（例如 AUTHOR）清空指令、保留結果", async () => {
    const out = await clean(await buildDocx({ body: field(' TOC \\o "1-3" \\h ', "目錄") + field(" PAGE ", "3") + field(" AUTHOR ", "陳大文") }));
    const xml = await partText(out, "word/document.xml");
    expect(xml).toContain('TOC \\o "1-3" \\h');
    expect(xml).toContain(" PAGE ");
    expect(xml).not.toContain("AUTHOR");
    expect(xml).toContain("陳大文");
  });

  test("只准 http、https、mailto 外部連結", async () => {
    const out = await clean(
      await buildDocx({
        links: { rIdOk: "https://example.com/", rIdBad: "javascript:alert(1)" },
        body: `<w:p><w:hyperlink r:id="rIdOk"><w:r><w:t>好</w:t></w:r></w:hyperlink><w:hyperlink r:id="rIdBad"><w:r><w:t>壞</w:t></w:r></w:hyperlink></w:p>`,
      }),
    );
    const rels = await partText(out, "word/_rels/document.xml.rels");
    expect(rels).toContain("https://example.com/");
    expect(rels).not.toContain("javascript");
    expect(allParagraphs((await readDocx(out)).doc).map(paragraphText)).toEqual(["好壞"]);
  });

  test("書籤名稱唔合規格就唔搬", async () => {
    const out = await clean(
      await buildDocx({
        body: `<w:p><w:bookmarkStart w:id="1" w:name="_Toc123"/><w:bookmarkEnd w:id="1"/><w:bookmarkStart w:id="2" w:name="AI generated by tool"/><w:r><w:t>x</w:t></w:r></w:p>`,
      }),
    );
    const xml = await partText(out, "word/document.xml");
    expect(xml).toContain('w:name="_Toc123"');
    expect(xml).not.toContain("AI generated");
  });
});

describe("writeCleanDocx：用家選擇", () => {
  test("刪除段落；有章節設定嘅段落只清空內容", async () => {
    const sect = `<w:p><w:pPr><w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:pPr><w:r><w:t>章節尾</w:t></w:r></w:p>`;
    const out = await clean(await buildDocx({ body: p("留") + p("刪") + sect }), { removeParagraphs: ["document:2", "document:3"], keepImages: [] });
    const xml = await partText(out, "word/document.xml");
    expect(allParagraphs((await readDocx(out)).doc).map(paragraphText)).toEqual(["留", ""]);
    expect(xml).toContain("<w:sectPr>");
  });

  const withImages = () =>
    buildDocx({
      body: imageP("rIdA", 1905000, 952500) + imageP("rIdB", 952500, 952500),
      media: [
        { rId: "rIdA", name: "a.png", bytes: pngWith("Exif\0\0c2pa") },
        { rId: "rIdB", name: "b.png", bytes: PNG_1PX },
      ],
    });

  test("保留嘅圖片清走 EXIF、C2PA，尺寸唔變", async () => {
    const out = await clean(await withImages(), { removeParagraphs: [], keepImages: [1, 2] });
    const doc = (await readDocx(out)).doc;
    expect(allImages(doc).map((i) => [i.widthEmu, i.heightEmu])).toEqual([[1905000, 952500], [952500, 952500]]);
    for (const bytes of Object.values(doc.media)) expect(imageTags(bytes)).toEqual([]);
  });

  test("冇揀保留嘅圖片換成同尺寸佔位圖", async () => {
    const out = await clean(await withImages(), { removeParagraphs: [], keepImages: [2] });
    const doc = (await readDocx(out)).doc;
    const [first, second] = allImages(doc);
    expect([first.widthEmu, first.heightEmu]).toEqual([1905000, 952500]);
    expect(doc.media[first.mediaPath]).not.toEqual(PNG_1PX);
    expect(doc.media[second.mediaPath]).toEqual(PNG_1PX);
  });
});
