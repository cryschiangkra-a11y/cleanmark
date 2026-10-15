import JSZip from "jszip";
import { describe, expect, test } from "vitest";
import { OoxmlPackage, resolvePath, ZipBudgetError } from "../src/core/ooxml/package";
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

describe("OoxmlPackage：防壓縮炸彈", () => {
  const zipOf = async (files: Record<string, string | Uint8Array>) => {
    const zip = new JSZip();
    for (const [k, v] of Object.entries(files)) zip.file(k, v);
    return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
  };

  test("單一檔案報稱超過上限：打開時已經拒絕", async () => {
    const bomb = await zipOf({ "word/document.xml": new Uint8Array(2 * 1024 * 1024) }); // 2 MB 嘅零，壓縮後好細
    expect(bomb.length).toBeLessThan(50 * 1024);
    await expect(OoxmlPackage.open(bomb, { maxTotalBytes: 1024 * 1024, maxEntries: 100 })).rejects.toBeInstanceOf(ZipBudgetError);
  });

  test("解壓時逐段計數：加埋超過上限就停止", async () => {
    const pkg = await OoxmlPackage.open(
      await zipOf({ "a.xml": new Uint8Array(700 * 1024), "b.xml": new Uint8Array(700 * 1024) }),
      { maxTotalBytes: 1024 * 1024, maxEntries: 100 },
    );
    expect((await pkg.bytes("a.xml"))?.length).toBe(700 * 1024);
    await expect(pkg.bytes("b.xml")).rejects.toBeInstanceOf(ZipBudgetError);
  });

  test("內部檔案太多就拒絕", async () => {
    const many = await zipOf(Object.fromEntries(Array.from({ length: 20 }, (_, i) => [`f${i}.xml`, "<x/>"])));
    await expect(OoxmlPackage.open(many, { maxTotalBytes: 1024 * 1024, maxEntries: 10 })).rejects.toBeInstanceOf(ZipBudgetError);
  });

  test("同一個檔案讀幾次都只計一次", async () => {
    const pkg = await OoxmlPackage.open(await zipOf({ "a.xml": "x".repeat(600) }), { maxTotalBytes: 1000, maxEntries: 10 });
    for (let i = 0; i < 5; i++) expect(await pkg.text("a.xml")).toBe("x".repeat(600));
  });
});
