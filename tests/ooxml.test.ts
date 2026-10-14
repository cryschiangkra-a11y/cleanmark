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
