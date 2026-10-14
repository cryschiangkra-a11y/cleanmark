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
  settings?: string; // <w:settings> 入面嘅內容
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
  if (f.settings !== undefined) {
    docRels.push(rel("rIdSettings", "settings", "settings.xml"));
    zip.file("word/settings.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ${NS_DECL}>${f.settings}</w:settings>`);
  }
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
