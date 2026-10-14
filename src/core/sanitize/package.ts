import JSZip from "jszip";
import type { Choices } from "../model";
import { partName } from "../ooxml/ids";
import type { OoxmlPackage, Relationship } from "../ooxml/package";
import type { Element } from "../ooxml/xml";
import { placeholderPng, stripImageMetadata } from "./images";
import { prepass } from "./prepass";
import { serializePart, type ImageRequest, type SerializeCtx } from "./serialize";

const MAIN = "word/document.xml";
const REL_BASE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const CT_BASE = "application/vnd.openxmlformats-officedocument.wordprocessingml";

/** 新檔案入面一定要有嘅 part 同佢哋嘅 relationship 類型 */
const CORE_PARTS: { type: string; name: string; contentType: string }[] = [
  { type: "styles", name: "styles.xml", contentType: `${CT_BASE}.styles+xml` },
  { type: "numbering", name: "numbering.xml", contentType: `${CT_BASE}.numbering+xml` },
  { type: "settings", name: "settings.xml", contentType: `${CT_BASE}.settings+xml` },
  { type: "fontTable", name: "fontTable.xml", contentType: `${CT_BASE}.fontTable+xml` },
  { type: "theme", name: "theme/theme1.xml", contentType: "application/vnd.openxmlformats-officedocument.theme+xml" },
  { type: "footnotes", name: "footnotes.xml", contentType: `${CT_BASE}.footnotes+xml` },
  { type: "endnotes", name: "endnotes.xml", contentType: `${CT_BASE}.endnotes+xml` },
];

const IMAGE_TYPES = { png: "image/png", jpg: "image/jpeg", gif: "image/gif", bmp: "image/bmp" } as const;
type ImageExt = keyof typeof IMAGE_TYPES;

function imageExt(path: string): ImageExt | null {
  const ext = path.slice(path.lastIndexOf(".") + 1).toLowerCase();
  if (ext === "jpeg" || ext === "jpg") return "jpg";
  return ext === "png" || ext === "gif" || ext === "bmp" ? ext : null;
}

const SAFE_URL = /^(https?:|mailto:)/i;

class Rels {
  private list: { id: string; type: string; target: string; external: boolean }[] = [];
  add(type: string, target: string, external = false): string {
    const id = `rId${this.list.length + 1}`;
    this.list.push({ id, type, target, external });
    return id;
  }
  get size(): number {
    return this.list.length;
  }
  xml(): string {
    const items = this.list
      .map((r) => {
        const target = r.target.replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;");
        return `<Relationship Id="${r.id}" Type="${REL_BASE}/${r.type}" Target="${target}"${r.external ? ' TargetMode="External"' : ""}/>`;
      })
      .join("");
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items}</Relationships>`;
  }
}

export interface CleanResult {
  bytes: Uint8Array;
  mediaCount: number;
  dropped: Record<string, number>;
}

/**
 * 白名單重新抄寫：由零開始砌一份新 .docx。
 * 只會搬 policy.ts 白名單入面、驗證過嘅元素同屬性；圖片清走 metadata；docProps 係全新、中性嘅。
 */
export async function writeCleanDocx(pkg: OoxmlPackage, choices: Choices): Promise<CleanResult> {
  const jszip = new JSZip();
  // 唔好自動加資料夾項目
  const zip = { file: (path: string, data: string | Uint8Array) => jszip.file(path, data, { createFolders: false }) };
  const docRels = new Rels();
  const docRelsOld = await pkg.rels(MAIN);
  const removeIds = new Set(choices.removeParagraphs);
  const keep = new Set(choices.keepImages);
  const dropped = new Map<string, number>();
  const overrides: [string, string][] = [];
  const usedImageTypes = new Set<ImageExt>();
  let mediaCount = 0;
  let drawingId = 0;
  const imageCounter = { n: 0 };
  const mediaCache = new Map<string, string>(); // 原圖或者佔位尺寸 → 新路徑

  async function addMedia(key: string, make: () => Promise<{ bytes: Uint8Array; ext: ImageExt }>): Promise<string> {
    const cached = mediaCache.get(key);
    if (cached) return cached;
    const { bytes, ext } = await make();
    const path = `media/image${++mediaCount}.${ext}`;
    zip.file(`word/${path}`, bytes);
    usedImageTypes.add(ext);
    mediaCache.set(key, path);
    return path;
  }

  /** 每個 part 有自己嘅 relationships */
  function partCtx(oldRels: Map<string, Relationship>, rels: Rels, allowGraphics: boolean, isSettings = false): SerializeCtx {
    let lastKeep = false;
    return {
      allowGraphics,
      isSettings,
      dropped,
      imageCounter,
      nextDrawingId: () => ++drawingId,
      hyperlink(old) {
        const rel = oldRels.get(old);
        if (!rel || !rel.external || !SAFE_URL.test(rel.target)) return null;
        return rels.add("hyperlink", rel.target, true);
      },
      headerFooter: () => null,
      async image(req: ImageRequest) {
        if (req.index !== null) lastKeep = keep.has(req.index);
        const rel = req.oldRelId ? oldRels.get(req.oldRelId) : undefined;
        const ext = rel && !rel.external ? imageExt(rel.target) : null;
        if (lastKeep && rel && ext) {
          const original = await pkg.bytes(rel.target);
          const stripped = original ? stripImageMetadata(original, ext) : null;
          if (stripped) {
            const path = await addMedia(`orig:${rel.target}`, async () => ({ bytes: stripped, ext }));
            return rels.add("image", path);
          }
        }
        // 冇揀保留、格式唔支援（例如 EMF）、或者讀唔到原圖：用同比例嘅佔位圖
        const key = `ph:${req.widthEmu}x${req.heightEmu}`;
        const path = await addMedia(key, async () => ({ bytes: await placeholderPng(req.widthEmu, req.heightEmu), ext: "png" }));
        return rels.add("image", path);
      },
    };
  }

  async function writePart(oldPath: string, newName: string, ctx: SerializeCtx, rels: Rels | null): Promise<boolean> {
    const root = await pkg.xml(oldPath);
    if (!root) return false;
    prepass(root, partName(oldPath), removeIds);
    zip.file(`word/${newName}`, await serializePart(root, ctx));
    if (rels && rels.size > 0) {
      const base = newName.slice(newName.lastIndexOf("/") + 1);
      const dir = newName.includes("/") ? newName.slice(0, newName.lastIndexOf("/") + 1) : "";
      zip.file(`word/${dir}_rels/${base}.rels`, rels.xml());
    }
    return true;
  }

  // 樣式、編號、設定、字型表、主題、註腳、尾註
  for (const part of CORE_PARTS) {
    const rel = [...docRelsOld.values()].find((r) => r.type === part.type && !r.external);
    if (!rel) continue;
    const rels = new Rels();
    const ctx = partCtx(await pkg.rels(rel.target), rels, false, part.type === "settings");
    if (await writePart(rel.target, part.name, ctx, rels)) {
      docRels.add(part.type, part.name);
      overrides.push([`/word/${part.name}`, part.contentType]);
    }
  }

  // 頁首頁尾：內文引用到先搬
  const headerFooterDone = new Map<string, string>();
  let headerCount = 0;
  let footerCount = 0;
  const mainCtx = partCtx(docRelsOld, docRels, true);
  const headerFooter = async (old: string): Promise<string | null> => {
    const rel = docRelsOld.get(old);
    if (!rel || rel.external || (rel.type !== "header" && rel.type !== "footer")) return null;
    const done = headerFooterDone.get(rel.target);
    if (done) return done;
    const name = rel.type === "header" ? `header${++headerCount}.xml` : `footer${++footerCount}.xml`;
    const rels = new Rels();
    if (!(await writePart(rel.target, name, partCtx(await pkg.rels(rel.target), rels, false), rels))) return null;
    overrides.push([`/word/${name}`, `${CT_BASE}.${rel.type}+xml`]);
    const id = docRels.add(rel.type, name);
    headerFooterDone.set(rel.target, id);
    return id;
  };
  // serializePart 係同步遞迴，所以頁首頁尾要預先處理：先掃描所有引用
  const docRoot = await pkg.xml(MAIN);
  if (!docRoot) throw new Error("word/document.xml 唔存在");
  const hfIds = new Map<string, string>();
  for (const ref of ["headerReference", "footerReference"]) {
    for (const el of Array.from(docRoot.getElementsByTagNameNS("http://schemas.openxmlformats.org/wordprocessingml/2006/main", ref))) {
      const old = (el as Element).getAttributeNS(REL_BASE, "id");
      if (old && !hfIds.has(old)) {
        const id = await headerFooter(old);
        if (id) hfIds.set(old, id);
      }
    }
  }
  mainCtx.headerFooter = (old) => hfIds.get(old) ?? null;

  prepass(docRoot, "document", removeIds);
  zip.file("word/document.xml", await serializePart(docRoot, mainCtx));
  zip.file("word/_rels/document.xml.rels", docRels.xml());

  // 全新、中性嘅 metadata
  const now = new Date().toISOString().replace(/\.\d{3}Z$/, "Z");
  zip.file(
    "docProps/core.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><cp:revision>1</cp:revision><dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`,
  );
  zip.file(
    "docProps/app.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"/>`,
  );
  zip.file(
    "_rels/.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL_BASE}/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="${REL_BASE}/extended-properties" Target="docProps/app.xml"/></Relationships>`,
  );
  const defaults = [
    `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>`,
    `<Default Extension="xml" ContentType="application/xml"/>`,
    ...[...usedImageTypes].map((ext) => `<Default Extension="${ext}" ContentType="${IMAGE_TYPES[ext]}"/>`),
  ];
  const allOverrides = [
    [`/word/document.xml`, `${CT_BASE}.document.main+xml`],
    ...overrides,
    ["/docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml"],
    ["/docProps/app.xml", "application/vnd.openxmlformats-officedocument.extended-properties+xml"],
  ].map(([p, t]) => `<Override PartName="${p}" ContentType="${t}"/>`);
  zip.file(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">${defaults.join("")}${allOverrides.join("")}</Types>`,
  );

  const bytes = await jszip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
  return { bytes, mediaCount, dropped: Object.fromEntries(dropped) };
}
