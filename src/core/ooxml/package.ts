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
