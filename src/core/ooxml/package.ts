import JSZip from "jszip";
import { attr, kids, NS, parseXml, type Element } from "./xml";

export interface Relationship {
  id: string;
  type: string; // 最後一節，例如 "image"、"hyperlink"、"header"
  target: string; // 已解析成套件內嘅完整路徑；外部連結就係原網址
  external: boolean;
}

/** 防壓縮炸彈：細細個嘅 .docx 解壓之後可以大到幾 GB，令網頁卡死 */
export class ZipBudgetError extends Error {}

export interface ZipLimits {
  maxTotalBytes: number; // 所有內部檔案解壓之後加埋嘅上限
  maxEntries: number; // 內部檔案數目上限
}

/** 正常 Word 檔遠遠唔會去到呢個數 */
export const ZIP_LIMITS: ZipLimits = { maxTotalBytes: 300 * 1024 * 1024, maxEntries: 5000 };

function concatChunks(chunks: Uint8Array[], size: number): Uint8Array {
  const out = new Uint8Array(size);
  let o = 0;
  for (const c of chunks) {
    out.set(c, o);
    o += c.length;
  }
  return out;
}

interface ZipStream {
  on(event: "data", cb: (chunk: Uint8Array) => void): ZipStream;
  on(event: "error", cb: (e: Error) => void): ZipStream;
  on(event: "end", cb: () => void): ZipStream;
  pause(): ZipStream;
  resume(): ZipStream;
}

export class OoxmlPackage {
  private used = 0;
  private cache = new Map<string, Uint8Array>();

  private constructor(
    private zip: JSZip,
    private limits: ZipLimits,
  ) {}

  static async open(bytes: Uint8Array, limits: ZipLimits = ZIP_LIMITS): Promise<OoxmlPackage> {
    const zip = await JSZip.loadAsync(bytes);
    const files = Object.values(zip.files).filter((f) => !f.dir);
    if (files.length > limits.maxEntries) throw new ZipBudgetError(`too many entries: ${files.length}`);
    // 快速檢查：壓縮檔自己報嘅大細（可以作假，所以真正嘅限制喺解壓時計）
    for (const f of files) {
      const declared = (f as unknown as { _data?: { uncompressedSize?: number } })._data?.uncompressedSize ?? 0;
      if (declared > limits.maxTotalBytes) throw new ZipBudgetError(`entry too large: ${f.name}`);
    }
    return new OoxmlPackage(zip, limits);
  }

  /** 解壓一個內部檔案，邊解壓邊計實際大細，超過上限即刻停止 */
  private read(path: string): Promise<Uint8Array | undefined> {
    const cached = this.cache.get(path);
    if (cached) return Promise.resolve(cached);
    const f = this.zip.file(path);
    if (!f || f.dir) return Promise.resolve(undefined);
    return new Promise((resolve, reject) => {
      const chunks: Uint8Array[] = [];
      let size = 0;
      let stopped = false;
      // JSZip 有 internalStream，但型別定義冇寫
      const stream = (f as unknown as { internalStream(t: "uint8array"): ZipStream }).internalStream("uint8array");
      stream
        .on("data", (chunk: Uint8Array) => {
          if (stopped) return;
          size += chunk.length;
          this.used += chunk.length;
          if (this.used > this.limits.maxTotalBytes) {
            stopped = true;
            stream.pause();
            reject(new ZipBudgetError(`decompressed size exceeds ${this.limits.maxTotalBytes} bytes`));
            return;
          }
          chunks.push(chunk);
        })
        .on("error", (e: Error) => {
          if (!stopped) reject(e);
        })
        .on("end", () => {
          if (stopped) return;
          const bytes = concatChunks(chunks, size);
          this.cache.set(path, bytes);
          resolve(bytes);
        })
        .resume();
    });
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
    const b = await this.read(path);
    return b === undefined ? undefined : new TextDecoder().decode(b);
  }

  async bytes(path: string): Promise<Uint8Array | undefined> {
    return this.read(path);
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
