import { hasC2pa, imageTags } from "../inspect/bytes";
import { inspectDocx } from "../inspect/docx";
import { readOoxmlMetadata } from "../inspect/metadata";
import { findVisibleDocx } from "../inspect/visible-docx";
import { readDocx } from "../readers/docx";

export type ProblemCode =
  | "unreadable"
  | "unexpectedPart"
  | "unexpectedMedia"
  | "metadata"
  | "visible"
  | "c2pa"
  | "imageMetadata"
  | "hiddenChars";

export interface Problem {
  code: ProblemCode;
  detail: string;
}

export interface SelfCheckResult {
  ok: boolean;
  problems: Problem[];
}

/** 輸出器應該產生嘅檔案，其他一律當有問題 */
const ALLOWED_PARTS: RegExp[] = [
  /^\[Content_Types\]\.xml$/,
  /^_rels\/\.rels$/,
  /^docProps\/(core|app|custom)\.xml$/,
  /^word\/(document|styles|settings|fontTable|numbering|footnotes|endnotes)\.xml$/,
  /^word\/(header|footer)\d+\.xml$/,
  /^word\/theme\/theme1\.xml$/,
  /^word\/media\/[^/]+$/,
  /^word\/_rels\/[^/]+\.xml\.rels$/,
];

/** 只准有建立時間、修改時間同版本號 */
const ALLOWED_METADATA = new Set(["cp:revision", "dcterms:created", "dcterms:modified"]);

/** expectedMedia：重新抄寫器話寫咗幾多張圖（保留嘅原圖 + 佔位圖） */
export async function selfCheckDocx(bytes: Uint8Array, expectedMedia: number): Promise<SelfCheckResult> {
  const problems: Problem[] = [];
  let read: Awaited<ReturnType<typeof readDocx>>;
  try {
    read = await readDocx(bytes);
  } catch (e) {
    return { ok: false, problems: [{ code: "unreadable", detail: String(e) }] };
  }
  const { doc, pkg } = read;

  const paths = pkg.paths();
  for (const path of paths) {
    if (!ALLOWED_PARTS.some((re) => re.test(path))) problems.push({ code: "unexpectedPart", detail: path });
  }
  const mediaCount = paths.filter((x) => x.startsWith("word/media/")).length;
  if (mediaCount > expectedMedia) problems.push({ code: "unexpectedMedia", detail: `${mediaCount} > ${expectedMedia}` });

  for (const m of await readOoxmlMetadata(pkg)) {
    if (!ALLOWED_METADATA.has(m.field)) problems.push({ code: "metadata", detail: `${m.field}=${m.value}` });
  }
  for (const v of await findVisibleDocx(pkg)) problems.push({ code: "visible", detail: `${v.code} ${v.part}` });

  for (const path of paths) {
    const b = await pkg.bytes(path);
    if (b && hasC2pa(b)) problems.push({ code: "c2pa", detail: path });
    // 圖片已經清走 metadata，唔應該仲有 EXIF / XMP / C2PA
    if (b && path.startsWith("word/media/") && imageTags(b).length > 0) {
      problems.push({ code: "imageMetadata", detail: `${path} ${imageTags(b).join("+")}` });
    }
  }

  const { hidden } = await inspectDocx(pkg, doc);
  const leftover = [...Object.entries(hidden.removed), ...Object.entries(hidden.replaced)];
  for (const [cp, n] of leftover) problems.push({ code: "hiddenChars", detail: `${cp}×${n}` });

  return { ok: problems.length === 0, problems };
}
