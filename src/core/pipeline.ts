import { inspectDocx } from "./inspect/docx";
import type { Choices, CmDocument, Report } from "./model";
import { OoxmlPackage, ZipBudgetError } from "./ooxml/package";
import { readDocx } from "./readers/docx";
import { writeCleanDocx } from "./sanitize/package";
import { selfCheckDocx, type Problem } from "./verify/self-check-docx";

export const MAX_BYTES = 50 * 1024 * 1024;

export type ErrorCode =
  | "legacyFormat" // .doc / .ppt
  | "unsupported" // 其他格式（只支援 Word）
  | "tooLarge"
  | "encrypted"
  | "corrupt"
  | "unsafeArchive" // 解壓之後大得唔正常（壓縮炸彈）
  | "selfCheckFailed";

export class CleanMarkError extends Error {
  constructor(
    public code: ErrorCode,
    public problems: Problem[] = [],
  ) {
    super(code);
  }
}

const ZIP = [0x50, 0x4b, 0x03, 0x04];
const CFB = [0xd0, 0xcf, 0x11, 0xe0]; // 舊 Office 格式、或者加咗密碼嘅新格式

const startsWith = (b: Uint8Array, sig: number[]) => sig.every((x, i) => b[i] === x);

export function extensionOf(name: string): string {
  const i = name.lastIndexOf(".");
  return i < 0 ? "" : name.slice(i + 1).toLowerCase();
}

export function checkInput(name: string, bytes: Uint8Array): "docx" {
  const ext = extensionOf(name);
  if (ext === "doc" || ext === "ppt") throw new CleanMarkError("legacyFormat");
  if (ext !== "docx") throw new CleanMarkError("unsupported");
  if (bytes.length > MAX_BYTES) throw new CleanMarkError("tooLarge");
  if (startsWith(bytes, CFB)) throw new CleanMarkError("encrypted");
  if (!startsWith(bytes, ZIP)) throw new CleanMarkError("corrupt");
  return "docx";
}

export interface Analysis {
  doc: CmDocument;
  report: Report;
}

export async function analyze(name: string, bytes: Uint8Array): Promise<Analysis> {
  checkInput(name, bytes);
  try {
    const { doc, pkg } = await readDocx(bytes);
    return { doc, report: await inspectDocx(pkg, doc) };
  } catch (e) {
    if (e instanceof CleanMarkError) throw e;
    if (e instanceof ZipBudgetError) throw new CleanMarkError("unsafeArchive");
    throw new CleanMarkError("corrupt");
  }
}

export interface BuiltFile {
  name: string;
  bytes: Uint8Array;
}

/** 由原檔白名單重新抄寫一份新檔案，自我檢查通過先回傳 */
export async function build(name: string, original: Uint8Array, choices: Choices): Promise<BuiltFile> {
  let out: Awaited<ReturnType<typeof writeCleanDocx>>;
  try {
    out = await writeCleanDocx(await OoxmlPackage.open(original), choices);
  } catch (e) {
    if (e instanceof ZipBudgetError) throw new CleanMarkError("unsafeArchive");
    throw e;
  }
  const check = await selfCheckDocx(out.bytes, out.mediaCount);
  if (!check.ok) throw new CleanMarkError("selfCheckFailed", check.problems);
  return { name: outputName(name, "docx"), bytes: out.bytes };
}

export function outputName(name: string, ext: string): string {
  const base = name.includes(".") ? name.slice(0, name.lastIndexOf(".")) : name;
  return `${base}-clean.${ext}`;
}
