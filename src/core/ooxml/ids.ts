import { NS, type Element } from "./xml";

/** part 名，例如 "word/header2.xml" → "header2" */
export function partName(path: string): string {
  return path.slice(path.lastIndexOf("/") + 1).replace(/\.xml$/, "");
}

/**
 * 段落 ID = part 名 + 段落喺個 part 入面嘅序號（按文件次序，包括表格同文字方塊入面嘅段落），
 * 例如 "document:12"。讀取器同重新抄寫器都用呢個，令用家揀刪除嘅段落兩邊對得上。
 */
export function paragraphIds(root: Element, part: string): Map<Element, string> {
  const out = new Map<Element, string>();
  Array.from(root.getElementsByTagNameNS(NS.w, "p")).forEach((p, i) => out.set(p as Element, `${part}:${i + 1}`));
  return out;
}
