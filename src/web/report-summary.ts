import type { Report } from "../core/model";

export function hiddenTotal(r: Report): number {
  const sum = (x: Record<string, number>) => Object.values(x).reduce((a, b) => a + b, 0);
  return sum(r.hidden.removed) + sum(r.hidden.replaced);
}

/** 檔案資料行顯示嘅「找到 N 項」：隱形字元當一項計 */
export function findingCount(r: Report): number {
  return r.visible.length + r.metadata.length + r.c2pa.length + (hiddenTotal(r) > 0 ? 1 : 0) + r.phrases.length;
}
