import type { Block, ListLevel } from "../core/model";

function roman(n: number): string {
  const table: [number, string][] = [
    [1000, "m"], [900, "cm"], [500, "d"], [400, "cd"], [100, "c"], [90, "xc"],
    [50, "l"], [40, "xl"], [10, "x"], [9, "ix"], [5, "v"], [4, "iv"], [1, "i"],
  ];
  let out = "";
  for (const [v, s] of table) while (n >= v) { out += s; n -= v; }
  return out;
}

function letters(n: number): string {
  let out = "";
  while (n > 0) { n--; out = String.fromCharCode(97 + (n % 26)) + out; n = Math.floor(n / 26); }
  return out;
}

export function formatNumber(n: number, format: string): string {
  switch (format) {
    case "lowerLetter": return letters(n);
    case "upperLetter": return letters(n).toUpperCase();
    case "lowerRoman": return roman(n);
    case "upperRoman": return roman(n).toUpperCase();
    case "decimalZero": return n < 10 ? `0${n}` : String(n);
    default: return String(n);
  }
}

/** 預覽用：計出每個列表段落前面嘅編號或者符號，key = 段落 id */
export function listLabels(blocks: Iterable<Block>, lists: Record<string, ListLevel[]>): Map<string, string> {
  const counters = new Map<string, number[]>();
  const out = new Map<string, string>();
  for (const b of blocks) {
    if (b.type !== "paragraph" || !b.list) continue;
    const levels = lists[b.list.numId];
    if (!levels) continue;
    const { level } = b.list;
    const c = counters.get(b.list.numId) ?? levels.map((l) => l.start - 1);
    c[level] = (c[level] ?? 0) + 1;
    for (let i = level + 1; i < c.length; i++) c[i] = levels[i].start - 1;
    counters.set(b.list.numId, c);
    const def = levels[level];
    out.set(
      b.id,
      def.format === "bullet"
        ? def.text
        : def.format === "none"
          ? ""
          : def.text.replace(/%(\d)/g, (_, d) => formatNumber(Math.max(c[Number(d) - 1] ?? 1, 1), levels[Number(d) - 1]?.format ?? "decimal")),
    );
  }
  return out;
}
