import type { HiddenCharSummary } from "../model";

// 規則見設計文件 5.6

function inRanges(cp: number, ranges: [number, number][]): boolean {
  return ranges.some(([lo, hi]) => cp >= lo && cp <= hi);
}

const ALWAYS_REMOVE: [number, number][] = [
  [0x00ad, 0x00ad], // 軟連字號
  [0x180e, 0x180e], // 蒙古文元音分隔符
  [0x200b, 0x200b], // 零寬度空格
  [0x202d, 0x202e], // 方向覆寫
  [0x2060, 0x2064], // 字詞連接符、隱形運算符
  [0xfeff, 0xfeff], // BOM
  [0xe0000, 0xe007f], // Tag 字元
];

// 只有窄不換行空格：有報道指某啲 AI 會用佢做記號。
// 全形空格、細空格等係正常排版用嘅空格（例如圖表目錄、方程式），原樣保留，否則排版會走樣
const REPLACE_WITH_SPACE: [number, number][] = [[0x202f, 0x202f]];

const JOINERS: [number, number][] = [[0x200c, 0x200d]]; // ZWNJ、ZWJ
const BIDI_MARKS: [number, number][] = [
  [0x061c, 0x061c],
  [0x200e, 0x200f],
  [0x202a, 0x202c],
  [0x2066, 0x2069],
];
const VARIATION_SELECTORS: [number, number][] = [[0xfe00, 0xfe0f]];

// 需要 ZWJ / ZWNJ 嘅文字同 emoji
const NEEDS_JOINER =
  /[\p{Script=Arabic}\p{Script=Syriac}\p{Script=Nko}\p{Script=Mongolian}\p{Script=Devanagari}\p{Script=Bengali}\p{Script=Gurmukhi}\p{Script=Gujarati}\p{Script=Oriya}\p{Script=Tamil}\p{Script=Telugu}\p{Script=Kannada}\p{Script=Malayalam}\p{Script=Sinhala}\p{Script=Khmer}\p{Script=Myanmar}\p{Extended_Pictographic}\p{Emoji_Modifier}\uFE0F]/u;
const NEEDS_VARIATION = /[\p{Extended_Pictographic}\p{Script=Han}0-9#*]/u;
const RTL = /[\p{Script=Arabic}\p{Script=Hebrew}\p{Script=Syriac}\p{Script=Thaana}\p{Script=Nko}]/u;

export function emptySummary(): HiddenCharSummary {
  return { removed: {}, replaced: {}, kept: 0 };
}

export function mergeSummary(a: HiddenCharSummary, b: HiddenCharSummary): HiddenCharSummary {
  const add = (x: Record<string, number>, y: Record<string, number>) => {
    const out = { ...x };
    for (const [k, v] of Object.entries(y)) out[k] = (out[k] ?? 0) + v;
    return out;
  };
  return { removed: add(a.removed, b.removed), replaced: add(a.replaced, b.replaced), kept: a.kept + b.kept };
}

export function hasRtl(text: string): boolean {
  return RTL.test(text);
}

function label(cp: number): string {
  return "U+" + cp.toString(16).toUpperCase().padStart(4, "0");
}

export interface CleanContext {
  before: string; // 前面嘅文字（同一段落）
  after: string; // 後面嘅文字（同一段落）
  rtl: boolean; // 段落有冇由右至左文字
}

export function cleanText(
  text: string,
  ctx: CleanContext = { before: "", after: "", rtl: hasRtl(text) },
): { text: string; summary: HiddenCharSummary } {
  const cps = Array.from(text);
  const prevOf = (i: number) => (i > 0 ? cps[i - 1] : Array.from(ctx.before).at(-1) ?? "");
  const nextOf = (i: number) => (i < cps.length - 1 ? cps[i + 1] : Array.from(ctx.after)[0] ?? "");
  const summary = emptySummary();
  let out = "";
  const bump = (bucket: Record<string, number>, cp: number) => {
    bucket[label(cp)] = (bucket[label(cp)] ?? 0) + 1;
  };

  cps.forEach((ch, i) => {
    const cp = ch.codePointAt(0)!;
    if (inRanges(cp, ALWAYS_REMOVE)) {
      bump(summary.removed, cp);
    } else if (inRanges(cp, REPLACE_WITH_SPACE)) {
      bump(summary.replaced, cp);
      out += " ";
    } else if (inRanges(cp, JOINERS)) {
      if (NEEDS_JOINER.test(prevOf(i)) || NEEDS_JOINER.test(nextOf(i))) {
        summary.kept++;
        out += ch;
      } else bump(summary.removed, cp);
    } else if (inRanges(cp, BIDI_MARKS)) {
      if (ctx.rtl) {
        summary.kept++;
        out += ch;
      } else bump(summary.removed, cp);
    } else if (inRanges(cp, VARIATION_SELECTORS)) {
      if (NEEDS_VARIATION.test(prevOf(i))) {
        summary.kept++;
        out += ch;
      } else bump(summary.removed, cp);
    } else {
      out += ch;
    }
  });
  return { text: out, summary };
}
