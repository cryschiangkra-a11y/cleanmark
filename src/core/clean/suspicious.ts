import type { Choices } from "../model";

// 可疑字元：AI 平台可能用嚟喺文字入面藏記號嘅「睇落正常」字元。
// 三樣都係一對一替換（一個 UTF-16 單位換一個），所以長度唔變，可以按原本格式段切返開，排版唔會走樣。

/** 睇落同英文字母一樣嘅西里爾 / 希臘字母。只收外形幾乎一樣嘅，避免誤改 */
const HOMOGLYPHS: Record<string, string> = {
  // 西里爾小楷
  "а": "a", "е": "e", "о": "o", "р": "p", "с": "c", "х": "x", "у": "y", "і": "i", "ј": "j", "ѕ": "s",
  "һ": "h", "ԁ": "d", "ԛ": "q", "ԝ": "w", "ӏ": "l",
  // 西里爾大楷
  "А": "A", "В": "B", "Е": "E", "К": "K", "М": "M", "Н": "H", "О": "O", "Р": "P", "С": "C", "Т": "T",
  "Х": "X", "У": "Y", "І": "I", "Ј": "J", "Ѕ": "S",
  // 希臘大楷同 omicron（其他希臘小楷例如 α、ν、ρ 喺數學同科學入面好常見，唔收）
  "Α": "A", "Β": "B", "Ε": "E", "Ζ": "Z", "Η": "H", "Ι": "I", "Κ": "K", "Μ": "M", "Ν": "N", "Ο": "O",
  "Ρ": "P", "Τ": "T", "Υ": "Y", "Χ": "X", "ο": "o",
};

const LATIN = /\p{Script=Latin}/u;
const WORD = /[\p{L}\p{M}]+/gu;

/** 撇號變體：U+02BC 修飾字母撇號、U+02B9 修飾字母 prime */
const APOSTROPHE_VARIANTS = new Set(["ʼ", "ʹ"]);

/** 特殊空格：U+00A0、U+2000–U+200A、U+205F。U+202F 喺隱形字元規則度處理；U+3000 全形空格喺中文好常見，唔計 */
function isSpecialSpace(ch: string): boolean {
  const cp = ch.charCodeAt(0);
  return cp === 0x00a0 || (cp >= 0x2000 && cp <= 0x200a) || cp === 0x205f;
}

const label = (ch: string) => "U+" + ch.charCodeAt(0).toString(16).toUpperCase().padStart(4, "0");

export interface HomoglyphHit {
  index: number;
  word: string;
  fixed: string;
}

export interface SuspiciousScan {
  homoglyphs: HomoglyphHit[];
  apostrophes: { index: number }[];
  specialSpaces: Record<string, number>; // 例如 { "U+00A0": 3 }
}

export interface SuspiciousOptions {
  homoglyphs: boolean;
  apostrophes: boolean;
  specialSpaces: boolean;
  apostrophe?: string; // 換成邊個撇號，預設 U+2019（Word 自動格式化用嘅）
}

export function scanSuspicious(text: string): SuspiciousScan {
  const homoglyphs: HomoglyphHit[] = [];
  for (const m of text.matchAll(WORD)) {
    const word = m[0];
    const chars = Array.from(word);
    // 同一個字入面有英文字母，又有睇落一樣嘅西里爾 / 希臘字母
    if (!chars.some((c) => LATIN.test(c)) || !chars.some((c) => c in HOMOGLYPHS)) continue;
    homoglyphs.push({ index: m.index!, word, fixed: chars.map((c) => HOMOGLYPHS[c] ?? c).join("") });
  }
  const apostrophes: { index: number }[] = [];
  const specialSpaces: Record<string, number> = {};
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (APOSTROPHE_VARIANTS.has(ch) && LATIN.test(text[i - 1] ?? "") && LATIN.test(text[i + 1] ?? "")) {
      apostrophes.push({ index: i });
    } else if (isSpecialSpace(ch)) {
      specialSpaces[label(ch)] = (specialSpaces[label(ch)] ?? 0) + 1;
    }
  }
  return { homoglyphs, apostrophes, specialSpaces };
}

/** 一對一替換，長度唔變 */
export function fixSuspicious(text: string, opts: SuspiciousOptions): string {
  if (!opts.homoglyphs && !opts.apostrophes && !opts.specialSpaces) return text;
  const out = text.split("");
  const scan = scanSuspicious(text);
  if (opts.homoglyphs) {
    for (const h of scan.homoglyphs) {
      for (let i = 0; i < h.word.length; i++) {
        const c = h.word[i];
        if (c in HOMOGLYPHS) out[h.index + i] = HOMOGLYPHS[c];
      }
    }
  }
  if (opts.apostrophes) for (const a of scan.apostrophes) out[a.index] = opts.apostrophe ?? "’";
  if (opts.specialSpaces) {
    for (let i = 0; i < out.length; i++) if (isSpecialSpace(out[i])) out[i] = " ";
  }
  return out.join("");
}

/** 將成段文字按原本每個格式段嘅長度切返開 */
export function splitByLengths(text: string, lengths: number[]): string[] {
  const out: string[] = [];
  let o = 0;
  for (const n of lengths) {
    out.push(text.slice(o, o + n));
    o += n;
  }
  return out;
}

/** 文件用直撇號定彎撇號多，就跟佢 */
export function preferredApostrophe(allText: string): string {
  const straight = (allText.match(/'/g) ?? []).length;
  const curly = (allText.match(/’/g) ?? []).length;
  return straight > curly ? "'" : "’";
}

/** 用家選擇 → 替換選項（同形字、撇號預設換；特殊空格預設唔換） */
export function suspiciousOptions(choices: Choices, apostrophe = "\u2019"): SuspiciousOptions {
  return {
    homoglyphs: choices.fixHomoglyphs ?? true,
    apostrophes: choices.fixApostrophes ?? true,
    specialSpaces: choices.replaceSpecialSpaces ?? false,
    apostrophe,
  };
}
