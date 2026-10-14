import type { ListLevel } from "../../model";
import { attr, kid, kids, NS, wnum, wval, type Element } from "../../ooxml/xml";

const BULLETS = ["•", "◦", "▪"];

/** Word 常用 Symbol / Wingdings 字型嘅私人區字元做項目符號，冇原字型就會變亂碼，所以換成通用符號 */
function normaliseBullet(text: string, level: number): string {
  const cp = text.codePointAt(0);
  if (!text || (cp !== undefined && cp >= 0xe000 && cp <= 0xf8ff)) return BULLETS[level % 3];
  return text;
}

function readLevel(lvl: Element, level: number): ListLevel {
  const format = wval(kid(lvl, NS.w, "numFmt")) ?? "decimal";
  const rawText = wval(kid(lvl, NS.w, "lvlText")) ?? "";
  const ind = kid(kid(lvl, NS.w, "pPr"), NS.w, "ind");
  return {
    format,
    text: format === "bullet" ? normaliseBullet(rawText, level) : rawText,
    start: wnum(kid(lvl, NS.w, "start")) ?? 1,
    indentLeft: wnum(ind, "left") ?? wnum(ind, "start") ?? 720 * (level + 1),
    hanging: wnum(ind, "hanging") ?? 360,
  };
}

/** numId → 九層設定 */
export function readNumbering(root: Element | undefined): Record<string, ListLevel[]> {
  const out: Record<string, ListLevel[]> = {};
  if (!root) return out;
  const abstracts = new Map<string, Element>();
  for (const a of kids(root, NS.w, "abstractNum")) {
    const id = attr(a, NS.w, "abstractNumId");
    if (id !== undefined) abstracts.set(id, a);
  }
  for (const num of kids(root, NS.w, "num")) {
    const numId = attr(num, NS.w, "numId");
    const absId = wval(kid(num, NS.w, "abstractNumId"));
    const abs = absId !== undefined ? abstracts.get(absId) : undefined;
    if (numId === undefined || !abs) continue;
    const levels: ListLevel[] = [];
    for (let i = 0; i < 9; i++) {
      const lvl = kids(abs, NS.w, "lvl").find((l) => attr(l, NS.w, "ilvl") === String(i));
      levels.push(
        lvl
          ? readLevel(lvl, i)
          : { format: "bullet", text: BULLETS[i % 3], start: 1, indentLeft: 720 * (i + 1), hanging: 360 },
      );
    }
    for (const ov of kids(num, NS.w, "lvlOverride")) {
      const i = Number(attr(ov, NS.w, "ilvl"));
      const start = wnum(kid(ov, NS.w, "startOverride"));
      if (levels[i] && start !== undefined) levels[i].start = start;
    }
    out[numId] = levels;
  }
  return out;
}
