import { cleanText, hasRtl } from "../clean/hidden-chars";
import { fixSuspicious, splitByLengths, type SuspiciousOptions } from "../clean/suspicious";
import { paragraphIds } from "../ooxml/ids";
import { kids, NS, type Element } from "../ooxml/xml";
import { FIELDS } from "./policy";

const NS_M = "http://schemas.openxmlformats.org/officeDocument/2006/math";

/** 文件次序列出所有元素 */
function allElements(root: Element): Element[] {
  const out: Element[] = [];
  const walk = (el: Element) => {
    out.push(el);
    for (const c of kids(el)) walk(c);
  };
  walk(root);
  return out;
}

function closestParagraph(el: Element): Element | null {
  for (let n = el.parentNode as Element | null; n; n = n.parentNode as Element | null) {
    if (n.namespaceURI === NS.w && n.localName === "p") return n;
  }
  return null;
}

/** 刪除用家揀咗嘅段落。有章節設定嘅段落、或者係儲存格唯一段落，就只清空內容 */
function removeParagraphs(root: Element, part: string, remove: Set<string>): void {
  for (const [p, id] of paragraphIds(root, part)) {
    if (!remove.has(id)) continue;
    const pPr = kids(p, NS.w, "pPr")[0];
    const hasSection = !!pPr && kids(pPr, NS.w, "sectPr").length > 0;
    const parent = p.parentNode as Element | null;
    const onlyInCell = parent?.localName === "tc" && kids(parent, NS.w, "p").length === 1;
    if (hasSection || onlyInCell || !parent) {
      for (const c of kids(p)) if (c !== pPr) p.removeChild(c);
    } else parent.removeChild(p);
  }
}

/** 唔准嘅欄位：清空指令，只保留顯示結果 */
function filterFields(all: Element[]): void {
  const stack: { instr: Element[]; decided: boolean }[] = [];
  const decide = (f: { instr: Element[]; decided: boolean }) => {
    if (f.decided) return;
    f.decided = true;
    const type = f.instr.map((e) => e.textContent ?? "").join("").trim().split(/\s+/)[0]?.toUpperCase() ?? "";
    if (!FIELDS.has(type)) for (const e of f.instr) e.textContent = "";
  };
  for (const el of all) {
    if (el.namespaceURI !== NS.w) continue;
    if (el.localName === "fldChar") {
      const t = el.getAttributeNS(NS.w, "fldCharType");
      if (t === "begin") stack.push({ instr: [], decided: false });
      else if (t === "separate" && stack.length) decide(stack[stack.length - 1]);
      else if (t === "end" && stack.length) decide(stack.pop()!);
    } else if (el.localName === "instrText" && stack.length) stack[stack.length - 1].instr.push(el);
  }
}

/** 按段落上下文清隱形字元（ZWJ、方向標記要睇前後文字） */
function cleanHiddenChars(all: Element[]): void {
  const byParagraph = new Map<Element | null, Element[]>();
  for (const el of all) {
    const isText = (el.namespaceURI === NS.w && el.localName === "t") || (el.namespaceURI === NS_M && el.localName === "t");
    if (!isText) continue;
    const p = closestParagraph(el);
    if (!byParagraph.has(p)) byParagraph.set(p, []);
    byParagraph.get(p)!.push(el);
  }
  for (const nodes of byParagraph.values()) {
    const texts = nodes.map((n) => n.textContent ?? "");
    const rtl = hasRtl(texts.join(""));
    nodes.forEach((n, i) => {
      // 方程式文字唔同前後文字連接，所以唔用上下文
      const math = n.namespaceURI === NS_M;
      const r = cleanText(texts[i], {
        before: math ? "" : texts.slice(0, i).join(""),
        after: math ? "" : texts.slice(i + 1).join(""),
        rtl,
      });
      if (r.text !== texts[i]) n.textContent = r.text;
    });
  }
}

/** 換可疑字元：按段落成段處理（一個字可能跨幾個格式段）；只處理內文 w:t，方程式唔郁 */
function fixSuspiciousChars(all: Element[], opts: SuspiciousOptions): void {
  const byParagraph = new Map<Element | null, Element[]>();
  for (const el of all) {
    if (el.namespaceURI !== NS.w || el.localName !== "t") continue;
    const p = closestParagraph(el);
    if (!byParagraph.has(p)) byParagraph.set(p, []);
    byParagraph.get(p)!.push(el);
  }
  for (const nodes of byParagraph.values()) {
    const texts = nodes.map((n) => n.textContent ?? "");
    const joined = texts.join("");
    const fixed = fixSuspicious(joined, opts);
    if (fixed === joined) continue;
    splitByLengths(fixed, texts.map((t) => t.length)).forEach((t, i) => {
      if (t !== texts[i]) nodes[i].textContent = t;
    });
  }
}

/** 全文用直撇號定彎撇號多（俾 preferredApostrophe 用） */
export function allBodyText(root: Element): string {
  return Array.from(root.getElementsByTagNameNS(NS.w, "t"))
    .map((t) => t.textContent ?? "")
    .join("");
}

export function prepass(root: Element, part: string, removeIds: Set<string>, fix: SuspiciousOptions): void {
  removeParagraphs(root, part, removeIds);
  const all = allElements(root);
  filterFields(all);
  fixSuspiciousChars(all, fix); // 一對一替換，要喺清隱形字元（會改長度）之前做
  cleanHiddenChars(all);
}
