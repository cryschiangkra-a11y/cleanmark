// 檔案層面比較：將原檔同新檔逐個元素整理成標準格式再比較。
// 原檔先套用「已知會拎走」嘅規則（rsid、段落 ID、內容控制項外殼等），之後兩邊應該完全一樣。
import JSZip from "jszip";
import { DOMParser, type Element } from "@xmldom/xmldom";
import { PREFIX_OF } from "../src/core/sanitize/policy";

const PART_TYPES = ["styles", "numbering", "settings", "fontTable", "theme", "footnotes", "endnotes"];

// 原檔入面已知會被拎走嘅屬性同元素
const DROP_ATTR = /^(w:rsid\w*|w14:\w+|w15:\w+|w16\w*:\w+|wp14:(anchorId|editId)|mc:Ignorable|xml:space|r:id|r:embed|id|name|descr|title)$/;
const DROP_EL = /^(w:rsids|w:zoom|w:proofState|w:lastRenderedPageBreak|w:tmpl|w:sdtPr|w:sdtEndPr|w14:\w+|w15:\w+|w:proofErr|w:docVars|w:attachedTemplate)$/;
const UNWRAP = /^(w:sdt|w:sdtContent|w:ins|w:moveTo|w:smartTag|w:customXml)$/;

function qn(el: { namespaceURI: string | null; localName: string | null; nodeName: string }): string {
  const p = el.namespaceURI ? PREFIX_OF[el.namespaceURI] : "";
  return p ? `${p}:${el.localName}` : el.nodeName;
}

function tokens(root: Element, isOriginal: boolean): string[] {
  const out: string[] = [];
  const walk = (el: Element) => {
    const name = qn(el);
    if (isOriginal && DROP_EL.test(name)) return;
    if (UNWRAP.test(name)) {
      for (let n = el.firstChild; n; n = n.nextSibling) if (n.nodeType === 1) walk(n as Element);
      return;
    }
    const attrs: string[] = [];
    for (let i = 0; i < el.attributes.length; i++) {
      const a = el.attributes.item(i)!;
      if (a.name.startsWith("xmlns")) continue;
      const an = qn(a);
      if (DROP_ATTR.test(an)) continue;
      attrs.push(`${an}=${a.value}`);
    }
    out.push(`<${name} ${attrs.sort().join(" ")}>`);
    for (let n = el.firstChild; n; n = n.nextSibling) {
      if (n.nodeType === 1) walk(n as Element);
      else if (n.nodeType === 3 || n.nodeType === 4) {
        // 文字元素（w:t、m:t 等）入面嘅空格都要比較；其他地方只係排版用嘅換行縮排
        const isTextEl = /:(t|instrText|posOffset|align)$/.test(name);
        if (isTextEl || (n.nodeValue ?? "").trim() !== "") out.push(`"${n.nodeValue}"`);
      }
    }
    out.push(`</${name}>`);
  };
  walk(root);
  return out;
}

async function relTargets(zip: JSZip): Promise<Map<string, string>> {
  const rels = (await zip.file("word/_rels/document.xml.rels")?.async("string")) ?? "";
  const map = new Map<string, string>();
  for (const m of rels.matchAll(/Type="[^"]*\/(\w+)"\s+Target="([^"]+)"/g)) if (!map.has(m[1])) map.set(m[1], `word/${m[2]}`);
  for (const m of rels.matchAll(/Target="([^"]+)"[^>]*Type="[^"]*\/(\w+)"/g)) if (!map.has(m[2])) map.set(m[2], `word/${m[1]}`);
  return map;
}

export async function xmlDiff(original: Uint8Array, cleaned: Uint8Array): Promise<string[]> {
  const [za, zb] = [await JSZip.loadAsync(original), await JSZip.loadAsync(cleaned)];
  const [ra, rb] = [await relTargets(za), await relTargets(zb)];
  const report: string[] = [];
  const parts: [string, string, string][] = [["document", "word/document.xml", "word/document.xml"]];
  for (const t of PART_TYPES) if (ra.has(t) && rb.has(t)) parts.push([t, ra.get(t)!, rb.get(t)!]);
  for (const t of PART_TYPES) if (ra.has(t) && !rb.has(t)) report.push(`${t}：新檔冇呢個 part`);
  for (const [label, pa, pb] of parts) {
    const parse = async (z: JSZip, p: string) =>
      new DOMParser().parseFromString((await z.file(p)!.async("string")), "text/xml").documentElement as unknown as Element;
    const ta = tokens(await parse(za, pa), true);
    const tb = tokens(await parse(zb, pb), false);
    let i = 0;
    while (i < ta.length && i < tb.length && ta[i] === tb[i]) i++;
    if (i === ta.length && i === tb.length) report.push(`${label}：${ta.length} 個元素同文字，完全一樣`);
    else report.push(`${label}：第 ${i} 個開始唔同\n  原檔：${ta.slice(i, i + 3).join(" ")}\n  新檔：${tb.slice(i, i + 3).join(" ")}`);
  }
  return report;
}
