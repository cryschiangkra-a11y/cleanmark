import type { MathNode } from "./model";
import { NS, type Element } from "./ooxml/xml";

export type { MathNode };

const NS_M = "http://schemas.openxmlformats.org/officeDocument/2006/math";

// 方程式入面只准有呢啲 Word 字型格式，其他（圖形、書籤等）一律唔要
const RPR_CHILDREN = new Set(["rFonts", "b", "i", "sz", "szCs", "color"]);

function nsAttrs(el: Element, ns: string, prefix: string): Record<string, string> {
  const out: Record<string, string> = {};
  for (let i = 0; i < el.attributes.length; i++) {
    const a = el.attributes.item(i)!;
    if (a.namespaceURI === ns && a.localName) out[`${prefix}:${a.localName}`] = a.value;
  }
  return out;
}

function sanitizeRPr(rPr: Element): MathNode {
  const children: MathNode[] = [];
  for (let n = rPr.firstChild; n; n = n.nextSibling) {
    if (n.nodeType !== 1) continue;
    const e = n as Element;
    if (e.namespaceURI === NS.w && e.localName && RPR_CHILDREN.has(e.localName)) {
      children.push({ tag: `w:${e.localName}`, attrs: nsAttrs(e, NS.w, "w"), children: [] });
    }
  }
  return { tag: "w:rPr", attrs: {}, children };
}

/** 將 OMML 元素轉成白名單樹：只留 m: 元素、m:t 文字，同 m:r / m:ctrlPr 入面嘅基本字型格式 */
export function sanitizeMath(el: Element): MathNode {
  const local = el.localName ?? "";
  const node: MathNode = { tag: `m:${local}`, attrs: nsAttrs(el, NS_M, "m"), children: [] };
  if (local === "t") {
    node.children.push(el.textContent ?? "");
    return node;
  }
  for (let n = el.firstChild; n; n = n.nextSibling) {
    if (n.nodeType !== 1) continue;
    const c = n as Element;
    if (c.namespaceURI === NS_M) node.children.push(sanitizeMath(c));
    else if (c.namespaceURI === NS.w && c.localName === "rPr" && (local === "r" || local === "ctrlPr")) {
      node.children.push(sanitizeRPr(c));
    }
  }
  return node;
}

export function isMathElement(el: Element): boolean {
  return el.namespaceURI === NS_M && (el.localName === "oMath" || el.localName === "oMathPara");
}

const escape = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

export function mathToXml(node: MathNode): string {
  const attrs = { ...node.attrs, ...(node.tag === "m:t" ? { "xml:space": "preserve" } : {}) };
  const attrText = Object.entries(attrs)
    .map(([k, v]) => ` ${k}="${escape(v)}"`)
    .join("");
  if (node.children.length === 0) return `<${node.tag}${attrText}/>`;
  const inner = node.children.map((c) => (typeof c === "string" ? escape(c) : mathToXml(c))).join("");
  return `<${node.tag}${attrText}>${inner}</${node.tag}>`;
}

/** 對每段文字執行 fn，回傳新樹（唔改原本） */
export function mapMathText(node: MathNode, fn: (text: string) => string): MathNode {
  return {
    ...node,
    attrs: { ...node.attrs },
    children: node.children.map((c) => (typeof c === "string" ? fn(c) : mapMathText(c, fn))),
  };
}

export function mathPlainText(node: MathNode): string {
  return node.children.map((c) => (typeof c === "string" ? c : mathPlainText(c))).join("");
}
