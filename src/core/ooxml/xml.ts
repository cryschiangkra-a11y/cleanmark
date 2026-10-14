import { DOMParser, type Element } from "@xmldom/xmldom";

export type { Element };

export const NS = {
  w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  rel: "http://schemas.openxmlformats.org/package/2006/relationships",
  wp: "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
  a: "http://schemas.openxmlformats.org/drawingml/2006/main",
  wps: "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
  mc: "http://schemas.openxmlformats.org/markup-compatibility/2006",
  v: "urn:schemas-microsoft-com:vml",
  cp: "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
  ep: "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
  custom: "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties",
} as const;

export function parseXml(text: string): Element {
  const doc = new DOMParser({
    onError: (level, msg) => {
      if (level === "fatalError") throw new Error(`XML parse error: ${msg}`);
    },
  }).parseFromString(text, "text/xml");
  const root = doc.documentElement;
  if (!root) throw new Error("XML parse error: no root element");
  return root;
}

/** 直接子元素 */
export function kids(el: Element, ns?: string, local?: string): Element[] {
  const out: Element[] = [];
  for (let n = el.firstChild; n; n = n.nextSibling) {
    if (n.nodeType !== 1) continue;
    const e = n as Element;
    if (ns !== undefined && e.namespaceURI !== ns) continue;
    if (local !== undefined && e.localName !== local) continue;
    out.push(e);
  }
  return out;
}

export function kid(el: Element | undefined, ns: string, local: string): Element | undefined {
  return el ? kids(el, ns, local)[0] : undefined;
}

/** 所有後代元素 */
export function desc(el: Element, ns: string, local: string): Element[] {
  return Array.from(el.getElementsByTagNameNS(ns, local)) as Element[];
}

export function attr(el: Element | undefined, ns: string | null, local: string): string | undefined {
  if (!el) return undefined;
  const v = ns ? el.getAttributeNS(ns, local) : el.getAttribute(local);
  return v === null || v === "" ? undefined : v;
}

/** w:val 屬性 */
export function wval(el: Element | undefined): string | undefined {
  return attr(el, NS.w, "val");
}

export function wnum(el: Element | undefined, local = "val"): number | undefined {
  const v = attr(el, NS.w, local);
  if (v === undefined) return undefined;
  const n = Number(v);
  return Number.isFinite(n) ? n : undefined;
}

/** OOXML 開關屬性：<w:b/> = true，<w:b w:val="0"/> = false，冇元素 = undefined */
export function toggle(el: Element | undefined): boolean | undefined {
  if (!el) return undefined;
  const v = wval(el);
  return !(v === "0" || v === "false" || v === "off" || v === "none");
}

export function textContent(el: Element): string {
  return el.textContent ?? "";
}
