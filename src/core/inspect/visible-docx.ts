import type { VisibleItem } from "../model";
import type { OoxmlPackage } from "../ooxml/package";
import { attr, desc, NS, type Element } from "../ooxml/xml";

/** 唔計 mc:Fallback 入面嘅嘢，避免同 mc:Choice 重複 */
function inFallback(el: Element): boolean {
  for (let n = el.parentNode; n; n = n.parentNode) {
    const e = n as Element;
    if (e.namespaceURI === NS.mc && e.localName === "Fallback") return true;
  }
  return false;
}

function boxText(el: Element): string | undefined {
  const t = desc(el, NS.w, "t").map((x) => x.textContent ?? "").join("").trim();
  return t || undefined;
}

function scanVml(part: string, root: Element, inHeader: boolean): VisibleItem[] {
  const out: VisibleItem[] = [];
  for (const shape of desc(root, NS.v, "shape")) {
    if (inFallback(shape)) continue;
    const id = attr(shape, null, "id") ?? "";
    const textpath = desc(shape, NS.v, "textpath")[0];
    const text = attr(textpath, null, "string");
    if (id.startsWith("PowerPlusWaterMarkObject")) out.push({ code: "wordTextWatermark", part, text });
    else if (id.startsWith("WordPictureWatermark")) out.push({ code: "wordPictureWatermark", part });
    else if (textpath) out.push({ code: inHeader ? "headerShape" : "textArt", part, text });
    else if (inHeader) {
      out.push(desc(shape, NS.v, "imagedata").length > 0 ? { code: "headerImage", part } : { code: "headerShape", part, text: boxText(shape) });
    }
  }
  return out;
}

function scanDrawings(part: string, root: Element, inHeader: boolean): VisibleItem[] {
  const out: VisibleItem[] = [];
  for (const d of desc(root, NS.w, "drawing")) {
    if (inFallback(d)) continue;
    const anchor = desc(d, NS.wp, "anchor")[0];
    const behind = ["1", "true"].includes(attr(anchor, null, "behindDoc") ?? "");
    const isPicture = desc(d, NS.a, "blip").length > 0;
    if (inHeader) out.push(isPicture ? { code: "headerImage", part } : { code: "headerShape", part, text: boxText(d) });
    else if (behind) out.push({ code: "behindText", part, text: boxText(d) });
  }
  return out;
}

export async function findVisibleDocx(pkg: OoxmlPackage): Promise<VisibleItem[]> {
  const out: VisibleItem[] = [];
  for (const part of pkg.paths()) {
    const isHeader = /^word\/(header|footer)\d*\.xml$/.test(part);
    const isBody = part === "word/document.xml";
    if (!isHeader && !isBody) continue;
    const root = await pkg.xml(part);
    if (!root) continue;
    out.push(...scanVml(part, root, isHeader), ...scanDrawings(part, root, isHeader));
    if (isBody && desc(root, NS.w, "background").length > 0) out.push({ code: "pageBackground", part });
  }
  return out;
}
