import { cleanText } from "../clean/hidden-chars";
import { attr, desc, kid, kids, NS, type Element } from "../ooxml/xml";
import { classifyDrawing, classifyVml } from "./drawing";
import {
  ATTRS,
  BOOKMARK_NAME,
  FIELDS,
  KEEP,
  NAMESPACES,
  PREFIX_OF,
  REL_ATTRS,
  SETTINGS_KEEP,
  SUPPORTED_REQUIRES,
  UNWRAP,
} from "./policy";

/** 圖片點處理：由呼叫者按「圖片 N」決定 */
export interface ImageRequest {
  index: number | null; // 「圖片 N」；圖片組合入面嘅第二張起都跟第一張，傳 null 表示唔計數
  oldRelId: string | null; // 原圖嘅 relationship；圖表等冇原圖就係 null
  widthEmu: number;
  heightEmu: number;
}

export interface SerializeCtx {
  /** 外部連結：回傳新 relationship ID；網址唔安全就回傳 null */
  hyperlink(oldRelId: string): string | null;
  /** 頁首頁尾：回傳新 relationship ID */
  headerFooter(oldRelId: string): string | null;
  /** 圖片：回傳新 relationship ID；null = 呢個繪圖唔搬 */
  image(req: ImageRequest): Promise<string | null>;
  /** 新嘅繪圖物件 ID（wp:docPr 等要求全文件唯一） */
  nextDrawingId(): number;
  allowGraphics: boolean;
  isSettings?: boolean;
  /** 被略過嘅元素統計，方便除錯 */
  dropped: Map<string, number>;
  /** 圖片編號，同讀取器一樣按文件次序 */
  imageCounter: { n: number };
}

const TEXT_ELEMENTS = new Set(["w:t", "w:instrText", "m:t", "wp:posOffset", "wp:align", "wp14:pctWidth", "wp14:pctHeight"]);
const FREE_TEXT_ATTRS = new Set(["w:name", "w:val"]); // 樣式名、字型名等：清隱形字元

const escapeText = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
const escapeAttr = (s: string) => escapeText(s).replace(/"/g, "&quot;");

export function qname(el: { namespaceURI: string | null; localName: string | null }): string | null {
  const prefix = el.namespaceURI ? PREFIX_OF[el.namespaceURI] : undefined;
  if (!prefix || !el.localName) return null;
  return `${prefix}:${el.localName}`;
}

function drop(ctx: SerializeCtx, name: string): string {
  ctx.dropped.set(name, (ctx.dropped.get(name) ?? 0) + 1);
  return "";
}

/** 屬性值：長度限制、清隱形字元 */
function safeValue(name: string, value: string): string | null {
  if (value.length > 256) return null;
  return FREE_TEXT_ATTRS.has(name) ? cleanText(value).text : value;
}

function attrsOf(el: Element, name: string, ctx: SerializeCtx, extra: Record<string, string> = {}): string | null {
  const out: Record<string, string> = {};
  for (let i = 0; i < el.attributes.length; i++) {
    const a = el.attributes.item(i)!;
    if (a.name === "xmlns" || a.prefix === "xmlns") continue;
    const prefix = a.namespaceURI ? PREFIX_OF[a.namespaceURI] : "";
    if (prefix === undefined) {
      drop(ctx, `@${a.name}`);
      continue;
    }
    const key = prefix ? `${prefix}:${a.localName}` : (a.localName ?? a.name);
    if (REL_ATTRS.has(key)) continue; // 由呼叫者處理
    if (!ATTRS.has(key)) {
      drop(ctx, `@${key}`);
      continue;
    }
    const v = safeValue(key, a.value);
    if (v !== null) out[key] = v;
  }
  if (name === "w:bookmarkStart" && !BOOKMARK_NAME.test(out["w:name"] ?? "")) return null;
  Object.assign(out, extra);
  return Object.entries(out)
    .map(([k, v]) => ` ${k}="${escapeAttr(v)}"`)
    .join("");
}

async function children(el: Element, ctx: SerializeCtx): Promise<string> {
  let s = "";
  for (const c of kids(el)) s += await element(c, ctx);
  return s;
}

function textOf(el: Element): string {
  let s = "";
  for (let n = el.firstChild; n; n = n.nextSibling) if (n.nodeType === 3 || n.nodeType === 4) s += n.nodeValue ?? "";
  return s;
}

function tag(name: string, attrs: string, inner: string): string {
  return inner ? `<${name}${attrs}>${inner}</${name}>` : `<${name}${attrs}/>`;
}

/** mc:AlternateContent：揀 Word 原生支援嘅 Choice，否則用 Fallback */
async function alternateContent(el: Element, ctx: SerializeCtx): Promise<string> {
  drop(ctx, "mc:AlternateContent(揀一個版本)");
  for (const choice of kids(el, NS.mc, "Choice")) {
    const req = (attr(choice, null, "Requires") ?? "").split(/\s+/).filter(Boolean);
    if (req.length > 0 && req.every((r) => SUPPORTED_REQUIRES.has(r))) return children(choice, ctx);
  }
  const fallback = kid(el, NS.mc, "Fallback");
  return fallback ? children(fallback, ctx) : "";
}

/** DrawingML 繪圖 */
async function drawing(el: Element, ctx: SerializeCtx): Promise<string> {
  const kind = classifyDrawing(el, ctx.allowGraphics);
  if (kind === "skip") return drop(ctx, "w:drawing(skip)");
  const holder = (kid(el, NS.wp, "inline") ?? kid(el, NS.wp, "anchor"))!;
  const extent = kid(holder, NS.wp, "extent");
  const cx = Number(attr(extent, null, "cx") ?? 0);
  const cy = Number(attr(extent, null, "cy") ?? 0);

  if (kind === "other") {
    // 圖表、SmartArt：換成同尺寸嘅佔位圖
    const index = ++ctx.imageCounter.n;
    const rid = await ctx.image({ index, oldRelId: null, widthEmu: cx, heightEmu: cy });
    if (!rid) return "";
    return `<w:drawing>${await pictureHolder(holder, ctx, rid, cx, cy)}</w:drawing>`;
  }

  const blips = desc(holder, NS.a, "blip");
  const relFor = new Map<Element, string | null>();
  if (kind === "picture") {
    // 同讀取器一樣：一個繪圖算一張圖；組合入面嘅其他圖跟第一張嘅決定
    const index = ++ctx.imageCounter.n;
    for (const [i, blip] of blips.entries()) {
      const old = attr(blip, NS.r, "embed") ?? null;
      relFor.set(blip, await ctx.image({ index: i === 0 ? index : null, oldRelId: old, widthEmu: cx, heightEmu: cy }));
    }
    if (blips.length > 0 && [...relFor.values()].every((r) => r === null)) return "";
  } else {
    for (const blip of blips) {
      const old = attr(blip, NS.r, "embed") ?? null;
      relFor.set(blip, await ctx.image({ index: null, oldRelId: old, widthEmu: cx, heightEmu: cy }));
    }
  }
  const inner = await element(holder, { ...ctx, blipRels: relFor } as SerializeCtx & { blipRels: Map<Element, string | null> });
  return inner ? `<w:drawing>${inner}</w:drawing>` : "";
}

/** 用佔位圖重建一個圖片繪圖（保留原本嘅位置同尺寸設定） */
async function pictureHolder(holder: Element, ctx: SerializeCtx, rid: string, cx: number, cy: number): Promise<string> {
  const name = qname(holder)!;
  const attrs = (await attrsOf(holder, name, ctx)) ?? "";
  let inner = "";
  for (const c of kids(holder)) {
    const q = qname(c);
    if (q === "a:graphic") continue;
    inner += await element(c, ctx);
  }
  const id = ctx.nextDrawingId();
  inner +=
    `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>` +
    `<pic:nvPicPr><pic:cNvPr id="${id}" name="Picture ${id}"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
    `</pic:pic></a:graphicData></a:graphic>`;
  return tag(name, attrs, inner);
}

/** 舊式 VML（w:pict、w:object） */
async function vml(el: Element, ctx: SerializeCtx): Promise<string> {
  const kind = classifyVml(el, ctx.allowGraphics);
  if (kind === "skip") return drop(ctx, `${qname(el)}(skip)`);
  const relFor = new Map<Element, string | null>();
  for (const img of desc(el, NS.v, "imagedata")) {
    const old = attr(img, NS.r, "id") ?? null;
    const index = kind === "picture" && relFor.size === 0 ? ++ctx.imageCounter.n : null;
    relFor.set(img, await ctx.image({ index, oldRelId: old, widthEmu: 0, heightEmu: 0 }));
  }
  if (kind === "picture" && [...relFor.values()].every((r) => r === null)) return "";
  // w:object（OLE 物件）只保留預覽圖，所以包成 w:pict
  const inner = await children(el, { ...ctx, blipRels: relFor } as SerializeCtx & { blipRels: Map<Element, string | null> });
  return inner ? `<w:pict>${inner}</w:pict>` : "";
}

export async function element(el: Element, ctx: SerializeCtx): Promise<string> {
  const name = qname(el);
  if (!name) return drop(ctx, `${el.namespaceURI ?? ""}:${el.localName}`);

  // 先處理要特別對待嘅元素
  if (name === "mc:AlternateContent") return alternateContent(el, ctx);
  if (name === "w:sdt") {
    drop(ctx, "w:sdt(拆走外殼)");
    const content = kid(el, NS.w, "sdtContent");
    return content ? children(content, ctx) : "";
  }
  if (UNWRAP.has(name)) {
    drop(ctx, `${name}(拆走外殼)`);
    return children(el, ctx);
  }
  if (name === "w:drawing") return drawing(el, ctx);
  if (name === "w:pict" || name === "w:object") return vml(el, ctx);
  if (name === "w:fldSimple") {
    const field = (attr(el, NS.w, "instr") ?? "").trim().split(/\s+/)[0]?.toUpperCase() ?? "";
    if (!FIELDS.has(field)) return children(el, ctx); // 唔准嘅欄位：只保留顯示結果
  }
  if (name === "w:hyperlink") {
    const old = attr(el, NS.r, "id");
    const anchor = attr(el, NS.w, "anchor");
    const rid = old ? ctx.hyperlink(old) : null;
    if (!rid && !(anchor && BOOKMARK_NAME.test(anchor))) return children(el, ctx);
    const attrs = (await attrsOf(el, name, ctx, rid ? { "r:id": rid } : {})) ?? "";
    return tag(name, attrs, await children(el, ctx));
  }
  if (name === "w:headerReference" || name === "w:footerReference") {
    const old = attr(el, NS.r, "id");
    const rid = old ? ctx.headerFooter(old) : null;
    if (!rid) return drop(ctx, name);
    return tag(name, (await attrsOf(el, name, ctx, { "r:id": rid })) ?? "", "");
  }

  const allowed = ctx.isSettings ? KEEP.has(name) || SETTINGS_KEEP.has(name) || name.startsWith("m:") : KEEP.has(name);
  if (!allowed) return drop(ctx, name);

  // 圖片框架嘅 ID 同名稱重新編號；替代文字唔搬
  if (name === "wp:docPr" || name === "pic:cNvPr" || name === "wps:cNvPr") {
    drop(ctx, `${name}(重新編號、拎走替代文字)`);
    const id = ctx.nextDrawingId();
    return tag(name, ` id="${id}" name="${name === "wp:docPr" ? "Picture" : "Object"} ${id}"`, "");
  }

  const extra: Record<string, string> = {};
  const blipRels = (ctx as SerializeCtx & { blipRels?: Map<Element, string | null> }).blipRels;
  if (name === "a:blip" || name === "v:imagedata") {
    const rid = blipRels?.get(el);
    if (!rid) return drop(ctx, `${name}(no image)`);
    extra[name === "a:blip" ? "r:embed" : "r:id"] = rid;
  }
  if (TEXT_ELEMENTS.has(name) && (name === "w:t" || name === "w:instrText" || name === "m:t")) extra["xml:space"] = "preserve";

  const attrs = await attrsOf(el, name, ctx, extra);
  if (attrs === null) return drop(ctx, `${name}(invalid)`);
  const inner = TEXT_ELEMENTS.has(name) ? escapeText(textOf(el)) : await children(el, ctx);
  return tag(name, attrs, inner);
}

/** 根元素：宣告所有可能用到嘅命名空間 */
export async function serializePart(root: Element, ctx: SerializeCtx): Promise<string> {
  const name = qname(root);
  if (!name || !KEEP.has(name)) throw new Error(`唔支援嘅根元素：${root.nodeName}`);
  const decl = Object.entries(NAMESPACES)
    .filter(([p]) => p !== "xml")
    .map(([p, uri]) => ` xmlns:${p}="${uri}"`)
    .join("");
  const attrs = (await attrsOf(root, name, ctx)) ?? "";
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<${name}${decl}${attrs}>${await children(root, ctx)}</${name}>`;
}
