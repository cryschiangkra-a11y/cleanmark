import { isMathElement, sanitizeMath } from "../../math";
import type { Block, ImageSlot, Inline, Paragraph, RunStyle, Table, TableCell } from "../../model";
import type { OoxmlPackage, Relationship } from "../../ooxml/package";
import { attr, desc, kid, kids, NS, wnum, wval, type Element } from "../../ooxml/xml";
import { classifyDrawing, classifyVml } from "../../sanitize/drawing";
import type { StyleSheet } from "./styles";

const CONTEXT_MAX = 80;
const EMU_PER_PT = 12700;
const PLACEHOLDER_EMU = { w: 2743200, h: 457200 }; // 冇尺寸資料時：3 吋 × 0.5 吋

/** 成份文件共用嘅計數器同收集器 */
export interface ReadState {
  pkg: OoxmlPackage;
  styles: StyleSheet;
  media: Record<string, Uint8Array>;
  imageCount: number;
  paraCount: Record<string, number>;
  lastHeading: string; // 最近一個標題，俾圖片做位置提示
  lastText: string; // 最近一段有字嘅段落
}

/** 每個 part（內文、頁首、註腳）各自嘅設定 */
export interface PartCtx {
  rels: Map<string, Relationship>;
  allowGraphics: boolean; // 頁首頁尾 = false：圖形一律唔搬
  part: string; // part 名，例如 "document"、"header1"
  ids: Map<Element, string>; // 段落 ID（見 ooxml/ids.ts）
}

function contentTypeOf(path: string): ImageSlot["contentType"] {
  const ext = path.slice(path.lastIndexOf(".") + 1).toLowerCase();
  if (ext === "png") return "png";
  if (ext === "jpg" || ext === "jpeg") return "jpg";
  if (ext === "gif") return "gif";
  if (ext === "bmp") return "bmp";
  if (ext === "webp") return "webp";
  if (ext === "tif" || ext === "tiff") return "tiff";
  return "other";
}

async function imageSlot(
  st: ReadState,
  ctx: PartCtx,
  rId: string | undefined,
  widthEmu: number,
  heightEmu: number,
): Promise<ImageSlot> {
  const rel = rId ? ctx.rels.get(rId) : undefined;
  let mediaPath = "";
  if (rel && !rel.external) {
    const bytes = await st.pkg.bytes(rel.target);
    if (bytes) {
      mediaPath = rel.target;
      st.media[mediaPath] = bytes;
    }
  }
  return {
    type: "image",
    index: ++st.imageCount,
    widthEmu: widthEmu || PLACEHOLDER_EMU.w,
    heightEmu: heightEmu || PLACEHOLDER_EMU.h,
    context: st.lastHeading || st.lastText,
    mediaPath,
    contentType: mediaPath ? contentTypeOf(mediaPath) : "other",
    keep: false,
  };
}

/** 段落讀取中途嘅狀態：遇到圖片、分頁、文字方塊就要切開段落 */
class ParaBuilder {
  out: Block[] = [];
  extra: Block[] = []; // 文字方塊內容，擺喺段落後面
  inlines: Inline[] = [];
  split = false; // 已經切開過（遇過圖片或者分頁）
  skipped = false; // 有圖形被略過：如果段落因此變空，就唔留低空段落

  constructor(
    private st: ReadState,
    private id: string,
    private base: Omit<Paragraph, "id" | "inlines">,
  ) {}

  addText(text: string, style: RunStyle, link?: string): void {
    const last = this.inlines.at(-1);
    if (last?.type === "text" && last.link === link && JSON.stringify(last.style) === JSON.stringify(style)) {
      last.text += text;
    } else {
      this.inlines.push({ type: "text", text, style: { ...style }, ...(link ? { link } : {}) });
    }
  }

  flush(force: boolean): void {
    if (this.inlines.length > 0 || (force && !this.split && !this.skipped)) {
      // 同一個 w:p 切開咗都用同一個 ID：刪除嗰陣成段一齊刪
      this.out.push({ ...this.base, id: this.id, inlines: this.inlines });
      this.remember();
    }
    this.inlines = [];
  }

  /** 記低最近嘅標題同文字，俾之後嘅圖片做位置提示 */
  private remember(): void {
    const text = this.inlines
      .map((i) => (i.type === "text" ? i.text : ""))
      .join("")
      .replace(/\s+/g, " ")
      .trim()
      .slice(0, CONTEXT_MAX);
    if (!text) return;
    this.st.lastText = text;
    if (this.base.heading > 0) this.st.lastHeading = text;
  }

  addBlock(b: Block): void {
    this.flush(false);
    this.split = true;
    this.out.push(b);
  }
}

export async function readBlocks(container: Element, st: ReadState, ctx: PartCtx): Promise<Block[]> {
  const out: Block[] = [];
  for (const el of kids(container)) {
    if (el.namespaceURI === NS.w) {
      if (el.localName === "p") out.push(...(await readParagraph(el, st, ctx)));
      else if (el.localName === "tbl") out.push(await readTable(el, st, ctx));
      else if (el.localName === "sdt") {
        const content = kid(el, NS.w, "sdtContent");
        if (content) out.push(...(await readBlocks(content, st, ctx)));
      } else if (el.localName === "customXml") out.push(...(await readBlocks(el, st, ctx)));
    } else if (isMathElement(el)) {
      // 直接喺內文層嘅方程式：包成一個段落
      const n = (st.paraCount[ctx.part] = (st.paraCount[ctx.part] ?? 0) + 1);
      const pb = new ParaBuilder(st, `${ctx.part}:math${n}`, { type: "paragraph", ...st.styles.resolvePara(undefined) });
      pb.inlines.push({ type: "math", display: true, root: sanitizeMath(el) });
      pb.flush(true);
      out.push(...pb.out);
    }
  }
  return out;
}

async function readParagraph(p: Element, st: ReadState, ctx: PartCtx): Promise<Block[]> {
  const pPr = kid(p, NS.w, "pPr");
  const props = st.styles.resolvePara(pPr);
  const pStyleId = st.styles.paraStyleId(pPr);
  const pb = new ParaBuilder(st, ctx.ids.get(p) ?? `${ctx.part}:?`, { type: "paragraph", ...props });
  await readInlineContainer(p, pb, st, ctx, pStyleId, undefined);
  pb.flush(true);
  return [...pb.out, ...pb.extra];
}

const PASS_THROUGH = new Set(["ins", "moveTo", "smartTag", "customXml", "fldSimple", "dir", "bdo"]);

async function readInlineContainer(
  el: Element,
  pb: ParaBuilder,
  st: ReadState,
  ctx: PartCtx,
  pStyleId: string | undefined,
  link: string | undefined,
): Promise<void> {
  for (const c of kids(el)) {
    if (isMathElement(c)) {
      // 方程式係純數學文字，藏唔到圖形水印，原樣保留（只留白名單元素）
      pb.inlines.push({ type: "math", display: c.localName === "oMathPara", root: sanitizeMath(c) });
      continue;
    }
    if (c.namespaceURI !== NS.w) continue;
    const name = c.localName ?? "";
    if (name === "r") await readRun(c, pb, st, ctx, pStyleId, link);
    else if (name === "hyperlink") {
      const rId = attr(c, NS.r, "id");
      const rel = rId ? ctx.rels.get(rId) : undefined;
      await readInlineContainer(c, pb, st, ctx, pStyleId, rel?.external ? rel.target : link);
    } else if (name === "sdt") {
      const content = kid(c, NS.w, "sdtContent");
      if (content) await readInlineContainer(content, pb, st, ctx, pStyleId, link);
    } else if (PASS_THROUGH.has(name)) {
      await readInlineContainer(c, pb, st, ctx, pStyleId, link);
    }
    // del、moveFrom（已刪除）、bookmark、proofErr 等一律略過
  }
}

async function readRun(
  r: Element,
  pb: ParaBuilder,
  st: ReadState,
  ctx: PartCtx,
  pStyleId: string | undefined,
  link: string | undefined,
): Promise<void> {
  const style = st.styles.resolveRun(pStyleId, kid(r, NS.w, "rPr"));
  await readRunChildren(kids(r), pb, st, ctx, style, link);
}

async function readRunChildren(
  children: Element[],
  pb: ParaBuilder,
  st: ReadState,
  ctx: PartCtx,
  style: RunStyle,
  link: string | undefined,
): Promise<void> {
  for (const c of children) {
    if (c.namespaceURI === NS.mc && c.localName === "AlternateContent") {
      // 只讀 Choice，避免同 Fallback 重複
      const choice = kid(c, NS.mc, "Choice") ?? kid(c, NS.mc, "Fallback");
      if (choice) await readRunChildren(kids(choice), pb, st, ctx, style, link);
      continue;
    }
    if (c.namespaceURI !== NS.w) continue;
    switch (c.localName) {
      case "t":
        pb.addText(c.textContent ?? "", style, link);
        break;
      case "tab":
        pb.addText("\t", style, link);
        break;
      case "noBreakHyphen":
        pb.addText("-", style, link);
        break;
      case "br":
        if (attr(c, NS.w, "type") === "page") pb.addBlock({ type: "pageBreak" });
        else pb.inlines.push({ type: "break" });
        break;
      case "cr":
        pb.inlines.push({ type: "break" });
        break;
      case "footnoteReference": {
        const id = wnum(c, "id");
        if (id !== undefined) pb.inlines.push({ type: "footnoteRef", id });
        break;
      }
      case "drawing":
        await readDrawing(c, pb, st, ctx);
        break;
      case "pict":
      case "object":
        await readVml(c, pb, st, ctx);
        break;
      // instrText（功能變數指令）、delText（已刪除文字）等略過
    }
  }
}

async function readDrawing(drawing: Element, pb: ParaBuilder, st: ReadState, ctx: PartCtx): Promise<void> {
  const kind = classifyDrawing(drawing, ctx.allowGraphics);
  // 頁首頁尾嘅圖形、擺喺文字後面嘅圖形 = 水印嘅常見做法，唔搬（檢查器會報告）
  if (kind === "skip") {
    pb.skipped = true;
    return;
  }
  const holder = (kid(drawing, NS.wp, "inline") ?? kid(drawing, NS.wp, "anchor"))!;
  const extent = kid(holder, NS.wp, "extent");
  const cx = Number(attr(extent, null, "cx") ?? 0);
  const cy = Number(attr(extent, null, "cy") ?? 0);
  if (kind === "textbox") {
    pb.extra.push(...(await readBlocks(desc(holder, NS.w, "txbxContent")[0], st, ctx)));
    return;
  }
  const blip = desc(holder, NS.a, "blip")[0];
  // 圖表、SmartArt（kind = "other"）冇原圖可以保留
  pb.addBlock(await imageSlot(st, ctx, kind === "picture" ? attr(blip, NS.r, "embed") : undefined, cx, cy));
}

function vmlSizeEmu(shape: Element | undefined): { w: number; h: number } {
  const style = attr(shape, null, "style") ?? "";
  const num = (key: string) => {
    const m = new RegExp(`(?:^|;)\\s*${key}:\\s*([\\d.]+)pt`).exec(style);
    return m ? Math.round(Number(m[1]) * EMU_PER_PT) : 0;
  };
  return { w: num("width"), h: num("height") };
}

async function readVml(pict: Element, pb: ParaBuilder, st: ReadState, ctx: PartCtx): Promise<void> {
  const kind = classifyVml(pict, ctx.allowGraphics);
  if (kind === "skip") {
    pb.skipped = true;
    return;
  }
  if (kind === "textbox") {
    pb.extra.push(...(await readBlocks(desc(pict, NS.w, "txbxContent")[0], st, ctx)));
    return;
  }
  const img = desc(pict, NS.v, "imagedata")[0];
  const { w, h } = vmlSizeEmu(desc(pict, NS.v, "shape")[0]);
  pb.addBlock(await imageSlot(st, ctx, attr(img, NS.r, "id"), w, h));
}

async function readTable(tbl: Element, st: ReadState, ctx: PartCtx): Promise<Table> {
  const tblPr = kid(tbl, NS.w, "tblPr");
  const grid = kids(kid(tbl, NS.w, "tblGrid") ?? tbl, NS.w, "gridCol").map((g) => wnum(g, "w") ?? 0);
  const rows: TableCell[][] = [];
  const open = new Map<number, TableCell>(); // 欄位置 → 未完嘅垂直合併起點

  const trs = kids(tbl, NS.w, "tr");
  for (const tr of trs) {
    const row: TableCell[] = [];
    let col = 0;
    const tcs = kids(tr).flatMap((c) =>
      c.namespaceURI === NS.w && c.localName === "sdt"
        ? kids(kid(c, NS.w, "sdtContent") ?? c, NS.w, "tc")
        : c.namespaceURI === NS.w && c.localName === "tc"
          ? [c]
          : [],
    );
    for (const tc of tcs) {
      const tcPr = kid(tc, NS.w, "tcPr");
      const span = wnum(kid(tcPr, NS.w, "gridSpan")) ?? 1;
      const vMerge = kid(tcPr, NS.w, "vMerge");
      const vVal = wval(vMerge);
      if (vMerge && vVal !== "restart" && open.has(col)) {
        open.get(col)!.rowSpan++;
        col += span;
        continue;
      }
      const cell: TableCell = { blocks: await readBlocks(tc, st, ctx), colSpan: span, rowSpan: 1 };
      if (vMerge && vVal === "restart") open.set(col, cell);
      else open.delete(col);
      row.push(cell);
      col += span;
    }
    rows.push(row);
  }
  const colCount = Math.max(1, ...rows.map((r) => r.reduce((n, c) => n + c.colSpan, 0)));
  const columnWidths = grid.length > 0 ? grid : Array<number>(colCount).fill(2000);
  return { type: "table", columnWidths, rows, borders: st.styles.tableHasBorders(tblPr) };
}
