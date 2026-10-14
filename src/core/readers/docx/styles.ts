import type { Align, LineSpacing, ListRef, RunStyle } from "../../model";
import { attr, kid, kids, NS, toggle, wnum, wval, type Element } from "../../ooxml/xml";

// 樣式繼承：文件預設 → 段落樣式（由最底嘅 basedOn 開始）→ 字元樣式 → 直接設定

interface StyleDef {
  type: string;
  basedOn?: string;
  name: string;
  pPr?: Element;
  rPr?: Element;
  tblPr?: Element;
}

type PartialRun = Partial<RunStyle>;

export interface ParaProps {
  heading: number;
  align: Align;
  indentLeft: number;
  indentFirstLine: number;
  spacingBefore: number;
  spacingAfter: number;
  lineSpacing: LineSpacing;
  list?: ListRef;
}

const DEFAULT_SIZE_PT = 10; // Word 喺冇任何設定時嘅預設字體大細

function applyRPr(acc: PartialRun, rPr: Element | undefined): void {
  if (!rPr) return;
  const w = (n: string) => kid(rPr, NS.w, n);
  const b = toggle(w("b"));
  if (b !== undefined) acc.bold = b;
  const i = toggle(w("i"));
  if (i !== undefined) acc.italic = i;
  const u = w("u");
  if (u) acc.underline = wval(u) !== "none";
  const strike = toggle(w("strike")) ?? toggle(w("dstrike"));
  if (strike !== undefined) acc.strike = strike;
  const color = wval(w("color"));
  if (color) acc.color = color === "auto" ? undefined : color.toUpperCase();
  const hl = wval(w("highlight"));
  if (hl) acc.highlight = hl === "none" ? undefined : hl;
  const sz = wnum(w("sz"));
  if (sz !== undefined) acc.size = sz / 2;
  const fonts = w("rFonts");
  if (fonts) {
    const ascii = attr(fonts, NS.w, "ascii") ?? attr(fonts, NS.w, "hAnsi");
    const ea = attr(fonts, NS.w, "eastAsia");
    if (ascii) acc.font = ascii;
    if (ea) acc.fontEastAsia = ea;
  }
  const va = wval(w("vertAlign"));
  if (va === "superscript" || va === "subscript") acc.vertAlign = va;
  else if (va === "baseline") acc.vertAlign = undefined;
}

interface PartialPara {
  heading?: number;
  align?: Align;
  indentLeft?: number;
  indentFirstLine?: number;
  spacingBefore?: number;
  spacingAfter?: number;
  lineSpacing?: LineSpacing;
  numId?: string;
  ilvl?: number;
}

const ALIGN: Record<string, Align> = {
  left: "left",
  start: "left",
  center: "center",
  right: "right",
  end: "right",
  both: "justify",
  distribute: "justify",
};

function applyPPr(acc: PartialPara, pPr: Element | undefined): void {
  if (!pPr) return;
  const w = (n: string) => kid(pPr, NS.w, n);
  const jc = wval(w("jc"));
  if (jc && ALIGN[jc]) acc.align = ALIGN[jc];
  const ind = w("ind");
  if (ind) {
    const left = wnum(ind, "left") ?? wnum(ind, "start");
    if (left !== undefined) acc.indentLeft = left;
    const first = wnum(ind, "firstLine");
    const hanging = wnum(ind, "hanging");
    if (hanging !== undefined) acc.indentFirstLine = -hanging;
    else if (first !== undefined) acc.indentFirstLine = first;
  }
  const sp = w("spacing");
  if (sp) {
    const before = wnum(sp, "before");
    const after = wnum(sp, "after");
    const line = wnum(sp, "line");
    if (before !== undefined) acc.spacingBefore = before;
    if (after !== undefined) acc.spacingAfter = after;
    if (line !== undefined) {
      const rule = attr(sp, NS.w, "lineRule");
      acc.lineSpacing = { line, rule: rule === "exact" ? "exact" : rule === "atLeast" ? "atLeast" : "auto" };
    }
  }
  const ol = wnum(w("outlineLvl"));
  if (ol !== undefined) acc.heading = ol >= 0 && ol <= 5 ? ol + 1 : 0;
  const numPr = w("numPr");
  if (numPr) {
    const numId = wval(kid(numPr, NS.w, "numId"));
    const ilvl = wnum(kid(numPr, NS.w, "ilvl"));
    if (numId !== undefined) acc.numId = numId;
    if (ilvl !== undefined) acc.ilvl = ilvl;
  }
}

function headingFromName(name: string): number | undefined {
  const m = /^heading\s*([1-6])$/i.exec(name.trim());
  if (m) return Number(m[1]);
  if (/^title$/i.test(name.trim())) return 1;
  return undefined;
}

export class StyleSheet {
  private styles = new Map<string, StyleDef>();
  private defaultRPr?: Element;
  private defaultPPr?: Element;
  private defaultParaStyle?: string;

  constructor(root: Element | undefined) {
    if (!root) return;
    const defaults = kid(root, NS.w, "docDefaults");
    this.defaultRPr = kid(kid(defaults, NS.w, "rPrDefault"), NS.w, "rPr");
    this.defaultPPr = kid(kid(defaults, NS.w, "pPrDefault"), NS.w, "pPr");
    for (const s of kids(root, NS.w, "style")) {
      const id = attr(s, NS.w, "styleId");
      if (!id) continue;
      const type = attr(s, NS.w, "type") ?? "paragraph";
      this.styles.set(id, {
        type,
        basedOn: wval(kid(s, NS.w, "basedOn")),
        name: wval(kid(s, NS.w, "name")) ?? id,
        pPr: kid(s, NS.w, "pPr"),
        rPr: kid(s, NS.w, "rPr"),
        tblPr: kid(s, NS.w, "tblPr"),
      });
      const isDefault = attr(s, NS.w, "default");
      if (type === "paragraph" && (isDefault === "1" || isDefault === "true")) this.defaultParaStyle = id;
    }
  }

  /** 由最底嘅 basedOn 開始排，避免循環 */
  private chain(id: string | undefined): StyleDef[] {
    const out: StyleDef[] = [];
    const seen = new Set<string>();
    let cur = id;
    while (cur && !seen.has(cur)) {
      seen.add(cur);
      const s = this.styles.get(cur);
      if (!s) break;
      out.unshift(s);
      cur = s.basedOn;
    }
    return out;
  }

  paraStyleId(pPr: Element | undefined): string | undefined {
    return wval(kid(pPr, NS.w, "pStyle")) ?? this.defaultParaStyle;
  }

  resolveRun(pStyleId: string | undefined, rPr: Element | undefined): RunStyle {
    const acc: PartialRun = {};
    applyRPr(acc, this.defaultRPr);
    for (const s of this.chain(pStyleId)) applyRPr(acc, s.rPr);
    const rStyle = wval(kid(rPr, NS.w, "rStyle"));
    for (const s of this.chain(rStyle)) applyRPr(acc, s.rPr);
    applyRPr(acc, rPr);
    return {
      font: acc.font,
      fontEastAsia: acc.fontEastAsia,
      size: acc.size ?? DEFAULT_SIZE_PT,
      bold: acc.bold ?? false,
      italic: acc.italic ?? false,
      underline: acc.underline ?? false,
      strike: acc.strike ?? false,
      color: acc.color,
      highlight: acc.highlight,
      vertAlign: acc.vertAlign,
    };
  }

  resolvePara(pPr: Element | undefined): ParaProps {
    const styleId = this.paraStyleId(pPr);
    const acc: PartialPara = {};
    applyPPr(acc, this.defaultPPr);
    for (const s of this.chain(styleId)) {
      const h = headingFromName(s.name);
      if (h !== undefined) acc.heading = h;
      applyPPr(acc, s.pPr);
    }
    applyPPr(acc, pPr);
    const list = acc.numId && acc.numId !== "0" ? { numId: acc.numId, level: acc.ilvl ?? 0 } : undefined;
    return {
      heading: acc.heading ?? 0,
      align: acc.align ?? "left",
      indentLeft: acc.indentLeft ?? 0,
      indentFirstLine: acc.indentFirstLine ?? 0,
      spacingBefore: acc.spacingBefore ?? 0,
      spacingAfter: acc.spacingAfter ?? 0,
      lineSpacing: acc.lineSpacing ?? { line: 240, rule: "auto" },
      list,
    };
  }

  tableHasBorders(tblPr: Element | undefined): boolean {
    const visible = (borders: Element | undefined) =>
      !!borders && kids(borders, NS.w).some((b) => !["none", "nil"].includes(wval(b) ?? "none"));
    const direct = kid(tblPr, NS.w, "tblBorders");
    if (direct) return visible(direct);
    const styleId = wval(kid(tblPr, NS.w, "tblStyle"));
    for (const s of this.chain(styleId).reverse()) {
      const b = kid(s.tblPr, NS.w, "tblBorders");
      if (b) return visible(b);
    }
    return false;
  }
}
