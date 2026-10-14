import type { MathNode } from "../core/model";

// 預覽用：將方程式轉成一行文字，例如 a/(b+c)、x^2、√(n)。新檔案入面嘅方程式唔受影響

const isNode = (c: MathNode | string): c is MathNode => typeof c !== "string";
const child = (n: MathNode, tag: string) => n.children.filter(isNode).find((c) => c.tag === tag);
const wrap = (s: string) => (Array.from(s).length > 1 ? `(${s})` : s);

/** 屬性（例如 m:chr）喺 m:xxxPr 入面 */
function prop(n: MathNode, prTag: string, name: string): string | undefined {
  return child(child(n, prTag) ?? { tag: "", attrs: {}, children: [] }, name)?.attrs["m:val"];
}

function render(n: MathNode | undefined): string {
  if (!n) return "";
  switch (n.tag) {
    case "m:t":
      return n.children.filter((c): c is string => typeof c === "string").join("");
    case "m:f":
      return `${wrap(render(child(n, "m:num")))}/${wrap(render(child(n, "m:den")))}`;
    case "m:sSup":
      return `${render(child(n, "m:e"))}^${wrap(render(child(n, "m:sup")))}`;
    case "m:sSub":
      return `${render(child(n, "m:e"))}_${wrap(render(child(n, "m:sub")))}`;
    case "m:sSubSup":
      return `${render(child(n, "m:e"))}_${wrap(render(child(n, "m:sub")))}^${wrap(render(child(n, "m:sup")))}`;
    case "m:rad":
      return `√(${render(child(n, "m:e"))})`;
    case "m:d": {
      const beg = prop(n, "m:dPr", "m:begChr") ?? "(";
      const end = prop(n, "m:dPr", "m:endChr") ?? ")";
      const sep = prop(n, "m:dPr", "m:sepChr") ?? ",";
      return beg + n.children.filter(isNode).filter((c) => c.tag === "m:e").map(render).join(sep) + end;
    }
    case "m:nary": {
      const sub = render(child(n, "m:sub"));
      const sup = render(child(n, "m:sup"));
      const chr = prop(n, "m:naryPr", "m:chr") ?? "∫";
      return `${chr}${sub ? `_${wrap(sub)}` : ""}${sup ? `^${wrap(sup)}` : ""} ${render(child(n, "m:e"))}`;
    }
    default:
      // 設定元素（m:fPr、w:rPr 等）唔顯示
      if (/Pr$/.test(n.tag)) return "";
      return n.children.map((c) => (typeof c === "string" ? c : render(c))).join("");
  }
}

export function mathPreviewText(root: MathNode): string {
  return render(root);
}
