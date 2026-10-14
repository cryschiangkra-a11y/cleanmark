import { attr, desc, kid, NS, type Element } from "../ooxml/xml";

const NS_PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture";

export type DrawingKind =
  | "skip" // 唔搬：頁首頁尾嘅圖形、擺喺文字後面嘅圖形（水印常見做法）
  | "textbox" // 文字方塊：搬內容
  | "picture" // 圖片：算一張圖
  | "other"; // 圖表、SmartArt 等：算一張圖，用佔位圖

/** 讀取器同重新抄寫器共用，確保「圖片 N」嘅編號兩邊一致 */
export function classifyDrawing(drawing: Element, allowGraphics: boolean): DrawingKind {
  const holder = kid(drawing, NS.wp, "inline") ?? kid(drawing, NS.wp, "anchor");
  if (!allowGraphics || !holder) return "skip";
  const behindText = holder.localName === "anchor" && ["1", "true"].includes(attr(holder, null, "behindDoc") ?? "");
  if (behindText) return "skip";
  if (desc(holder, NS.w, "txbxContent").length > 0) return "textbox";
  if (desc(holder, NS.a, "blip").length > 0 || desc(holder, NS_PIC, "pic").length > 0) return "picture";
  return "other";
}

export type VmlKind = "skip" | "textbox" | "picture";

/** 舊式 VML（w:pict、w:object）。WordArt（v:textpath）通常就係水印，唔搬 */
export function classifyVml(pict: Element, allowGraphics: boolean): VmlKind {
  if (!allowGraphics || desc(pict, NS.v, "textpath").length > 0) return "skip";
  if (desc(pict, NS.w, "txbxContent").length > 0) return "textbox";
  if (desc(pict, NS.v, "imagedata").length > 0) return "picture";
  return "skip";
}
