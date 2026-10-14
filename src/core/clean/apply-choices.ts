import {
  KEEPABLE_IMAGE_TYPES,
  type Block,
  type Choices,
  type CmDocument,
  type HiddenCharSummary,
  type Inline,
  type Paragraph,
  type TextInline,
} from "../model";
import { mapMathText } from "../math";
import { mapBlocks } from "../walk";
import { cleanText, emptySummary, hasRtl, mergeSummary } from "./hidden-chars";

/** 清走段落入面嘅隱形字元；上下文用成個段落嘅文字 */
export function cleanParagraph(p: Paragraph): { paragraph: Paragraph; summary: HiddenCharSummary } {
  const texts = p.inlines.map((i) => (i.type === "text" ? i.text : ""));
  const rtl = hasRtl(texts.join(""));
  let summary = emptySummary();
  const inlines: Inline[] = [];
  p.inlines.forEach((inl, idx) => {
    if (inl.type === "math") {
      const root = mapMathText(inl.root, (text) => {
        const r = cleanText(text, { before: "", after: "", rtl });
        summary = mergeSummary(summary, r.summary);
        return r.text;
      });
      inlines.push({ ...inl, root });
      return;
    }
    if (inl.type !== "text") {
      inlines.push(inl);
      return;
    }
    const before = texts.slice(0, idx).join("");
    const after = texts.slice(idx + 1).join("");
    const r = cleanText(inl.text, { before, after, rtl });
    summary = mergeSummary(summary, r.summary);
    if (r.text) inlines.push({ ...inl, text: r.text } satisfies TextInline);
  });
  return { paragraph: { ...p, inlines }, summary };
}

export function cleanParagraphSummary(p: Paragraph): HiddenCharSummary {
  return cleanParagraph(p).summary;
}

/** 按用家選擇產生最終內容清單：刪段落、清隱形字元、決定圖片保留定佔位 */
export function applyChoices(doc: CmDocument, choices: Choices): CmDocument {
  const remove = new Set(choices.removeParagraphs);
  const keep = new Set(choices.keepImages);
  const keptMedia: Record<string, Uint8Array> = {};

  const fn = (b: Block): Block | null => {
    if (b.type === "paragraph") return remove.has(b.id) ? null : cleanParagraph(b).paragraph;
    if (b.type === "image") {
      const canKeep = keep.has(b.index) && KEEPABLE_IMAGE_TYPES.has(b.contentType) && !!doc.media[b.mediaPath];
      if (canKeep) keptMedia[b.mediaPath] = doc.media[b.mediaPath];
      return { ...b, keep: canKeep };
    }
    return b;
  };

  const footnotes: CmDocument["footnotes"] = {};
  for (const [id, blocks] of Object.entries(doc.footnotes)) footnotes[Number(id)] = mapBlocks(blocks, fn);

  return {
    ...doc,
    body: mapBlocks(doc.body, fn),
    header: mapBlocks(doc.header, fn),
    footer: mapBlocks(doc.footer, fn),
    footnotes,
    media: keptMedia, // 冇揀保留嘅原圖唔會再出現喺記憶體入面嘅內容清單
  };
}
