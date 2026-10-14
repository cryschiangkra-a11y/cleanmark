import type { Block, CmDocument, ImageSlot, Paragraph, TextInline } from "./model";

/** 逐個區塊行一次（包括表格入面），由上至下 */
export function* walkBlocks(blocks: Block[]): Generator<Block> {
  for (const b of blocks) {
    yield b;
    if (b.type === "table") for (const row of b.rows) for (const cell of row) yield* walkBlocks(cell.blocks);
  }
}

/** 文件入面所有區塊：內文、頁首、頁尾、註腳 */
export function* allBlocks(doc: CmDocument): Generator<Block> {
  yield* walkBlocks(doc.body);
  yield* walkBlocks(doc.header);
  yield* walkBlocks(doc.footer);
  for (const fn of Object.values(doc.footnotes)) yield* walkBlocks(fn);
}

export function allParagraphs(doc: CmDocument): Paragraph[] {
  return [...allBlocks(doc)].filter((b): b is Paragraph => b.type === "paragraph");
}

export function allImages(doc: CmDocument): ImageSlot[] {
  return [...allBlocks(doc)].filter((b): b is ImageSlot => b.type === "image");
}

export function paragraphText(p: Paragraph): string {
  return p.inlines
    .filter((i): i is TextInline => i.type === "text")
    .map((i) => i.text)
    .join("");
}

/** 由頂層開始重建區塊樹：fn 回傳 null 就刪除嗰個區塊 */
export function mapBlocks(blocks: Block[], fn: (b: Block) => Block | null): Block[] {
  const out: Block[] = [];
  for (const b of blocks) {
    const mapped = fn(b);
    if (!mapped) continue;
    if (mapped.type === "table") {
      out.push({
        ...mapped,
        rows: mapped.rows.map((row) => row.map((cell) => ({ ...cell, blocks: mapBlocks(cell.blocks, fn) }))),
      });
    } else out.push(mapped);
  }
  return out;
}
