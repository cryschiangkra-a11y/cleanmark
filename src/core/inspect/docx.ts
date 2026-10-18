import { cleanParagraphSummary } from "../clean/apply-choices";
import { emptySummary, mergeSummary } from "../clean/hidden-chars";
import { scanSuspicious } from "../clean/suspicious";
import { KEEPABLE_IMAGE_TYPES, type CmDocument, type ImageItem, type Report, type SuspiciousSummary } from "../model";
import type { OoxmlPackage } from "../ooxml/package";
import { allImages, allParagraphs, paragraphText } from "../walk";
import { hasC2pa, imageTags } from "./bytes";
import { readOoxmlMetadata } from "./metadata";
import { findPhrases } from "./phrases";
import { findVisibleDocx } from "./visible-docx";

export async function inspectDocx(pkg: OoxmlPackage, doc: CmDocument): Promise<Report> {
  const c2pa: string[] = [];
  for (const path of pkg.paths()) {
    const bytes = await pkg.bytes(path);
    if (bytes && hasC2pa(bytes)) c2pa.push(path);
  }

  const images: ImageItem[] = allImages(doc).map((img) => ({
    index: img.index,
    context: img.context,
    keepable: KEEPABLE_IMAGE_TYPES.has(img.contentType) && !!doc.media[img.mediaPath],
    tags: doc.media[img.mediaPath] ? imageTags(doc.media[img.mediaPath]) : [],
  }));

  let hidden = emptySummary();
  for (const p of allParagraphs(doc)) hidden = mergeSummary(hidden, cleanParagraphSummary(p));

  const suspicious: SuspiciousSummary = { homoglyphs: [], apostrophes: 0, specialSpaces: {} };
  for (const p of allParagraphs(doc)) {
    const scan = scanSuspicious(paragraphText(p)); // 只睇內文文字，方程式唔計
    for (const h of scan.homoglyphs) suspicious.homoglyphs.push({ paragraphId: p.id, word: h.word, fixed: h.fixed });
    suspicious.apostrophes += scan.apostrophes.length;
    for (const [k, n] of Object.entries(scan.specialSpaces)) suspicious.specialSpaces[k] = (suspicious.specialSpaces[k] ?? 0) + n;
  }

  return {
    fileKind: "docx",
    visible: await findVisibleDocx(pkg),
    metadata: await readOoxmlMetadata(pkg),
    c2pa,
    hidden,
    phrases: findPhrases(doc),
    images,
    suspicious,
  };
}
